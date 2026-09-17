"""Complete schema-driven editor for reporting-contract profiles."""

from __future__ import annotations

import asyncio
from collections.abc import Callable, Mapping
from pathlib import Path
from typing import Any

from nicegui import events, ui

from qc_tool.config.editor import (
    DraftPath,
    ProfileDraft,
    Schema,
    default_for_schema,
    profile_to_yaml,
    resolve_schema,
    save_draft,
    schema_at_path,
    schema_leaf_paths,
    yaml_to_profile,
)
from qc_tool.config.lint import LintIssue, lint_profile
from qc_tool.config.profile import DeliverableProfile, default_profile, profile_path
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.ppt.extract import load_deck_snapshot
from qc_tool.progress import CancellationToken, RunCancelled

_profile_to_yaml = profile_to_yaml
_yaml_to_profile = yaml_to_profile

EDITOR_SECTIONS: dict[str, tuple[tuple[str, ...], ...]] = {
    "core": (
        ("name",),
        ("contract_id",),
        ("description",),
        ("tolerance",),
        ("restatement_windows",),
        ("waivers",),
        ("severity",),
        ("materiality_severity",),
        ("excel", "controls"),
        ("crosscheck",),
    ),
    "advanced": (
        ("excel", "ignore_sheets"),
        ("excel", "sheets"),
        ("excel", "comparison_prerequisites"),
        ("excel", "formula_engine"),
        ("excel", "members"),
        ("ppt",),
        ("review_policy",),
        ("input_contract",),
    ),
}

_PROFILE_HELP: dict[DraftPath, str] = {
    ("name",): "The reusable profile name shown in setup and run history.",
    ("contract_id",): "A stable internal identifier. QC Tool manages this value.",
    ("description",): "Plain-language purpose and ownership of this reusable policy.",
    ("tolerance",): (
        "Numeric differences inside these global bounds remain visible as Info; "
        "they are not deleted. Use stricter per-sheet acceptance bands when only "
        "one range needs a tolerance."
    ),
    ("restatement_windows",): (
        "How many recent weekly, monthly, or quarterly periods may be treated as "
        "recent restatements. Older material changes remain Critical."
    ),
    ("waivers",): (
        "Time-limited approved exceptions. Findings stay in evidence as Expected; "
        "use a narrow sheet/location and an expiry date."
    ),
    ("severity",): "Optional finding-class severity overrides for this deliverable.",
    ("materiality_severity",): (
        "Optional severity overrides by materiality tier. Prefer the defaults unless "
        "the review policy has been agreed in advance."
    ),
    ("excel", "controls"): (
        "Reusable workbook assertions: required cells, unique business keys, numeric "
        "bounds, and arithmetic tie-outs."
    ),
    ("crosscheck",): (
        "Confirmed links between PowerPoint figures and workbook source cells for "
        "final-package reconciliation."
    ),
    ("excel", "ignore_sheets"): (
        "Whole sheets omitted from Excel QC. Use only for genuine scratch/helper tabs, "
        "not to silence a noisy deliverable sheet. Scanned sheet names are suggested."
    ),
    ("excel", "sheets"): (
        "Per-sheet reusable overrides. Add a sheet only when it needs ignored or "
        "expected-refresh ranges, cadence/availability rules, chart-window policy, "
        "manual regions, or row identity. Most sheets need no entry."
    ),
    ("excel", "comparison_prerequisites"): (
        "Scenario/filter/parameter cells that must be non-blank and exactly equal "
        "between baseline and current before comparison starts."
    ),
    ("excel", "formula_engine"): (
        "XLSB formula-text adapter. Auto is recommended; an unavailable explicit "
        "engine degrades to formula-presence checks and is disclosed."
    ),
    ("excel", "members"): (
        "Per-workbook rules for multi-workbook packages, keyed by the same stable "
        "member ID used for baseline/current files. Leave empty for one workbook."
    ),
    ("ppt",): (
        "Reusable slide pairing, required-slide, draft-token, chart-window, and "
        "availability policy for PowerPoint."
    ),
    ("review_policy",): (
        "How repeated findings are grouped for review. This changes presentation, "
        "not the underlying evidence."
    ),
    ("input_contract",): (
        "The logical workbook/sheet/region contract saved from Configure & Run. "
        "Normally update it through setup rather than editing it by hand."
    ),
}

_PROFILE_FIELD_HELP = {
    "ignore": "Exclude this configured sheet from Excel QC.",
    "ignore_ranges": "A1 ranges omitted from comparison; reserve for non-evidence cells.",
    "refresh_ranges": "Value changes stay visible but are classified as Expected.",
    "acceptance_bands": "Range-specific numeric differences that stay visible as Info.",
    "regions": "Manual table boundaries when automatic region detection needs correction.",
    "cadence_bands": "Explicit weekly/monthly/quarterly/date axes for adjacent period blocks.",
    "availability_rules": "Where future-period blanks are allowed after a required-through point.",
    "chart_windows": "Force a chart or series to rolling-window or full-history semantics.",
    "row_identity_rules": "Saved business-key row matching for ranked or re-sorted tables.",
    "required_ranges": "Cells or ranges that must be populated in the current workbook.",
    "unique_ranges": "Ranges whose row tuples must be unique.",
    "numeric_bounds": "Minimum/maximum checks for numeric cells or ranges.",
    "tie_outs": "Targets that must reconcile to additive or signed source terms.",
    "slide_pins": "Explicit baseline-title to current-title slide pairings.",
    "match_threshold": "Minimum fuzzy title similarity accepted for automatic slide pairing.",
    "required_slides": "Slide titles that must exist in the final deck.",
    "draft_tokens": "Case-insensitive text tokens that flag unfinished deck content.",
}


def editor_leaf_paths() -> set[tuple[str, ...]]:
    """Return the exact schema leaves assigned to the typed editor sections."""
    paths: set[tuple[str, ...]] = set()
    for roots in EDITOR_SECTIONS.values():
        for root in roots:
            paths.update(schema_leaf_paths(schema_at_path(root), root))
    return paths


class ValidationTask:
    """One cooperative, bounded selected-file profile validation."""

    def __init__(self) -> None:
        self.task: asyncio.Task[None] | None = None
        self.token: CancellationToken | None = None

    @property
    def running(self) -> bool:
        return self.task is not None and not self.task.done()

    def cancel(self) -> bool:
        if self.token is None or not self.running:
            return False
        self.token.cancel()
        return True

    def start(
        self,
        profile: DeliverableProfile,
        files: Mapping[str, Path],
        passwords: Mapping[str, str],
        progress: Callable[[str], None],
        complete: Callable[[list[LintIssue] | Exception], None],
    ) -> None:
        if self.running:
            raise RuntimeError("file-backed validation is already running")
        token = CancellationToken()
        self.token = token

        async def run() -> None:
            try:
                await self._run(
                    profile,
                    dict(files),
                    dict(passwords),
                    token,
                    progress,
                    complete,
                )
            finally:
                if self.token is token:
                    self.token = None
                    self.task = None

        self.task = asyncio.create_task(run())

    @staticmethod
    async def _run(
        profile: DeliverableProfile,
        files: dict[str, Path],
        passwords: dict[str, str],
        token: CancellationToken,
        progress: Callable[[str], None],
        complete: Callable[[list[LintIssue] | Exception], None],
    ) -> None:
        excel_role = next(
            (role for role in ("current_excel", "baseline_excel") if role in files),
            None,
        )
        ppt_role = next(
            (role for role in ("current_ppt", "baseline_ppt") if role in files),
            None,
        )
        try:
            workbook = None
            if excel_role is not None:
                progress(f"Loading {excel_role.replace('_', ' ')} locally...")
                workbook = await asyncio.to_thread(
                    load_workbook_snapshot,
                    files[excel_role],
                    password=passwords.get(excel_role),
                    allow_large_workbook=False,
                    cancellation_token=token,
                )
            deck = None
            if ppt_role is not None:
                progress(f"Loading {ppt_role.replace('_', ' ')} locally...")
                deck = await asyncio.to_thread(
                    load_deck_snapshot,
                    files[ppt_role],
                    password=passwords.get(ppt_role),
                    cancellation_token=token,
                )
            progress("Checking contract references...")
            complete(lint_profile(profile, workbook=workbook, deck=deck))
        except Exception as exc:
            complete(exc)


class ProfileEditorSession:
    """Load one named source into the lossless draft boundary."""

    def __init__(self, profiles_dir: Path, selected_name: str) -> None:
        self.profiles_dir = profiles_dir
        self.draft = self._load(selected_name)

    def _load(self, name: str) -> ProfileDraft:
        if name == "default":
            return ProfileDraft.from_profile(default_profile())
        return ProfileDraft.from_path(profile_path(self.profiles_dir, name))

    def load(self, name: str) -> None:
        self.draft = self._load(name)


def _label(name: str) -> str:
    return name.replace("_", " ").title()


def _item_summary(value: object, index: int) -> str:
    if isinstance(value, dict):
        for key in ("name", "sheet", "slide", "label", "range", "target"):
            if value.get(key):
                return f"{index + 1}. {value[key]}"
    return f"Item {index + 1}"


class ProfileEditorController:
    """NiceGUI controller over one complete profile draft."""

    def __init__(
        self,
        container: ui.element,
        profiles_dir: Path,
        editing: ui.select,
        *,
        on_close: Callable[[], None],
        selected_files: Callable[[], Mapping[str, Path]],
        selected_passwords: Callable[[], Mapping[str, str]],
        on_saved: Callable[[str, str], None],
        mapping_key_suggestions: Callable[[DraftPath], tuple[str, ...]] | None = None,
    ) -> None:
        self.container = container
        self.profiles_dir = profiles_dir
        self.editing = editing
        self.on_close = on_close
        self.selected_files = selected_files
        self.selected_passwords = selected_passwords
        self.on_saved = on_saved
        self.mapping_key_suggestions = mapping_key_suggestions or (lambda _path: ())
        self.session = ProfileEditorSession(
            profiles_dir, str(editing.value or "default")
        )
        self.validation = ValidationTask()
        self.yaml_pending = False
        self.syncing_yaml = False
        self.setting_selection = False
        self.disposed = False
        self._build()
        self.editing.on_value_change(self._selection_changed)
        self._render_forms()
        self._sync_yaml()
        self._update_state()

    @property
    def dirty(self) -> bool:
        return self.session.draft.dirty or self.yaml_pending

    @property
    def immutable(self) -> bool:
        return self.session.draft.selected_name == "default"

    def _build(self) -> None:
        with self.container:
            ui.label(
                "This is the advanced editor for reusable policy. Most one-off "
                "runs only need the setup tabs behind this dialog. Use these "
                "fields when a tolerance, control, sheet rule, waiver, or mapping "
                "should apply again; the form and YAML edit the same profile."
            ).classes("note")
            with ui.tabs().props("dense no-caps align=left") as tabs:
                ui.tab("core", label="Core contract")
                ui.tab("advanced", label="Advanced Excel/PPT")
                ui.tab("yaml", label="Advanced YAML")
                ui.tab("validate", label="Validate")
            with ui.tab_panels(
                tabs, value="core", animated=False, keep_alive=True
            ).classes("profile-editor-panels w-full"):
                with ui.tab_panel("core"):
                    self.core_box = ui.column().classes("w-full gap-1")
                with ui.tab_panel("advanced"):
                    self.advanced_box = ui.column().classes("w-full gap-1")
                with ui.tab_panel("yaml"):
                    self.yaml_editor = (
                        ui.textarea("Canonical profile YAML")
                        .classes("w-full profile-yaml")
                        .props("outlined rows=24")
                    )
                    self.yaml_editor.on_value_change(self._yaml_changed)
                    with ui.row().classes("items-center gap-2"):
                        self.apply_yaml_button = ui.button(
                            "Apply YAML", on_click=self._apply_yaml
                        ).classes("ghostbtn").props("flat no-caps")
                        self.reset_yaml_button = ui.button(
                            "Reset YAML", on_click=self._reset_yaml
                        ).classes("ghostbtn").props("flat no-caps")
                with ui.tab_panel("validate"):
                    ui.label(
                        "Validation is local and read-only. It uses the selected "
                        "current file when available, otherwise the baseline, and "
                        "keeps workbook workload safeguards active."
                    ).classes("note")
                    ui.label(
                        "XLSB profile-reference validation remains bounded. Formula "
                        "text adapter availability and any degradation are disclosed "
                        "by the subsequent QC run, not inferred by this lint."
                    ).classes("notecard")
                    with ui.row().classes("items-center gap-2"):
                        ui.button(
                            "Validate selected files",
                            on_click=self._validate_selected_files,
                        ).classes("runbtn").props("no-caps")
                        ui.button(
                            "Cancel validation", on_click=self._cancel_validation
                        ).classes("ghostbtn").props("flat no-caps")
                    self.validation_status = ui.label().classes("note preline")
            self.dirty_status = ui.label().classes("note")
            with ui.row().classes("items-center gap-2 profile-editor-actions"):
                self.save_button = ui.button(
                    "Save profile", on_click=self._save
                ).classes("runbtn").props("no-caps")
                ui.button("Close", on_click=self.request_close).props("flat no-caps")

    def _update_state(self) -> None:
        if self.yaml_pending:
            self.dirty_status.set_text("YAML edits are not applied to the draft")
        elif self.session.draft.dirty:
            self.dirty_status.set_text("Unsaved profile changes")
        else:
            self.dirty_status.set_text("Profile matches its saved source")
        if self.immutable:
            self.save_button.disable()
            self.apply_yaml_button.disable()
            self.reset_yaml_button.disable()
            self.yaml_editor.disable()
        else:
            self.save_button.enable()
            self.apply_yaml_button.enable()
            self.reset_yaml_button.enable()
            self.yaml_editor.enable()

    def _sync_yaml(self) -> None:
        self.syncing_yaml = True
        self.yaml_editor.value = self.session.draft.yaml()
        self.yaml_editor.update()
        self.syncing_yaml = False
        self.yaml_pending = False

    def _yaml_changed(self, event: events.ValueChangeEventArguments) -> None:
        if self.syncing_yaml or self.immutable:
            return
        self.yaml_pending = str(event.value or "") != self.session.draft.yaml()
        self._update_state()

    def _apply_yaml(self) -> None:
        try:
            self.session.draft.apply_yaml(str(self.yaml_editor.value or ""))
        except Exception as exc:
            ui.notify(f"Invalid YAML: {exc}", type="negative")
            return
        self.yaml_pending = False
        self._render_forms()
        self._sync_yaml()
        self._update_state()
        ui.notify("YAML applied to the typed form")

    def _reset_yaml(self) -> None:
        self._sync_yaml()
        self._update_state()

    def _allow_form_change(self) -> bool:
        if self.immutable:
            return False
        if self.yaml_pending:
            ui.notify("Apply or reset the YAML edits before changing the form", type="warning")
            self._render_forms()
            return False
        return True

    def _set_value(self, path: DraftPath, value: object) -> None:
        if not self._allow_form_change():
            return
        self.session.draft.set(path, value)
        self._sync_yaml()
        self._update_state()

    def _text_changed(
        self,
        event: events.ValueChangeEventArguments,
        path: DraftPath,
        nullable: bool,
    ) -> None:
        text = str(event.value or "")
        self._set_value(path, None if nullable and not text else text)

    def _number_changed(
        self,
        event: events.ValueChangeEventArguments,
        path: DraftPath,
        integer: bool,
    ) -> None:
        value = event.value
        self._set_value(
            path,
            None if value in {None, ""} else int(value) if integer else float(value),
        )

    def _render_forms(self) -> None:
        self.core_box.clear()
        with self.core_box:
            if self.immutable:
                ui.label(
                    "The built-in default profile is immutable. Create a named "
                    "profile to edit this contract."
                ).classes("notecard")
            for path in EDITOR_SECTIONS["core"]:
                self._render_node(schema_at_path(path), path, path[-1], 0)
        self.advanced_box.clear()
        with self.advanced_box:
            for path in EDITOR_SECTIONS["advanced"]:
                self._render_node(schema_at_path(path), path, path[-1], 0)

    def _render_node(
        self,
        schema: Schema,
        path: DraftPath,
        name: str,
        depth: int,
    ) -> None:
        resolved, _ = resolve_schema(schema)
        kind = resolved.get("type")
        if kind == "object" and resolved.get("properties"):
            with ui.expansion(_label(name), value=depth == 0).classes(
                "w-full profile-section"
            ):
                self._render_help(path, name)
                self._render_object(resolved, path, depth + 1)
        elif kind == "object":
            self._render_mapping(resolved, path, name, depth)
        elif kind == "array":
            self._render_array(resolved, path, name, depth)
        else:
            self._render_scalar(schema, path, name)

    def _render_help(self, path: DraftPath, name: str) -> None:
        help_text = _PROFILE_HELP.get(path) or _PROFILE_FIELD_HELP.get(name)
        if help_text:
            ui.label(help_text).classes("profile-help note")

    def _render_object(self, schema: Schema, path: DraftPath, depth: int) -> None:
        for name, child_schema in schema.get("properties", {}).items():
            self._render_node(child_schema, (*path, name), name, depth)

    def _lock(self, control: Any) -> None:
        if self.immutable:
            control.disable()

    def _draft_value(self, path: DraftPath, schema: Schema) -> object:
        """Read a field, projecting schema defaults omitted by serialization."""
        try:
            return self.session.draft.get(path)
        except KeyError:
            return default_for_schema(schema)

    def _render_scalar(
        self,
        schema: Schema,
        path: DraftPath,
        name: str,
        *,
        classes: str = "w-full",
    ) -> None:
        resolved, nullable = resolve_schema(schema)
        value = self._draft_value(path, schema)
        enum_values = [str(item) for item in resolved.get("enum", [])]
        if enum_values:
            options = {item: item.replace("_", " ") for item in enum_values}
            if nullable:
                options = {"": "Not set", **options}
            control = ui.select(
                options,
                value="" if value is None else str(value),
                label=_label(name),
                on_change=lambda event: self._set_value(
                    path, None if nullable and event.value == "" else event.value
                ),
            ).classes(classes).props("outlined dense options-dense")
            self._lock(control)
            self._render_help(path, name)
            return
        kind = resolved.get("type")
        if kind == "boolean":
            control = ui.checkbox(
                _label(name),
                value=bool(value),
                on_change=lambda event: self._set_value(path, bool(event.value)),
            ).classes(classes)
        elif kind in {"integer", "number"}:
            props = "outlined dense"
            minimum = resolved.get("minimum")
            maximum = resolved.get("maximum")
            if minimum is not None:
                props += f" min={minimum}"
            if maximum is not None:
                props += f" max={maximum}"
            numeric_value = float(value) if isinstance(value, (int, float)) else None
            control = ui.number(
                _label(name),
                value=numeric_value,
                on_change=lambda event: self._number_changed(
                    event, path, kind == "integer"
                ),
            ).classes(classes).props(props)
        else:
            factory = ui.textarea if name == "description" else ui.input
            control = factory(
                _label(name),
                value="" if value is None else str(value),
                on_change=lambda event: self._text_changed(event, path, nullable),
            ).classes(classes).props("outlined dense")
            if resolved.get("format") == "date":
                control.props("type=date")
            if name == "contract_id":
                control.props("readonly")
        self._lock(control)
        self._render_help(path, name)

    def _icon_button(
        self,
        icon: str,
        tooltip: str,
        handler: Callable[[], None],
        *,
        disabled: bool = False,
    ) -> None:
        button = (
            ui.button(icon=icon, on_click=handler)
            .classes("profile-icon-button")
            .props(f'flat round dense aria-label="{tooltip}"')
            .tooltip(tooltip)
        )
        if disabled or self.immutable:
            button.disable()

    def _move_handler(
        self,
        path: DraftPath,
        index: int,
        offset: int,
    ) -> Callable[[], None]:
        def move() -> None:
            if not self._allow_form_change():
                return
            self.session.draft.move(path, index, offset)
            self._sync_yaml()
            self._update_state()
            self._render_forms()

        return move

    def _remove_handler(self, path: DraftPath, index: int) -> Callable[[], None]:
        def remove() -> None:
            if not self._allow_form_change():
                return
            self.session.draft.remove(path, index)
            self._sync_yaml()
            self._update_state()
            self._render_forms()

        return remove

    def _append_handler(
        self,
        path: DraftPath,
        item_schema: Schema,
    ) -> Callable[[], None]:
        def append() -> None:
            if not self._allow_form_change():
                return
            self.session.draft.append(path, item_schema)
            self._sync_yaml()
            self._update_state()
            self._render_forms()

        return append

    def _append_value_handler(
        self,
        path: DraftPath,
        item_schema: Schema,
        control: Any,
    ) -> Callable[[], None]:
        def append() -> None:
            if not self._allow_form_change():
                return
            value = str(control.value or "").strip()
            if not value:
                ui.notify("Choose or enter a value first", type="warning")
                return
            current = self.session.draft.get(path)
            if isinstance(current, list) and value in current:
                ui.notify(f"{value!r} is already configured", type="warning")
                return
            index = len(current) if isinstance(current, list) else 0
            self.session.draft.append(path, item_schema)
            self.session.draft.set((*path, index), value)
            self._sync_yaml()
            self._update_state()
            self._render_forms()

        return append

    def _render_array(
        self,
        schema: Schema,
        path: DraftPath,
        name: str,
        depth: int,
    ) -> None:
        values = self._draft_value(path, schema)
        if not isinstance(values, list):
            return
        item_schema = schema.get("items", {"type": "string"})
        item_resolved, _ = resolve_schema(item_schema)
        with ui.expansion(f"{_label(name)} ({len(values)})", value=depth == 0).classes(
            "w-full profile-section"
        ):
            self._render_help(path, name)
            for index, value in enumerate(values):
                if item_resolved.get("type") == "object":
                    with ui.expansion(_item_summary(value, index)).classes("w-full"):
                        with ui.row().classes("items-center gap-1"):
                            self._icon_button(
                                "arrow_upward",
                                "Move item up",
                                self._move_handler(path, index, -1),
                                disabled=index == 0,
                            )
                            self._icon_button(
                                "arrow_downward",
                                "Move item down",
                                self._move_handler(path, index, 1),
                                disabled=index == len(values) - 1,
                            )
                            self._icon_button(
                                "delete",
                                "Remove item",
                                self._remove_handler(path, index),
                            )
                        self._render_object(item_resolved, (*path, index), depth + 1)
                else:
                    with ui.row().classes("items-center gap-1 w-full no-wrap"):
                        self._render_scalar(
                            item_schema,
                            (*path, index),
                            f"{_label(name)} {index + 1}",
                            classes="flex-1",
                        )
                        self._icon_button(
                            "arrow_upward",
                            "Move item up",
                            self._move_handler(path, index, -1),
                            disabled=index == 0,
                        )
                        self._icon_button(
                            "arrow_downward",
                            "Move item down",
                            self._move_handler(path, index, 1),
                            disabled=index == len(values) - 1,
                        )
                        self._icon_button(
                            "delete",
                            "Remove item",
                            self._remove_handler(path, index),
                        )
            suggestions = tuple(
                item
                for item in self.mapping_key_suggestions(path)
                if item and item not in values
            )
            if item_resolved.get("type") == "string" and suggestions:
                with ui.row().classes("items-end gap-2 w-full"):
                    value_control = ui.select(
                        list(dict.fromkeys(suggestions)),
                        label=f"Add {_label(name)}",
                        with_input=True,
                        new_value_mode="add-unique",
                    ).classes("flex-1").props("outlined dense")
                    add_button = ui.button(
                        "Add",
                        icon="add",
                        on_click=self._append_value_handler(
                            path, item_schema, value_control
                        ),
                    ).classes("ghostbtn").props("flat no-caps dense")
            else:
                add_button = ui.button(
                    f"Add {_label(name)} item",
                    icon="add",
                    on_click=self._append_handler(path, item_schema),
                ).classes("ghostbtn").props("flat no-caps dense")
            self._lock(add_button)

    def _rename_mapping_handler(
        self,
        path: DraftPath,
        old_key: str,
        control: Any,
    ) -> Callable[[], None]:
        def rename() -> None:
            if not self._allow_form_change():
                return
            try:
                self.session.draft.rename_mapping_key(
                    path, old_key, str(control.value or "")
                )
            except ValueError as exc:
                ui.notify(str(exc), type="warning")
                return
            self._sync_yaml()
            self._update_state()
            self._render_forms()

        return rename

    def _remove_mapping_handler(
        self,
        path: DraftPath,
        key: str,
    ) -> Callable[[], None]:
        def remove() -> None:
            if not self._allow_form_change():
                return
            self.session.draft.remove_mapping(path, key)
            self._sync_yaml()
            self._update_state()
            self._render_forms()

        return remove

    def _add_mapping_handler(
        self,
        path: DraftPath,
        key_control: Any,
        value_schema: Schema,
    ) -> Callable[[], None]:
        def add() -> None:
            if not self._allow_form_change():
                return
            try:
                self.session.draft.set_mapping(
                    path,
                    str(key_control.value or ""),
                    default_for_schema(value_schema),
                )
            except ValueError as exc:
                ui.notify(str(exc), type="warning")
                return
            self._sync_yaml()
            self._update_state()
            self._render_forms()

        return add

    def _render_mapping(
        self,
        schema: Schema,
        path: DraftPath,
        name: str,
        depth: int,
    ) -> None:
        mapping = self._draft_value(path, schema)
        if not isinstance(mapping, dict):
            return
        value_schema = schema.get("additionalProperties", {"type": "string"})
        value_resolved, _ = resolve_schema(value_schema)
        key_schema, _ = resolve_schema(schema.get("propertyNames", {"type": "string"}))
        enum_keys = [str(value) for value in key_schema.get("enum", [])]
        with ui.expansion(f"{_label(name)} ({len(mapping)})", value=depth == 0).classes(
            "w-full profile-section"
        ):
            self._render_help(path, name)
            for raw_key in list(mapping):
                key = str(raw_key)
                with ui.expansion(key.replace("_", " ")).classes("w-full"):
                    with ui.row().classes("items-center gap-1 w-full"):
                        if not enum_keys:
                            key_input = ui.input("Key", value=key).classes(
                                "flex-1"
                            ).props("outlined dense")
                            self._lock(key_input)
                            self._icon_button(
                                "drive_file_rename_outline",
                                "Rename entry",
                                self._rename_mapping_handler(path, key, key_input),
                            )
                        else:
                            ui.label(key.replace("_", " ")).classes("flex-1")
                        self._icon_button(
                            "delete",
                            "Remove entry",
                            self._remove_mapping_handler(path, key),
                        )
                    if value_resolved.get("type") == "object":
                        self._render_object(value_resolved, (*path, key), depth + 1)
                    else:
                        self._render_scalar(value_schema, (*path, key), "Value")
            available_keys = [key for key in enum_keys if key not in mapping]
            if enum_keys and not available_keys:
                ui.label("All available keys are configured.").classes("note")
                return
            suggested_keys = tuple(
                key
                for key in self.mapping_key_suggestions(path)
                if key and key not in mapping
            )
            key_label = {
                "members": "New workbook member ID",
                "sheets": "New sheet",
            }.get(name, "New key")
            key_control = (
                ui.select(available_keys, label=key_label)
                if enum_keys
                else ui.select(
                    list(dict.fromkeys(suggested_keys)),
                    label=key_label,
                    with_input=True,
                    new_value_mode="add-unique",
                )
                if suggested_keys
                else ui.input(key_label)
            )
            key_control.classes("w-full").props("outlined dense")
            self._lock(key_control)
            add_button = ui.button(
                "Add entry",
                icon="add",
                on_click=self._add_mapping_handler(path, key_control, value_schema),
            ).classes("ghostbtn").props("flat no-caps dense")
            self._lock(add_button)

    def _save(self) -> None:
        if self.yaml_pending:
            ui.notify("Apply or reset the YAML edits before saving", type="warning")
            return
        old_name = self.session.draft.selected_name
        try:
            profile = save_draft(self.session.draft, self.profiles_dir)
        except Exception as exc:
            ui.notify(str(exc), type="negative")
            return
        self.on_saved(old_name, profile.name)
        self._set_selection(profile.name)
        self._sync_yaml()
        self._update_state()
        ui.notify("Profile saved")

    def _validate_selected_files(self) -> None:
        try:
            profile = self.session.draft.profile()
        except Exception as exc:
            ui.notify(f"Invalid profile: {exc}", type="negative")
            return
        files = dict(self.selected_files())
        if not files:
            ui.notify("Select files on Compare before validation", type="warning")
            return

        def progress(message: str) -> None:
            if not self.disposed:
                self.validation_status.set_text(message)

        def complete(result: list[LintIssue] | Exception) -> None:
            if self.disposed:
                return
            if isinstance(result, RunCancelled):
                self.validation_status.set_text("Validation cancelled")
                return
            if isinstance(result, Exception):
                self.validation_status.set_text(f"Validation failed: {result}")
                return
            errors = [issue for issue in result if issue.level == "error"]
            warnings = [issue for issue in result if issue.level == "warning"]
            lines = [
                f"{issue.level}: {issue.where}: {issue.message}"
                for issue in result[:8]
            ]
            summary = f"{len(errors)} errors, {len(warnings)} warnings"
            self.validation_status.set_text("\n".join([summary, *lines]))

        try:
            self.validation.start(
                profile,
                files,
                dict(self.selected_passwords()),
                progress,
                complete,
            )
        except RuntimeError as exc:
            ui.notify(str(exc), type="warning")

    def _cancel_validation(self) -> None:
        if self.validation.cancel():
            self.validation_status.set_text("Cancellation requested...")
        else:
            ui.notify("No file-backed validation is running")

    def _set_selection(self, name: str) -> None:
        self.setting_selection = True
        self.editing.value = name
        self.editing.update()
        self.setting_selection = False

    def _selection_changed(self, event: events.ValueChangeEventArguments) -> None:
        if self.setting_selection:
            return
        self.request_load(str(event.value or "default"))

    def request_load(self, name: str) -> None:
        if name == self.session.draft.selected_name:
            return
        if not self.dirty:
            self._load(name)
            return
        self._set_selection(self.session.draft.selected_name)
        with ui.dialog() as dialog, ui.card().classes("w-[30rem] max-w-full"):
            ui.label("Discard profile changes?").classes("runhead")
            ui.label(
                "The current form or YAML has unsaved changes."
            ).classes("note")

            def discard() -> None:
                dialog.close()
                self._load(name)

            with ui.row().classes("items-center gap-2"):
                ui.button("Discard and switch", on_click=discard).classes(
                    "runbtn"
                ).props("no-caps")
                ui.button("Keep editing", on_click=dialog.close).props("flat no-caps")
        dialog.open()

    def _load(self, name: str) -> None:
        self.validation.cancel()
        try:
            self.session.load(name)
        except Exception as exc:
            ui.notify(f"Failed to load profile: {exc}", type="negative")
            self._set_selection(self.session.draft.selected_name)
            return
        self._set_selection(name)
        self.yaml_pending = False
        self._render_forms()
        self._sync_yaml()
        self._update_state()

    def request_close(self) -> None:
        if not self.dirty:
            self._finish_close()
            return
        with ui.dialog() as dialog, ui.card().classes("w-[30rem] max-w-full"):
            ui.label("Discard profile changes?").classes("runhead")
            ui.label(
                "Closing now will discard the current form or YAML changes."
            ).classes("note")

            def discard() -> None:
                dialog.close()
                self._finish_close()

            with ui.row().classes("items-center gap-2"):
                ui.button("Discard and close", on_click=discard).classes(
                    "runbtn"
                ).props("no-caps")
                ui.button("Keep editing", on_click=dialog.close).props("flat no-caps")
        dialog.open()

    def _finish_close(self) -> None:
        self.dispose()
        self.on_close()

    def dispose(self) -> None:
        self.disposed = True
        self.validation.cancel()


def open_profile_editor(
    container: ui.element,
    profiles_dir: Path,
    editing: ui.select,
    *,
    on_close: Callable[[], None],
    selected_files: Callable[[], Mapping[str, Path]],
    selected_passwords: Callable[[], Mapping[str, str]],
    on_saved: Callable[[str, str], None],
    mapping_key_suggestions: Callable[[DraftPath], tuple[str, ...]] | None = None,
) -> ProfileEditorController:
    """Attach and return the complete typed profile editor controller."""
    return ProfileEditorController(
        container,
        profiles_dir,
        editing,
        on_close=on_close,
        selected_files=selected_files,
        selected_passwords=selected_passwords,
        on_saved=on_saved,
        mapping_key_suggestions=mapping_key_suggestions,
    )
