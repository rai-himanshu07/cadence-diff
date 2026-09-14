"""Full-page, mode-aware configuration workspace (plan-20260913, Step 7).

The mandatory intermediate step between file intake (still handled by the
existing "/" page's uploads/mode/profile controls) and actual QC
submission. Auto-starts a bounded setup-analysis scan, lets the analyst
review structure and pick a profile only after the scan completes, and
performs the run/save actions through the same `RunRequest` shape the main
page's own retry paths use (`qc_tool.ui.app.build_run_request`).

Scope boundary (disclosed, matches the plan's own step ownership): this
module builds the SHELL -- auto-start/cancel/reconnect, post-analysis
profile selection with a field diff, a current/baseline preview toggle,
per-region mode/anchor/identity-column decisions, warning acknowledgement,
and the five final actions. It does not yet implement every region/column
semantic (blank-key threshold nuance, period bands, selector
prerequisites, PowerPoint/final-package slide review) -- those are Step
8/9's job, extending this same shell and its underlying
`qc_tool.ui.config_review` builder.
"""

from __future__ import annotations

import asyncio
import dataclasses
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

from nicegui import events, ui
from openpyxl.utils import get_column_letter

from qc_tool.config.profile import (
    DeliverableProfile,
    list_profiles,
    load_profile_by_name,
    profile_path,
    profile_sha256,
    save_profile,
)
from qc_tool.coverage import QCRunMode
from qc_tool.history.config_session import ConfigSessionStore
from qc_tool.history.store import sha256_file
from qc_tool.package import PackageManifest
from qc_tool.runqueue import QueueBusyError, RunQueueManager, run_exclusive
from qc_tool.setup.models import MemberSetupProfile, SetupAnalysisResult
from qc_tool.setup.preview_worker import (
    PreviewWindowRequest,
    SetupScanRequest,
    run_preview_window_worker,
    run_setup_scan_worker,
)
from qc_tool.ui.config_review import (
    ConfigWorkspaceState,
    DeckReview,
    MemberReview,
    SelectorDecision,
    add_selector,
    apply_anchor_click,
    apply_manual_range,
    apply_region_transform,
    build_input_contract,
    build_resolved_configuration,
    compute_required_slide_warnings,
    compute_sheet_pairing_warnings,
    compute_slide_pairing_warnings,
    compute_warnings,
    confirm_all_regions,
    deck_review_from_titles,
    diff_profile_against_scan,
    effective_slide_pairing,
    is_clean_profile_diff,
    member_review_from_scan,
    parse_a1_cell,
    remove_selector,
    resolve_slide_anchors,
    set_expected_refresh_columns,
    set_identity_columns,
    set_ignore_columns,
    set_ordinal_columns,
    set_sheet_rename,
    set_slide_included,
    set_slide_rename,
    sheet_pairing,
    unresolved_blockers,
    update_region_decision,
)
from qc_tool.ui.theme import page_frame, section, status_chip

#: Excel role prefixes this workspace's setup scan understands.
_EXCEL_BASELINE_PREFIX = "baseline_excel"
_EXCEL_CURRENT_PREFIX = "current_excel"
#: PowerPoint role prefixes (Step 9's own slide-review scope).
_PPT_BASELINE_PREFIX = "baseline_ppt"
_PPT_CURRENT_PREFIX = "current_ppt"

#: One preview "window" fetched at a time -- comfortably inside
#: `preview_worker.MAX_PREVIEW_ROWS`/`MAX_PREVIEW_COLS`, small enough to stay
#: a lightweight DOM table.
_PREVIEW_WINDOW_ROWS = 20
_PREVIEW_WINDOW_COLS = 12


JobStatus = Literal["idle", "running", "done", "cancelled", "failed"]


@dataclass(slots=True)
class SetupJob:
    """One in-memory setup-scan job, keyed by config session id so a
    browser refresh/reconnect (a fresh NiceGUI page load) observes the
    SAME job's live status instead of restarting it (Step 7's
    "auto-start/cancel/reconnect" criterion).

    Tracks one status per member so a multi-member package's scans can run
    (sequentially, sharing the one exclusive slot) without one member's
    outcome overwriting another's.
    """

    member_status: dict[str, JobStatus] = field(default_factory=dict)
    member_results: dict[str, MemberSetupProfile] = field(default_factory=dict)
    member_disclosures: dict[str, str] = field(default_factory=dict)
    cancel_event: threading.Event = field(default_factory=threading.Event)
    missing_credential_roles: dict[str, tuple[str, ...]] = field(default_factory=dict)
    started: bool = False

    @property
    def overall_status(self) -> JobStatus:
        statuses = set(self.member_status.values())
        if not statuses:
            return "idle"
        if "running" in statuses:
            return "running"
        if "failed" in statuses:
            return "failed"
        if "cancelled" in statuses:
            return "cancelled"
        return "done"

    @property
    def result(self) -> SetupAnalysisResult | None:
        if not self.member_results:
            return None
        return SetupAnalysisResult(members=dict(self.member_results))


class SetupJobRegistry:
    """Per-work-dir in-memory registry of setup-scan jobs."""

    def __init__(self) -> None:
        self._jobs: dict[str, SetupJob] = {}
        self._lock = threading.Lock()

    def get_or_create(self, session_key: str) -> SetupJob:
        with self._lock:
            job = self._jobs.get(session_key)
            if job is None:
                job = SetupJob()
                self._jobs[session_key] = job
            return job

    def reset(self, session_key: str) -> SetupJob:
        with self._lock:
            job = SetupJob()
            self._jobs[session_key] = job
            return job


_REGISTRIES: dict[Path, SetupJobRegistry] = {}
_REGISTRIES_LOCK = threading.Lock()


def get_setup_job_registry(work_dir: Path) -> SetupJobRegistry:
    """Process-global registry for one managed data directory, mirroring
    `qc_tool.runqueue.get_manager`'s own per-work-dir singleton pattern.
    """
    key = work_dir.resolve()
    with _REGISTRIES_LOCK:
        registry = _REGISTRIES.get(key)
        if registry is None:
            registry = SetupJobRegistry()
            _REGISTRIES[key] = registry
        return registry


def _excel_members(files: dict[str, str]) -> dict[str, tuple[str | None, str | None]]:
    """``{member_id: (baseline_path_or_None, current_path_or_None)}`` for
    every Excel role present, parsing the same ``role`` / ``role:member``
    convention the rest of the app already uses.
    """
    members: dict[str, list[str | None]] = {}

    def _member(prefix: str, role: str) -> str:
        _, _, member_id = role.partition(":")
        return member_id or "primary"

    for role, path in files.items():
        if role == _EXCEL_BASELINE_PREFIX or role.startswith(_EXCEL_BASELINE_PREFIX + ":"):
            member_id = _member(_EXCEL_BASELINE_PREFIX, role)
            members.setdefault(member_id, [None, None])[0] = path
        elif role == _EXCEL_CURRENT_PREFIX or role.startswith(_EXCEL_CURRENT_PREFIX + ":"):
            member_id = _member(_EXCEL_CURRENT_PREFIX, role)
            members.setdefault(member_id, [None, None])[1] = path
    return {member_id: (pair[0], pair[1]) for member_id, pair in members.items()}


def _config_session_store(work_dir: Path) -> ConfigSessionStore:
    return ConfigSessionStore(work_dir / "history.sqlite3")


def build_session_choices(
    *,
    mode: QCRunMode,
    profile_name: str,
    files: dict[str, str],
    file_hashes: dict[str, str],
    output_mode: str,
    allow_large_workbooks: bool,
    allow_dependency_indexing: bool,
    acceptance_absolute: float,
    acceptance_percent: float,
    rerun_of: int | None,
) -> dict[str, object]:
    """The primitive-only payload persisted into `ConfigSessionStore` when
    the main intake page hands off to this workspace.
    """
    return {
        "mode": mode.value,
        "profile_name": profile_name,
        "files": dict(files),
        "file_hashes": dict(file_hashes),
        "output_mode": output_mode,
        "allow_large_workbooks": allow_large_workbooks,
        "allow_dependency_indexing": allow_dependency_indexing,
        "acceptance_absolute": acceptance_absolute,
        "acceptance_percent": acceptance_percent,
        "rerun_of": rerun_of,
    }


def start_setup_scan(
    work_dir: Path,
    job: SetupJob,
    members: dict[str, tuple[str, str]],
    passwords: dict[str, str],
) -> None:
    """Launch (once) a background thread that scans every member
    sequentially, sharing the one exclusive slot with any other heavy work.
    A no-op if a scan has already started for this job.
    """
    if job.started:
        return
    job.started = True
    for member_id in members:
        job.member_status[member_id] = "running"

    def _run_all() -> None:
        for member_id, (baseline_path, current_path) in members.items():
            if job.cancel_event.is_set():
                job.member_status[member_id] = "cancelled"
                continue
            _run_one_member_scan(
                work_dir, job, member_id, baseline_path, current_path, passwords
            )

    threading.Thread(target=_run_all, daemon=True).start()


def _run_one_member_scan(
    work_dir: Path,
    job: SetupJob,
    member_id: str,
    baseline_path: str,
    current_path: str,
    passwords: dict[str, str],
) -> None:
    """Runs on a worker thread: acquires the shared exclusive slot, spawns
    the disposable owned scan subprocess, and records this member's
    outcome on `job`.
    """
    try:
        baseline_hash = sha256_file(Path(baseline_path)) if Path(baseline_path).exists() else ""
        current_hash = sha256_file(Path(current_path)) if Path(current_path).exists() else ""
        request = SetupScanRequest(
            member_id=member_id,
            baseline_path=baseline_path,
            current_path=current_path,
            baseline_hash=baseline_hash,
            current_hash=current_hash,
            passwords=dict(passwords),
        )
        try:
            outcome = run_exclusive(
                work_dir,
                f"setup-scan:{member_id}",
                lambda: run_setup_scan_worker(request, cancel_event=job.cancel_event),
            )
        except QueueBusyError as exc:
            job.member_status[member_id] = "failed"
            job.member_disclosures[member_id] = str(exc)
            return
        if job.cancel_event.is_set():
            job.member_status[member_id] = "cancelled"
            return
        if outcome.missing_credential_roles:
            job.member_status[member_id] = "failed"
            job.missing_credential_roles[member_id] = outcome.missing_credential_roles
            job.member_disclosures[member_id] = (
                "a password is required before setup analysis can continue"
            )
            return
        if outcome.result_payload is None:
            job.member_status[member_id] = "failed"
            job.member_disclosures[member_id] = outcome.disclosure or "setup analysis failed"
            return
        job.member_results[member_id] = _member_profile_from_payload(outcome.result_payload)
        job.member_status[member_id] = "done"
    except Exception as exc:  # the UI thread must never crash from a scan failure
        job.member_status[member_id] = "failed"
        job.member_disclosures[member_id] = (
            f"setup analysis failed unexpectedly ({type(exc).__name__})"
        )


def _member_profile_from_payload(payload: dict[str, Any]) -> MemberSetupProfile:
    """Rehydrate one member's typed scan result from its own already
    self-produced, already-validated JSON round-trip (see
    ``qc_tool.setup.preview_worker``'s own "asdict() then JSON round-trip"
    contract) -- not external/untrusted input, so plain key access plus a
    permissive ``Any`` shape is appropriate here rather than a defensive
    isinstance wall.
    """
    from qc_tool.excel.complexity import WorkbookComplexity
    from qc_tool.excel.ranked_identity import RankedTableCandidate
    from qc_tool.excel.regions import TableRegion
    from qc_tool.setup.models import DetectedRegion, SheetSetupProfile, XlsbRiskProfile

    def _region(raw: dict[str, Any]) -> DetectedRegion:
        region_raw: dict[str, Any] = raw["region"]
        candidate_raw: dict[str, Any] | None = raw.get("ranked_candidate")
        candidate = None
        if candidate_raw is not None:
            candidate = RankedTableCandidate(
                columns=tuple(candidate_raw["columns"]),
                non_blank_coverage=candidate_raw["non_blank_coverage"],
                unique_ratio=candidate_raw["unique_ratio"],
                key_overlap=candidate_raw["key_overlap"],
                formula_ratio=candidate_raw["formula_ratio"],
                displaced_ratio=candidate_raw["displaced_ratio"],
                mismatch_reduction=candidate_raw["mismatch_reduction"],
                projected_positional_mismatches=candidate_raw[
                    "projected_positional_mismatches"
                ],
                ordinal_columns=tuple(candidate_raw.get("ordinal_columns", ())),
                header_row=candidate_raw.get("header_row"),
                manual_review=bool(candidate_raw.get("manual_review", False)),
            )
        return DetectedRegion(
            region=TableRegion(
                sheet=region_raw["sheet"],
                min_row=region_raw["min_row"],
                min_col=region_raw["min_col"],
                max_row=region_raw["max_row"],
                max_col=region_raw["max_col"],
                orientation=region_raw["orientation"],
                header_row=region_raw.get("header_row"),
                key_col=region_raw.get("key_col"),
                period_axis=region_raw["period_axis"],
            ),
            ranked_candidate=candidate,
        )

    def _sheet(raw: dict[str, Any]) -> SheetSetupProfile:
        return SheetSetupProfile(
            sheet_name=raw["sheet_name"],
            hidden=bool(raw.get("hidden", False)),
            very_hidden=bool(raw.get("very_hidden", False)),
            regions=tuple(_region(r) for r in raw.get("regions", ())),
            failure_detail=raw.get("failure_detail", ""),
        )

    def _risk(raw: dict[str, Any] | None):
        if raw is None:
            return None
        return XlsbRiskProfile(
            risky_features=tuple(raw.get("risky_features", ())),
            passive_features=tuple(raw.get("passive_features", ())),
            blocking_features=tuple(raw.get("blocking_features", ())),
            unknown_external_features=tuple(raw.get("unknown_external_features", ())),
            safe_for_external_engine=bool(raw.get("safe_for_external_engine", False)),
        )

    def _complexity(raw: dict[str, Any] | None):
        if raw is None:
            return None
        return WorkbookComplexity(
            formula_count=raw.get("formula_count", 0),
            lexical_formula_count=raw.get("lexical_formula_count", 0),
            reference_operands=raw.get("reference_operands", 0),
            resolved_range_cells=raw.get("resolved_range_cells", 0),
            projected_concrete_edges=raw.get("projected_concrete_edges", 0),
            interaction_rule_count=raw.get("interaction_rule_count", 0),
            warning_reasons=tuple(raw.get("warning_reasons", ())),
            override_used=bool(raw.get("override_used", False)),
            sampled_formulas=raw.get("sampled_formulas", 0),
            extras=dict(raw.get("extras", {})),
        )

    return MemberSetupProfile(
        member_id=payload["member_id"],
        baseline_hash=payload["baseline_hash"],
        current_hash=payload["current_hash"],
        baseline_sheets=tuple(_sheet(s) for s in payload.get("baseline_sheets", ())),
        current_sheets=tuple(_sheet(s) for s in payload.get("current_sheets", ())),
        baseline_complexity=_complexity(payload.get("baseline_complexity")),
        current_complexity=_complexity(payload.get("current_complexity")),
        baseline_xlsb_risk=_risk(payload.get("baseline_xlsb_risk")),
        current_xlsb_risk=_risk(payload.get("current_xlsb_risk")),
        failure_detail=payload.get("failure_detail", ""),
    )


_MODE_LABELS = {
    QCRunMode.CYCLE_COMPARISON: "Cycle comparison",
    QCRunMode.CURRENT_FILE_PREFLIGHT: "Current-file preflight",
    QCRunMode.FINAL_PACKAGE: "Final-package QC",
}
_REGION_MODE_OPTIONS = {
    "automatic": "Automatic",
    "keyed": "Match rows by key",
    "positional": "Compare by position",
    "excluded": "Exclude from this run",
}


def _coerce_str_dict(value: object) -> dict[str, str]:
    """Narrow a `ConfigSessionRecord.choices` value (stored as plain
    `dict[str, object]` since the session store round-trips arbitrary JSON)
    back to the flat string-keyed/string-valued shape `build_session_choices`
    always writes.
    """
    if not isinstance(value, dict):
        return {}
    return {str(key): str(item) for key, item in value.items()}


def _coerce_str_tuple(value: object) -> tuple[str, ...]:
    if not isinstance(value, (list, tuple)):
        return ()
    return tuple(str(item) for item in value)


def _coerce_float(value: object, default: float) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return default
    return float(value)


def _coerce_int(value: object, default: int) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        return default
    return value


def _coerce_optional_int(value: object) -> int | None:
    if isinstance(value, bool) or not isinstance(value, int):
        return None
    return value


def _profile_hash_or_none(profiles_dir: Path, name: str) -> str | None:
    """The saved profile file's content hash, or `None` when no file exists
    under that name yet (Step 11's "optimistic conflict protection"
    criterion: this is the value captured at the moment a profile name was
    selected in this session, compared against the file's hash again right
    before a save actually overwrites it).
    """
    path = profile_path(profiles_dir, name)
    if not path.exists():
        return None
    return sha256_file(path)


def _role_display_name(path_str: str) -> str:
    try:
        return Path(path_str).name
    except (TypeError, ValueError):
        return path_str


def _serialize_region(region) -> dict[str, object]:
    return {
        "region_id": region.region_id,
        "mode": region.mode,
        "header_intent": region.header_intent,
        "first_data_row": region.first_data_row,
        "preamble_rows": region.preamble_rows,
        "footer_rows": region.footer_rows,
        "baseline_range": region.baseline_range,
        "identity_columns": list(region.identity_columns),
        "ordinal_columns": list(region.ordinal_columns),
        "ignore_columns": list(region.ignore_columns),
        "expected_refresh_columns": list(region.expected_refresh_columns),
        "trim_identity_whitespace": region.trim_identity_whitespace,
        "blank_key_policy": region.blank_key_policy,
        "duplicate_key_policy": region.duplicate_key_policy,
        "exclusion_reason": region.exclusion_reason,
        "exclusion_expires_on": region.exclusion_expires_on,
        "ignore_columns_reason": region.ignore_columns_reason,
        "ignore_columns_expires_on": region.ignore_columns_expires_on,
        "confirmed": region.confirmed,
    }


def _apply_saved_region_choices(
    member: MemberReview, saved: dict[str, dict[str, object]]
) -> MemberReview:
    """Overlay previously saved decisions (draft recovery) onto a freshly
    scanned member review, matched by ``region_id``. A region the fresh
    scan no longer has is silently dropped -- the scan is authoritative for
    structure, the session store only for choices.
    """
    from dataclasses import replace as _replace

    new_sheets = []
    for sheet in member.current_sheets:
        new_regions = []
        for region in sheet.regions:
            saved_region = saved.get(region.region_id)
            if saved_region is None:
                new_regions.append(region)
                continue
            new_regions.append(
                _replace(
                    region,
                    mode=saved_region.get("mode", region.mode),
                    header_intent=saved_region.get("header_intent", region.header_intent),
                    first_data_row=(
                        _coerce_optional_int(saved_region["first_data_row"])
                        if "first_data_row" in saved_region
                        else region.first_data_row
                    ),
                    preamble_rows=_coerce_int(
                        saved_region.get("preamble_rows"), region.preamble_rows
                    ),
                    footer_rows=_coerce_int(saved_region.get("footer_rows"), region.footer_rows),
                    baseline_range=str(
                        saved_region.get("baseline_range", region.baseline_range) or ""
                    ),
                    identity_columns=_coerce_str_tuple(saved_region.get("identity_columns", ())),
                    ordinal_columns=_coerce_str_tuple(saved_region.get("ordinal_columns", ())),
                    ignore_columns=_coerce_str_tuple(saved_region.get("ignore_columns", ())),
                    expected_refresh_columns=_coerce_str_tuple(
                        saved_region.get("expected_refresh_columns", ())
                    ),
                    trim_identity_whitespace=bool(
                        saved_region.get(
                            "trim_identity_whitespace", region.trim_identity_whitespace
                        )
                    ),
                    blank_key_policy=saved_region.get("blank_key_policy", region.blank_key_policy),
                    duplicate_key_policy=saved_region.get(
                        "duplicate_key_policy", region.duplicate_key_policy
                    ),
                    exclusion_reason=saved_region.get("exclusion_reason", ""),
                    exclusion_expires_on=saved_region.get("exclusion_expires_on", ""),
                    ignore_columns_reason=saved_region.get("ignore_columns_reason", ""),
                    ignore_columns_expires_on=saved_region.get("ignore_columns_expires_on", ""),
                    confirmed=bool(saved_region.get("confirmed", False)),
                )
            )
        new_sheets.append(_replace(sheet, regions=tuple(new_regions)))
    return _replace(member, current_sheets=tuple(new_sheets))


def _serialize_selectors(member: MemberReview) -> list[dict[str, object]]:
    entries: list[dict[str, object]] = []
    for sheet in member.current_sheets:
        for selector in sheet.selectors:
            entries.append(
                {
                    "sheet_name": sheet.sheet_name,
                    "selector_id": selector.selector_id,
                    "label": selector.label,
                    "cell": selector.cell,
                }
            )
    return entries


def _apply_saved_selectors(member: MemberReview, saved: list[dict[str, object]]) -> MemberReview:
    """Restore analyst-declared selector prerequisites (draft recovery),
    matched by sheet name -- the scan is authoritative for which sheets
    exist, the session store only for which selectors an analyst added.
    """
    from dataclasses import replace as _replace

    by_sheet: dict[str, list[SelectorDecision]] = {}
    for entry in saved:
        sheet_name = str(entry.get("sheet_name", ""))
        if not sheet_name:
            continue
        by_sheet.setdefault(sheet_name, []).append(
            SelectorDecision(
                selector_id=str(entry.get("selector_id", "")),
                label=str(entry.get("label", "")),
                cell=str(entry.get("cell", "")),
            )
        )
    if not by_sheet:
        return member
    new_sheets = []
    for sheet in member.current_sheets:
        restored = by_sheet.get(sheet.sheet_name)
        if not restored:
            new_sheets.append(sheet)
            continue
        new_sheets.append(_replace(sheet, selectors=tuple(restored)))
    return _replace(member, current_sheets=tuple(new_sheets))


def _render_preview_grid(outcome: Any, *, on_cell_click) -> None:
    """Render a bounded, read-only preview window as a plain HTML table.
    Every cell is clickable (Step 7's "click controls" criterion) so an
    armed anchor-pick action can resolve a cell without drag selection.
    """
    with ui.element("div").classes("previewscroll"), ui.element("table").classes("previewgrid"):
        with ui.element("thead"), ui.element("tr"):
            ui.element("th")
            for col_offset in range(len(outcome.rows[0]) if outcome.rows else 0):
                col_number = outcome.resolved_min_col + col_offset
                with ui.element("th"):
                    ui.label(get_column_letter(col_number))
        with ui.element("tbody"):
            for row_offset, row_values in enumerate(outcome.rows):
                row_number = outcome.resolved_min_row + row_offset
                formula_row = (
                    outcome.formula_cells[row_offset]
                    if row_offset < len(outcome.formula_cells)
                    else []
                )
                with ui.element("tr"):
                    with ui.element("th"):
                        ui.label(str(row_number))
                    for col_offset, text in enumerate(row_values):
                        col_number = outcome.resolved_min_col + col_offset
                        is_formula = (
                            bool(formula_row[col_offset])
                            if col_offset < len(formula_row)
                            else False
                        )
                        cell = ui.element("td").classes("previewcell")
                        if is_formula:
                            cell.classes("previewcell-formula")
                        with cell:
                            ui.label(text if text else "\u00a0")
                        cell.on(
                            "click",
                            lambda _event, r=row_number, c=col_number: on_cell_click(r, c),
                        )


def render_config_workspace(
    work_dir: Path,
    profiles_dir: Path,
    session_key: str,
    *,
    queue_manager: RunQueueManager,
    network_mode: str = "local",
    on_settings=None,
    on_shutdown=None,
) -> None:
    """Render the full-page mode-aware configuration workspace for one
    already-created config session. Registered by `qc_tool.ui.app` at
    ``/configure``.
    """
    session_store = _config_session_store(work_dir)
    record = session_store.get(session_key)
    if record is None or "files" not in record.choices:
        with page_frame(
            "run", network_mode=network_mode, on_settings=on_settings, on_shutdown=on_shutdown
        ):
            ui.label("This configuration session was not found").classes("runhead")
            ui.label(
                "It may have expired, or the link is stale. Return to Compare and "
                "select your files again."
            ).classes("note")
            ui.link("Back to Compare", "/").classes("guide-jump")
        return

    choices = record.choices
    raw_files: dict[str, str] = _coerce_str_dict(choices.get("files", {}))
    file_hashes: dict[str, str] = _coerce_str_dict(choices.get("file_hashes", {}))
    try:
        mode = QCRunMode(str(choices.get("mode", "cycle_comparison")))
    except ValueError:
        mode = QCRunMode.CYCLE_COMPARISON

    invalid_roles: list[str] = []
    files: dict[str, Path] = {}
    for role, path_str in raw_files.items():
        path = Path(path_str)
        expected_hash = file_hashes.get(role)
        if not path.exists() or (expected_hash and sha256_file(path) != expected_hash):
            invalid_roles.append(role)
            continue
        files[role] = path

    with page_frame(
        "run", network_mode=network_mode, on_settings=on_settings, on_shutdown=on_shutdown
    ):
        ui.label(f"Configure & run — {_MODE_LABELS[mode]}").classes("pagetitle")
        ui.label(
            "Review the structure of your files, resolve how rows and columns "
            "match, and confirm before QC runs. This scan is bounded and never "
            "becomes evidence -- the actual run reloads your files fresh."
        ).classes("lede")

        if invalid_roles:
            ui.label(
                "Some selected files are missing or changed since intake: "
                + ", ".join(sorted(invalid_roles))
            ).classes("notecard")
            ui.link("Back to Compare to select files again", "/").classes("guide-jump")
            return

        section("1 · Input roles")
        with ui.column().classes("gap-1 w-full"):
            for role in sorted(files):
                with ui.row().classes("items-center gap-2"):
                    status_chip("ok", role.replace("_", " "), _role_display_name(raw_files[role]))

        members = _excel_members({role: str(path) for role, path in files.items()})
        # Preflight/final-package have no baseline side -- self-scan the
        # current file so structural facts (sheets/regions/visibility) are
        # still reviewable without inventing a fake baseline concept.
        scan_members: dict[str, tuple[str, str]] = {}
        for member_id, (baseline_path, current_path) in members.items():
            if current_path is None:
                continue
            scan_members[member_id] = (baseline_path or current_path, current_path)

        registry = get_setup_job_registry(work_dir)
        job = registry.get_or_create(session_key)

        section("2 · Setup analysis")
        scan_status_box = ui.column().classes("gap-2 w-full")
        profile_box = ui.column().classes("gap-2 w-full")
        regions_box = ui.column().classes("gap-2 w-full")
        deck_box = ui.column().classes("gap-2 w-full")
        section("3 · Preview data")
        preview_box = ui.column().classes("gap-2 w-full")
        warnings_box = ui.column().classes("gap-2 w-full")
        advanced_box = ui.column().classes("gap-2 w-full")
        actions_box = ui.row().classes("items-center gap-2 flex-wrap")

        workspace_state: dict[str, ConfigWorkspaceState] = {
            "value": ConfigWorkspaceState(
                mode=mode,
                profile_name=str(choices.get("profile_name", "default")),
                allow_large_workbooks=bool(choices.get("allow_large_workbooks", False)),
                allow_dependency_indexing=bool(choices.get("allow_dependency_indexing", False)),
                profile_opened_hash=_profile_hash_or_none(
                    profiles_dir, str(choices.get("profile_name", "default"))
                ),
            )
        }
        selected_profile: dict[str, DeliverableProfile | None] = {"value": None}
        #: One in-memory preview panel: member/sheet/window position plus the
        #: last fetched outcome. Never persisted -- a refresh regenerates it
        #: on demand (matches the plan's own "preview content regenerates,
        #: choices restore" boundary).
        preview_state: dict[str, Any] = {
            "member_id": None,
            "sheet_name": None,
            "min_row": 1,
            "min_col": 1,
            "reveal_formulas": False,
            "loading": False,
            "error": "",
            "outcome": None,
            #: (member_id, sheet_name, region_id) currently armed to receive
            #: the next preview-grid cell click as its new anchor, or None.
            "anchor_target": None,
        }

        def _persist_choices() -> None:
            state = workspace_state["value"]
            region_choices: dict[str, dict[str, object]] = {}
            selector_choices: dict[str, list[dict[str, object]]] = {}
            rename_choices: dict[str, dict[str, str]] = {}
            for member in state.member_reviews:
                for sheet in member.current_sheets:
                    for region in sheet.regions:
                        region_choices[region.region_id] = _serialize_region(region)
                member_selectors = _serialize_selectors(member)
                if member_selectors:
                    selector_choices[member.member_id] = member_selectors
                if member.sheet_renames:
                    rename_choices[member.member_id] = dict(member.sheet_renames)
            deck = state.deck_review
            excluded_slides = (
                [s.slide_index for s in deck.current_slides if not s.included]
                if deck is not None
                else []
            )
            slide_renames = dict(deck.slide_renames) if deck is not None else {}
            session_store.save_choices(
                session_key,
                profile_name=state.profile_name,
                choices={
                    **choices,
                    "region_decisions": region_choices,
                    "selectors": selector_choices,
                    "sheet_renames": rename_choices,
                    "excluded_slides": excluded_slides,
                    "slide_renames": slide_renames,
                    "allow_large_workbooks": state.allow_large_workbooks,
                    "allow_dependency_indexing": state.allow_dependency_indexing,
                },
            )

        def _seed_member_reviews() -> None:
            result = job.result
            if result is None:
                return
            saved_regions_raw = choices.get("region_decisions", {})
            saved_regions: dict[str, dict[str, object]] = (
                dict(saved_regions_raw) if isinstance(saved_regions_raw, dict) else {}
            )
            saved_selectors_raw = choices.get("selectors", {})
            saved_selectors: dict[str, object] = (
                dict(saved_selectors_raw) if isinstance(saved_selectors_raw, dict) else {}
            )
            saved_renames_raw = choices.get("sheet_renames", {})
            saved_renames: dict[str, object] = (
                dict(saved_renames_raw) if isinstance(saved_renames_raw, dict) else {}
            )
            reviews = []
            for member_id, member_profile in result.members.items():
                review = member_review_from_scan(member_id, member_profile)
                if saved_regions:
                    review = _apply_saved_region_choices(review, saved_regions)
                member_selectors = saved_selectors.get(member_id)
                if isinstance(member_selectors, list):
                    review = _apply_saved_selectors(review, member_selectors)
                member_renames = saved_renames.get(member_id)
                if isinstance(member_renames, dict):
                    review = dataclasses.replace(
                        review,
                        sheet_renames={str(k): str(v) for k, v in member_renames.items()},
                    )
                reviews.append(review)
            workspace_state["value"] = dataclasses.replace(
                workspace_state["value"], member_reviews=tuple(reviews)
            )

        def _seed_deck_review() -> None:
            """Peek PowerPoint slide titles (Step 9's "cycle slide inclusion
            and explicit renamed-slide pins" / "preflight current-side
            setup" criteria) -- a cheap, direct, in-process read (unlike
            Excel's disposable-subprocess setup scan; a deck is small
            enough that no bounded worker is warranted). Runs exactly
            once: `deck_review` starts `None` and this always leaves it a
            real `DeckReview` (empty when there is no PPT file at all), so
            the `poll()` guard below never re-seeds and silently discards
            an analyst's in-progress inclusion/rename choices.
            """
            from qc_tool.io.peek import peek_slide_titles

            current_ppt = files.get(_PPT_CURRENT_PREFIX)
            if current_ppt is None:
                workspace_state["value"] = dataclasses.replace(
                    workspace_state["value"], deck_review=DeckReview()
                )
                return
            baseline_ppt = files.get(_PPT_BASELINE_PREFIX)
            baseline_titles = peek_slide_titles(baseline_ppt) if baseline_ppt is not None else []
            current_titles = peek_slide_titles(current_ppt)
            deck = deck_review_from_titles(baseline_titles, current_titles)
            saved_excluded_raw = choices.get("excluded_slides", [])
            if isinstance(saved_excluded_raw, list):
                excluded = {int(i) for i in saved_excluded_raw if str(i).lstrip("-").isdigit()}
                if excluded:
                    deck = dataclasses.replace(
                        deck,
                        current_slides=tuple(
                            dataclasses.replace(s, included=False)
                            if s.slide_index in excluded
                            else s
                            for s in deck.current_slides
                        ),
                    )
            saved_slide_renames_raw = choices.get("slide_renames", {})
            if isinstance(saved_slide_renames_raw, dict):
                deck = dataclasses.replace(
                    deck,
                    slide_renames={
                        str(k): str(v) for k, v in saved_slide_renames_raw.items()
                    },
                )
            workspace_state["value"] = dataclasses.replace(
                workspace_state["value"], deck_review=deck
            )

        def _region_by_id(region_id: str):
            for member in workspace_state["value"].member_reviews:
                for sheet in member.current_sheets:
                    for region in sheet.regions:
                        if region.region_id == region_id:
                            return member.member_id, sheet.sheet_name, region
            return None, None, None

        def _update_region(member_id: str, sheet_name: str, region_id: str, **changes) -> None:
            state = workspace_state["value"]
            member = state.member_review(member_id)
            if member is None:
                return
            updated_member = update_region_decision(member, sheet_name, region_id, **changes)
            workspace_state["value"] = state.with_member_review(updated_member)
            _persist_choices()
            refresh()

        def _apply_region_transform(
            member_id: str, sheet_name: str, region_id: str, transform
        ) -> None:
            """Backs the "click controls"/"A1 controls" criterion: a
            transform that must recompute several region fields together
            (anchor, range, available columns) instead of one field at a
            time. A `ValueError` raised by `transform` (an unparseable typed
            range) is surfaced as a notification, never a crash.
            """
            state = workspace_state["value"]
            member = state.member_review(member_id)
            if member is None:
                return
            try:
                updated_member = apply_region_transform(member, sheet_name, region_id, transform)
            except ValueError as exc:
                ui.notify(str(exc), type="warning")
                return
            workspace_state["value"] = state.with_member_review(updated_member)
            _persist_choices()
            refresh()

        def _confirm_all_regions() -> None:
            """Bulk confirmation (Step 7 criterion): mark every currently
            valid, detected region across every member as analyst-confirmed
            in one action.
            """
            state = workspace_state["value"]
            new_reviews = tuple(confirm_all_regions(member) for member in state.member_reviews)
            workspace_state["value"] = dataclasses.replace(state, member_reviews=new_reviews)
            _persist_choices()
            refresh()

        def _update_sheet_rename(
            member_id: str, current_name: str, baseline_name: str | None
        ) -> None:
            """Declare or clear a renamed-sheet pairing (Step 8's "renamed
            sheets require explicit mapping" criterion).
            """
            state = workspace_state["value"]
            member = state.member_review(member_id)
            if member is None:
                return
            updated_member = set_sheet_rename(member, current_name, baseline_name)
            workspace_state["value"] = state.with_member_review(updated_member)
            _persist_choices()
            refresh()

        def _render_sheet_rename_controls(
            member: MemberReview, added: tuple[str, ...], removed: tuple[str, ...]
        ) -> None:
            """One dropdown per added (current-only) sheet, offering every
            still-unclaimed removed (baseline-only) sheet as a rename
            target -- promotes an added+removed pair into one logical sheet
            instead of two separate acknowledgements.
            """
            claimed = set(member.sheet_renames.values())
            for current_name in added:
                declared = member.sheet_renames.get(current_name)
                options = {"": "New sheet (no rename)"}
                for baseline_name in removed:
                    if baseline_name == declared or baseline_name not in claimed:
                        options[baseline_name] = f"Same as {baseline_name!r}"

                def _on_rename_change(
                    event: events.ValueChangeEventArguments,
                    member_id=member.member_id,
                    current_name=current_name,
                ) -> None:
                    value = str(event.value or "")
                    _update_sheet_rename(member_id, current_name, value or None)

                with ui.row().classes("items-center gap-2"):
                    ui.label(f"{current_name!r} (new)").classes("note")
                    ui.select(
                        options, value=declared or "", on_change=_on_rename_change
                    ).props("outlined dense").classes("w-56")
            still_removed = [name for name in removed if name not in claimed]
            for baseline_name in still_removed:
                ui.label(f"{baseline_name!r} removed from the current file").classes("note")

        def _set_slide_included(slide_index: int, included: bool) -> None:
            state = workspace_state["value"]
            deck = state.deck_review
            if deck is None:
                return
            workspace_state["value"] = dataclasses.replace(
                state, deck_review=set_slide_included(deck, slide_index, included)
            )
            _persist_choices()
            refresh()

        def _update_slide_rename(current_title: str, baseline_title: str | None) -> None:
            """Declare or clear a renamed-slide pin (Step 9's "explicit
            renamed-slide pins" criterion) -- mirrors `_update_sheet_rename`.
            """
            state = workspace_state["value"]
            deck = state.deck_review
            if deck is None:
                return
            workspace_state["value"] = dataclasses.replace(
                state, deck_review=set_slide_rename(deck, current_title, baseline_title)
            )
            _persist_choices()
            refresh()

        def _render_slide_rename_controls(
            deck: DeckReview, added: tuple[str, ...], removed: tuple[str, ...]
        ) -> None:
            """Mirrors `_render_sheet_rename_controls` exactly, for slide
            titles instead of sheet names.
            """
            claimed = set(deck.slide_renames.values())
            for current_title in added:
                declared = deck.slide_renames.get(current_title)
                options = {"": "New slide (no rename)"}
                for baseline_title in removed:
                    if baseline_title == declared or baseline_title not in claimed:
                        options[baseline_title] = f"Same as {baseline_title!r}"

                def _on_rename_change(
                    event: events.ValueChangeEventArguments, current_title=current_title
                ) -> None:
                    value = str(event.value or "")
                    _update_slide_rename(current_title, value or None)

                with ui.row().classes("items-center gap-2"):
                    ui.label(f"{current_title!r} (new)").classes("note")
                    ui.select(
                        options, value=declared or "", on_change=_on_rename_change
                    ).props("outlined dense").classes("w-56")
            still_removed = [title for title in removed if title not in claimed]
            for baseline_title in still_removed:
                ui.label(f"{baseline_title!r} removed from the current deck").classes("note")

        def _add_selector(member_id: str, sheet_name: str, label: str, cell: str) -> None:
            state = workspace_state["value"]
            member = state.member_review(member_id)
            if member is None:
                return
            try:
                updated_member = add_selector(member, sheet_name, label=label, cell=cell)
            except ValueError as exc:
                ui.notify(str(exc), type="warning")
                return
            workspace_state["value"] = state.with_member_review(updated_member)
            _persist_choices()
            refresh()

        def _remove_selector(member_id: str, sheet_name: str, selector_id: str) -> None:
            state = workspace_state["value"]
            member = state.member_review(member_id)
            if member is None:
                return
            updated_member = remove_selector(member, sheet_name, selector_id)
            workspace_state["value"] = state.with_member_review(updated_member)
            _persist_choices()
            refresh()

        def _render_add_selector_control(member: MemberReview) -> None:
            """A compact, always-available control (Step 8's "explicit
            selector prerequisites for dropdown/filter/scenario/parameter
            cells" criterion) -- reachable even for a sheet with no
            detected table regions, since a selector cell rarely lives
            inside one.
            """
            sheet_names = [sheet.sheet_name for sheet in member.current_sheets]
            if not sheet_names:
                return
            with ui.row().classes("items-center gap-2 flex-wrap"):
                ui.label("Add selector prerequisite:").classes("note")
                sheet_select = (
                    ui.select(sheet_names, value=sheet_names[0], label="Sheet")
                    .props("outlined dense")
                    .classes("w-40")
                )
                cell_input = ui.input("Cell (A1)").props("outlined dense").classes("w-32")
                label_input = ui.input("Label").props("outlined dense").classes("w-40")

                def _on_add(member_id=member.member_id) -> None:
                    _add_selector(
                        member_id,
                        str(sheet_select.value),
                        str(label_input.value or ""),
                        str(cell_input.value or ""),
                    )
                    cell_input.set_value("")
                    label_input.set_value("")

                ui.button("Add", on_click=_on_add).props("flat dense no-caps")

        def _render_selector_row(member_id: str, sheet_name: str, selector) -> None:
            with ui.row().classes("items-center gap-2"):
                ui.label(f"{selector.label} ({selector.cell})").classes("note")

                def _on_remove(
                    member_id=member_id, sheet_name=sheet_name, selector_id=selector.selector_id
                ) -> None:
                    _remove_selector(member_id, sheet_name, selector_id)

                ui.button(icon="close", on_click=_on_remove).props("flat dense round size=sm")

        def _toggle_warning(code: str, value: bool) -> None:
            state = workspace_state["value"]
            acknowledged = set(state.warnings_acknowledged)
            if value:
                acknowledged.add(code)
            else:
                acknowledged.discard(code)
            workspace_state["value"] = dataclasses.replace(
                state, warnings_acknowledged=frozenset(acknowledged)
            )
            refresh()

        def render_scan_status() -> None:
            scan_status_box.clear()
            with scan_status_box:
                if not scan_members:
                    ui.label(
                        "No Excel workbook is available to analyze for this mode yet."
                    ).classes("note")
                    return
                status = job.overall_status
                if status == "idle":
                    ui.label("Setup analysis will start automatically.").classes("note")
                elif status == "running":
                    with ui.row().classes("items-center gap-2"):
                        ui.spinner(size="1.2rem")
                        ui.label("Analyzing structure...").classes("note")
                        ui.button(
                            "Cancel", on_click=lambda: (job.cancel_event.set(), refresh())
                        ).props("flat no-caps dense")
                elif status == "done":
                    total_regions = sum(
                        len(sheet.regions)
                        for member in job.result.members.values()
                        for sheet in member.current_sheets
                    ) if job.result else 0
                    status_chip(
                        "ok",
                        "Analysis complete",
                        f"{len(scan_members)} member(s), {total_regions} region(s) detected",
                    )
                elif status == "cancelled":
                    ui.label("Setup analysis was cancelled.").classes("notecard")
                    ui.button(
                        "Retry analysis",
                        on_click=lambda: (registry.reset(session_key), refresh()),
                    ).props("flat no-caps dense")
                else:
                    disclosures = "; ".join(sorted(set(job.member_disclosures.values())))
                    ui.label(f"Setup analysis failed: {disclosures}").classes("notecard")
                    ui.button(
                        "Retry analysis",
                        on_click=lambda: (registry.reset(session_key), refresh()),
                    ).props("flat no-caps dense")

        def render_profile_section() -> None:
            profile_box.clear()
            state = workspace_state["value"]
            with profile_box:
                if job.overall_status != "done":
                    ui.label(
                        "Choose a saved profile once setup analysis finishes."
                    ).classes("note")
                    return
                options = list_profiles(profiles_dir)
                current_name = state.profile_name if state.profile_name in options else "default"

                def _on_profile_change(event: events.ValueChangeEventArguments) -> None:
                    new_name = str(event.value)
                    workspace_state["value"] = dataclasses.replace(
                        workspace_state["value"],
                        profile_name=new_name,
                        profile_opened_hash=_profile_hash_or_none(profiles_dir, new_name),
                    )
                    _persist_choices()
                    refresh()

                ui.select(
                    options,
                    value=current_name,
                    label="Deliverable profile",
                    on_change=_on_profile_change,
                ).classes("w-64").props("outlined dense")
                try:
                    profile = load_profile_by_name(profiles_dir, current_name)
                except Exception:
                    profile = None
                selected_profile["value"] = profile
                if profile is None or job.result is None:
                    return
                diff = diff_profile_against_scan(profile, job.result, mode=mode)
                if is_clean_profile_diff(diff) or not diff:
                    ui.label(
                        "This profile matches the current scan -- nothing new to "
                        "confirm."
                    ).classes("note")
                else:
                    with ui.expansion(
                        f"Configuration diff ({len(diff)} item(s))", value=False
                    ).classes("w-full"):
                        for entry in diff:
                            ui.label(f"[{entry.kind}] {entry.description}").classes("note")

        def render_regions_section() -> None:
            regions_box.clear()
            state = workspace_state["value"]
            with regions_box:
                if job.overall_status != "done" or not state.member_reviews:
                    return
                any_region = any(
                    sheet.regions
                    for member in state.member_reviews
                    for sheet in member.current_sheets
                )
                if any_region:
                    ui.button(
                        "Confirm all detected regions", on_click=lambda: _confirm_all_regions()
                    ).props("flat no-caps dense")
                for member in state.member_reviews:
                    pairs, added, removed = sheet_pairing(member)
                    if added or removed:
                        note = []
                        if pairs:
                            note.append(f"{len(pairs)} sheet(s) matched by name")
                        ui.label(f"[{member.member_id}] " + "; ".join(note)).classes("note")
                        _render_sheet_rename_controls(member, added, removed)
                    _render_add_selector_control(member)
                    for sheet in member.current_sheets:
                        if not sheet.regions and not sheet.selectors:
                            continue
                        with ui.card().classes("w-full p-3"):
                            visibility = (
                                "very hidden"
                                if sheet.very_hidden
                                else "hidden"
                                if sheet.hidden
                                else "visible"
                            )
                            ui.label(f"{sheet.sheet_name} ({visibility})").classes("runhead")
                            for region in sheet.regions:
                                _render_region_row(member.member_id, sheet.sheet_name, region)
                            if sheet.selectors:
                                ui.label("Selector prerequisites").classes("dk")
                                for selector in sheet.selectors:
                                    _render_selector_row(
                                        member.member_id, sheet.sheet_name, selector
                                    )

        def render_deck_section() -> None:
            """PowerPoint slide review (Step 9): cycle slide inclusion plus
            explicit renamed-slide pins for CYCLE_COMPARISON; a plain
            current-side inventory for preflight/final-package (no
            baseline deck, so no pairing/inclusion concept applies).
            """
            deck_box.clear()
            state = workspace_state["value"]
            deck = state.deck_review
            with deck_box:
                if deck is None or not deck.current_slides:
                    return
                ui.label("PowerPoint slides").classes("dk")
                if mode is QCRunMode.CYCLE_COMPARISON and deck.baseline_slides:
                    pairs, added, removed = effective_slide_pairing(deck)
                    if pairs:
                        ui.label(f"{len(pairs)} slide(s) matched by title").classes("note")
                    if added or removed:
                        _render_slide_rename_controls(deck, added, removed)
                for slide in deck.current_slides:
                    with ui.row().classes("items-center gap-2"):
                        if mode is QCRunMode.CYCLE_COMPARISON:
                            ui.checkbox(
                                value=slide.included,
                                on_change=(
                                    lambda event, idx=slide.slide_index: _set_slide_included(
                                        idx, bool(event.value)
                                    )
                                ),
                            ).mark(f"slide-include-{slide.slide_index}")
                        ui.label(f"{slide.slide_index}. {slide.title}").classes("note")

        def _render_region_row(member_id: str, sheet_name: str, region) -> None:
            with ui.column().classes("gap-1 w-full"), ui.card().classes("p-2 w-full"):
              with ui.row().classes("items-center gap-2 flex-wrap w-full"):
                ui.label(f"anchor {region.anchor_cell}").classes("dk")

                def _on_range_change(
                    event: events.ValueChangeEventArguments,
                    member_id=member_id,
                    sheet_name=sheet_name,
                    region_id=region.region_id,
                ) -> None:
                    text = str(event.value or "")
                    _apply_region_transform(
                        member_id, sheet_name, region_id, lambda r: apply_manual_range(r, text)
                    )

                ui.input(
                    "Range (A1)", value=region.current_range, on_change=_on_range_change
                ).props("outlined dense").classes("w-32")

                def _on_pick_anchor(
                    member_id=member_id, sheet_name=sheet_name, region_id=region.region_id
                ) -> None:
                    preview_state["member_id"] = member_id
                    preview_state["sheet_name"] = sheet_name
                    preview_state["anchor_target"] = (member_id, sheet_name, region_id)
                    render_preview_section()

                is_armed = preview_state["anchor_target"] == (
                    member_id,
                    sheet_name,
                    region.region_id,
                )
                ui.button(
                    "Click preview to set anchor" if is_armed else "Pick anchor from preview",
                    on_click=_on_pick_anchor,
                ).props("flat dense no-caps" + (" color=primary" if is_armed else ""))
                if region.ranked_candidate_pending and region.mode == "automatic":
                    ui.label("looks ranked/sorted").classes("notecard")

                def _on_mode_change(
                    event: events.ValueChangeEventArguments,
                    member_id=member_id,
                    sheet_name=sheet_name,
                    region_id=region.region_id,
                ) -> None:
                    _update_region(member_id, sheet_name, region_id, mode=str(event.value))

                ui.select(
                    _REGION_MODE_OPTIONS, value=region.mode, on_change=_on_mode_change
                ).props("outlined dense").classes("w-56")

                if region.mode == "keyed":

                    def _on_identity_change(
                        event: events.ValueChangeEventArguments,
                        member_id=member_id,
                        sheet_name=sheet_name,
                        region_id=region.region_id,
                    ) -> None:
                        cols = tuple(event.value or ())
                        _apply_region_transform(
                            member_id,
                            sheet_name,
                            region_id,
                            lambda r: dataclasses.replace(
                                set_identity_columns(r, cols), confirmed=True
                            ),
                        )

                    ui.select(
                        list(region.available_columns),
                        value=list(region.identity_columns),
                        multiple=True,
                        label="Identity columns",
                        on_change=_on_identity_change,
                    ).props("outlined dense use-chips").classes("w-56")

                    def _on_ordinal_change(
                        event: events.ValueChangeEventArguments,
                        member_id=member_id,
                        sheet_name=sheet_name,
                        region_id=region.region_id,
                    ) -> None:
                        cols = tuple(event.value or ())
                        _apply_region_transform(
                            member_id, sheet_name, region_id, lambda r: set_ordinal_columns(r, cols)
                        )

                    ui.select(
                        list(region.available_columns),
                        value=list(region.ordinal_columns),
                        multiple=True,
                        label="Ignore order in",
                        on_change=_on_ordinal_change,
                    ).props("outlined dense use-chips").classes("w-56")

                    def _on_trim_change(
                        event: events.ValueChangeEventArguments,
                        member_id=member_id,
                        sheet_name=sheet_name,
                        region_id=region.region_id,
                    ) -> None:
                        _update_region(
                            member_id,
                            sheet_name,
                            region_id,
                            trim_identity_whitespace=bool(event.value),
                        )

                    ui.checkbox(
                        "Trim outer whitespace",
                        value=region.trim_identity_whitespace,
                        on_change=_on_trim_change,
                    )

                    def _on_blank_key_change(
                        event: events.ValueChangeEventArguments,
                        member_id=member_id,
                        sheet_name=sheet_name,
                        region_id=region.region_id,
                    ) -> None:
                        _update_region(
                            member_id, sheet_name, region_id, blank_key_policy=str(event.value)
                        )

                    ui.select(
                        {
                            "system_default": "System default",
                            "tolerate": "Tolerate blank keys",
                            "block": "Block on blank keys",
                        },
                        value=region.blank_key_policy,
                        label="Blank key policy",
                        on_change=_on_blank_key_change,
                    ).props("outlined dense").classes("w-48")
                elif region.mode == "excluded":

                    def _on_reason_change(
                        event: events.ValueChangeEventArguments,
                        member_id=member_id,
                        sheet_name=sheet_name,
                        region_id=region.region_id,
                    ) -> None:
                        _update_region(
                            member_id,
                            sheet_name,
                            region_id,
                            exclusion_reason=str(event.value or ""),
                        )

                    ui.input(
                        "Reason", value=region.exclusion_reason, on_change=_on_reason_change
                    ).props(
                        "outlined dense"
                    ).classes("w-56")

                    def _on_expiry_change(
                        event: events.ValueChangeEventArguments,
                        member_id=member_id,
                        sheet_name=sheet_name,
                        region_id=region.region_id,
                    ) -> None:
                        _update_region(
                            member_id,
                            sheet_name,
                            region_id,
                            exclusion_expires_on=str(event.value or ""),
                            confirmed=True,
                        )

                    ui.input(
                        "Expires on", value=region.exclusion_expires_on, on_change=_on_expiry_change
                    ).props('outlined dense type=date').classes("w-40")
                if not region.is_valid:
                    ui.label("needs more detail").classes("notecard")

              with ui.row().classes("items-center gap-2 flex-wrap w-full"):

                def _on_header_intent_change(
                    event: events.ValueChangeEventArguments,
                    member_id=member_id,
                    sheet_name=sheet_name,
                    region_id=region.region_id,
                ) -> None:
                    _update_region(member_id, sheet_name, region_id, header_intent=str(event.value))

                ui.select(
                    {
                        "automatic": "Header: automatic",
                        "no_header": "Header: none",
                        "first_data_row": "Header: first data row",
                    },
                    value=region.header_intent,
                    on_change=_on_header_intent_change,
                ).props("outlined dense").classes("w-52")

                if region.header_intent == "first_data_row":

                    def _on_first_data_row_change(
                        event: events.ValueChangeEventArguments,
                        member_id=member_id,
                        sheet_name=sheet_name,
                        region_id=region.region_id,
                    ) -> None:
                        raw = str(event.value or "").strip()
                        _update_region(
                            member_id,
                            sheet_name,
                            region_id,
                            first_data_row=int(raw) if raw.isdigit() else None,
                        )

                    ui.input(
                        "First data row",
                        value=str(region.first_data_row or ""),
                        on_change=_on_first_data_row_change,
                    ).props("outlined dense type=number").classes("w-32")

                def _on_preamble_change(
                    event: events.ValueChangeEventArguments,
                    member_id=member_id,
                    sheet_name=sheet_name,
                    region_id=region.region_id,
                ) -> None:
                    raw = str(event.value or "0").strip()
                    _update_region(
                        member_id,
                        sheet_name,
                        region_id,
                        preamble_rows=int(raw) if raw.isdigit() else 0,
                    )

                ui.input(
                    "Preamble rows", value=str(region.preamble_rows), on_change=_on_preamble_change
                ).props("outlined dense type=number").classes("w-32")

                def _on_footer_change(
                    event: events.ValueChangeEventArguments,
                    member_id=member_id,
                    sheet_name=sheet_name,
                    region_id=region.region_id,
                ) -> None:
                    raw = str(event.value or "0").strip()
                    _update_region(
                        member_id,
                        sheet_name,
                        region_id,
                        footer_rows=int(raw) if raw.isdigit() else 0,
                    )

                ui.input(
                    "Footer rows", value=str(region.footer_rows), on_change=_on_footer_change
                ).props("outlined dense type=number").classes("w-32")

                def _on_baseline_range_change(
                    event: events.ValueChangeEventArguments,
                    member_id=member_id,
                    sheet_name=sheet_name,
                    region_id=region.region_id,
                ) -> None:
                    _update_region(
                        member_id, sheet_name, region_id, baseline_range=str(event.value or "")
                    )

                ui.input(
                    "Baseline range override (A1)",
                    value=region.baseline_range,
                    on_change=_on_baseline_range_change,
                ).props("outlined dense").classes("w-56")

              with (
                  ui.expansion("Ignore / expected-refresh columns").classes("w-full"),
                  ui.row().classes("items-center gap-2 flex-wrap w-full"),
              ):

                    def _on_ignore_change(
                        event: events.ValueChangeEventArguments,
                        member_id=member_id,
                        sheet_name=sheet_name,
                        region_id=region.region_id,
                    ) -> None:
                        cols = tuple(event.value or ())
                        _apply_region_transform(
                            member_id, sheet_name, region_id, lambda r: set_ignore_columns(r, cols)
                        )

                    ui.select(
                        list(region.available_columns),
                        value=list(region.ignore_columns),
                        multiple=True,
                        label="Ignore columns",
                        on_change=_on_ignore_change,
                    ).props("outlined dense use-chips").classes("w-56")

                    if region.ignore_columns:

                        def _on_ignore_reason_change(
                            event: events.ValueChangeEventArguments,
                            member_id=member_id,
                            sheet_name=sheet_name,
                            region_id=region.region_id,
                        ) -> None:
                            _update_region(
                                member_id,
                                sheet_name,
                                region_id,
                                ignore_columns_reason=str(event.value or ""),
                            )

                        ui.input(
                            "Reason",
                            value=region.ignore_columns_reason,
                            on_change=_on_ignore_reason_change,
                        ).props("outlined dense").classes("w-56")

                        def _on_ignore_expiry_change(
                            event: events.ValueChangeEventArguments,
                            member_id=member_id,
                            sheet_name=sheet_name,
                            region_id=region.region_id,
                        ) -> None:
                            _update_region(
                                member_id,
                                sheet_name,
                                region_id,
                                ignore_columns_expires_on=str(event.value or ""),
                            )

                        ui.input(
                            "Expires on",
                            value=region.ignore_columns_expires_on,
                            on_change=_on_ignore_expiry_change,
                        ).props("outlined dense type=date").classes("w-40")

                    def _on_expected_refresh_change(
                        event: events.ValueChangeEventArguments,
                        member_id=member_id,
                        sheet_name=sheet_name,
                        region_id=region.region_id,
                    ) -> None:
                        cols = tuple(event.value or ())
                        _apply_region_transform(
                            member_id,
                            sheet_name,
                            region_id,
                            lambda r: set_expected_refresh_columns(r, cols),
                        )

                    ui.select(
                        list(region.available_columns),
                        value=list(region.expected_refresh_columns),
                        multiple=True,
                        label="Expected-refresh columns",
                        on_change=_on_expected_refresh_change,
                    ).props("outlined dense use-chips").classes("w-56")

        def render_preview_section() -> None:
            """Current-first preview with an explicit baseline toggle (Step
            7 criterion): loads only the side/window the analyst asked for,
            never both sides, never the whole sheet.
            """
            preview_box.clear()
            state = workspace_state["value"]
            with preview_box:
                if job.overall_status != "done" or job.result is None:
                    ui.label(
                        "Preview becomes available once setup analysis finishes."
                    ).classes("note")
                    return
                member_ids = sorted(job.result.members)
                if not member_ids:
                    return
                if preview_state["member_id"] not in member_ids:
                    preview_state["member_id"] = member_ids[0]
                active_member_id = preview_state["member_id"]
                member_profile = job.result.members[active_member_id]
                sheet_names = [sheet.sheet_name for sheet in member_profile.current_sheets]
                if not sheet_names:
                    ui.label("No sheets available to preview.").classes("note")
                    return
                if preview_state["sheet_name"] not in sheet_names:
                    preview_state["sheet_name"] = sheet_names[0]
                    preview_state["min_row"], preview_state["min_col"] = 1, 1
                active_sheet_name = preview_state["sheet_name"]

                with ui.row().classes("items-center gap-2 flex-wrap"):
                    if len(member_ids) > 1:

                        def _on_member_change(event: events.ValueChangeEventArguments) -> None:
                            preview_state["member_id"] = str(event.value)
                            preview_state["sheet_name"] = None
                            preview_state["outcome"] = None
                            render_preview_section()

                        ui.select(
                            member_ids,
                            value=active_member_id,
                            label="Member",
                            on_change=_on_member_change,
                        ).props("outlined dense").classes("w-36")

                    def _on_sheet_change(event: events.ValueChangeEventArguments) -> None:
                        preview_state["sheet_name"] = str(event.value)
                        preview_state["min_row"], preview_state["min_col"] = 1, 1
                        preview_state["outcome"] = None
                        render_preview_section()

                    ui.select(
                        sheet_names,
                        value=active_sheet_name,
                        label="Sheet",
                        on_change=_on_sheet_change,
                    ).props("outlined dense").classes("w-48")

                    if mode is QCRunMode.CYCLE_COMPARISON:

                        def _on_side_change(event: events.ValueChangeEventArguments) -> None:
                            workspace_state["value"] = dataclasses.replace(
                                workspace_state["value"], preview_side=str(event.value)
                            )
                            preview_state["outcome"] = None
                            render_preview_section()

                        ui.toggle(
                            {"current": "Current", "baseline": "Baseline"},
                            value=state.preview_side,
                            on_change=_on_side_change,
                        ).props("dense")

                    def _on_jump_change(event: events.ValueChangeEventArguments) -> None:
                        try:
                            row, col = parse_a1_cell(str(event.value))
                        except ValueError as exc:
                            ui.notify(str(exc), type="warning")
                            return
                        preview_state["min_row"], preview_state["min_col"] = row, col
                        preview_state["outcome"] = None
                        render_preview_section()

                    ui.input(
                        "Jump to cell",
                        value=f"{get_column_letter(preview_state['min_col'])}{preview_state['min_row']}",
                        on_change=_on_jump_change,
                    ).props("outlined dense").classes("w-32")

                    def _page(delta_rows: int) -> None:
                        preview_state["min_row"] = max(1, preview_state["min_row"] + delta_rows)
                        preview_state["outcome"] = None
                        render_preview_section()

                    ui.button(
                        icon="chevron_left", on_click=lambda: _page(-_PREVIEW_WINDOW_ROWS)
                    ).props("flat dense")
                    ui.button(
                        icon="chevron_right", on_click=lambda: _page(_PREVIEW_WINDOW_ROWS)
                    ).props("flat dense")

                    def _on_reveal_change(event: events.ValueChangeEventArguments) -> None:
                        preview_state["reveal_formulas"] = bool(event.value)
                        preview_state["outcome"] = None
                        render_preview_section()

                    ui.checkbox(
                        "Reveal formula text",
                        value=preview_state["reveal_formulas"],
                        on_change=_on_reveal_change,
                    )

                    ui.button(
                        "Load preview",
                        on_click=lambda: asyncio.create_task(_load_preview()),
                    ).props("no-caps flat dense")

                if preview_state["anchor_target"] is not None:
                    _, target_sheet, target_region_id = preview_state["anchor_target"]
                    ui.label(
                        f"Click a cell below to set the anchor for {target_region_id!r} "
                        f"on {target_sheet!r}. (Load preview first if the grid is empty.)"
                    ).classes("notecard")

                if preview_state["loading"]:
                    ui.spinner(size="1.2rem")
                    return
                if preview_state["error"]:
                    ui.label(preview_state["error"]).classes("notecard")
                    return
                outcome = preview_state["outcome"]
                if outcome is None:
                    ui.label(
                        "Click Load preview to fetch a bounded view of this sheet."
                    ).classes("note")
                    return
                if outcome.missing_credential:
                    ui.label("A password is required to preview this source.").classes(
                        "notecard"
                    )
                    return
                if outcome.disclosure:
                    ui.label(outcome.disclosure).classes("notecard")
                    return

                def _on_cell_click(row: int, col: int) -> None:
                    target = preview_state["anchor_target"]
                    if target is None:
                        return
                    target_member_id, target_sheet, target_region_id = target
                    preview_state["anchor_target"] = None
                    _apply_region_transform(
                        target_member_id,
                        target_sheet,
                        target_region_id,
                        lambda region: apply_anchor_click(region, row, col),
                    )

                _render_preview_grid(outcome, on_cell_click=_on_cell_click)

        async def _load_preview() -> None:
            active_member_id = preview_state["member_id"]
            active_sheet_name = preview_state["sheet_name"]
            if active_member_id is None or active_sheet_name is None:
                return
            baseline_path, current_path = members.get(active_member_id, (None, None))
            active_side = (
                workspace_state["value"].preview_side
                if mode is QCRunMode.CYCLE_COMPARISON
                else "current"
            )
            source_path_str = current_path if active_side == "current" else baseline_path
            if source_path_str is None:
                preview_state["error"] = "No file is available for that side."
                render_preview_section()
                return
            preview_state["loading"] = True
            preview_state["error"] = ""
            render_preview_section()
            try:
                source_path = Path(source_path_str)
                source_hash = await asyncio.to_thread(sha256_file, source_path)
                request = PreviewWindowRequest(
                    side=active_side,
                    path=str(source_path),
                    source_hash=source_hash,
                    sheet=active_sheet_name,
                    min_row=preview_state["min_row"],
                    min_col=preview_state["min_col"],
                    max_row=preview_state["min_row"] + _PREVIEW_WINDOW_ROWS - 1,
                    max_col=preview_state["min_col"] + _PREVIEW_WINDOW_COLS - 1,
                    reveal_formulas=preview_state["reveal_formulas"],
                )
                try:
                    outcome = await asyncio.to_thread(
                        run_exclusive,
                        work_dir,
                        f"setup-preview:{active_member_id}",
                        lambda: run_preview_window_worker(request),
                    )
                except QueueBusyError as exc:
                    preview_state["error"] = str(exc)
                    return
                preview_state["outcome"] = outcome
            finally:
                preview_state["loading"] = False
                render_preview_section()

        def render_warnings_section() -> None:
            warnings_box.clear()
            state = workspace_state["value"]
            with warnings_box:
                if job.overall_status != "done" or job.result is None:
                    return
                warnings = current_warnings()
                if not warnings:
                    return
                ui.label("Warnings").classes("runhead")
                for warning in warnings:
                    with ui.row().classes("items-center gap-2"):
                        ui.checkbox(
                            warning.message,
                            value=warning.code in state.warnings_acknowledged,
                            on_change=lambda event, code=warning.code: _toggle_warning(
                                code, bool(event.value)
                            ),
                        )

        def current_warnings():
            if job.result is None:
                return ()
            state = workspace_state["value"]
            warnings = compute_warnings(job.result) + compute_sheet_pairing_warnings(state)
            deck = state.deck_review
            warnings += compute_slide_pairing_warnings(deck)
            profile = selected_profile["value"]
            if profile is not None:
                warnings += compute_required_slide_warnings(
                    deck, tuple(profile.ppt.required_slides)
                )
                if mode is QCRunMode.FINAL_PACKAGE:
                    anchor_titles = tuple(
                        dict.fromkeys(m.slide for m in profile.crosscheck.mappings)
                    )
                    warnings += resolve_slide_anchors(deck, anchor_titles)
            return warnings

        async def do_run_once() -> None:
            await _finalize(save=False, run=True, as_new_name=None)

        async def do_save_profile() -> None:
            await _finalize(save=True, run=False, as_new_name=None)

        async def do_save_profile_and_run() -> None:
            await _finalize(save=True, run=True, as_new_name=None)

        async def do_export_configuration() -> None:
            profile = selected_profile["value"] or load_profile_by_name(
                profiles_dir, workspace_state["value"].profile_name
            )
            resolved = build_resolved_configuration(
                workspace_state["value"],
                profile=profile,
                profile_sha256=profile_sha256(profile),
                warnings_acknowledged=tuple(sorted(workspace_state["value"].warnings_acknowledged)),
                file_hashes=file_hashes,
            )
            import json
            import tempfile

            bundle = {
                "profile": profile.model_dump(mode="json", by_alias=True),
                "resolved_input_configuration": resolved.model_dump(mode="json"),
                "resolved_input_digest": resolved.canonical_sha256(),
            }
            with tempfile.NamedTemporaryFile(
                mode="w", suffix=".json", delete=False, encoding="utf-8"
            ) as handle:
                json.dump(bundle, handle, indent=2)
                export_path = handle.name
            ui.download(export_path, filename="qc-configuration.json")

        async def _finalize(*, save: bool, run: bool, as_new_name: str | None) -> None:
            state = workspace_state["value"]
            blockers = unresolved_blockers(state, current_warnings())
            if blockers:
                ui.notify("; ".join(blockers), type="warning")
                return
            try:
                profile = load_profile_by_name(profiles_dir, state.profile_name)
            except Exception as exc:
                ui.notify(f"Profile failed to load: {exc}", type="negative")
                return
            if save:
                contract = build_input_contract(state)
                target_name = as_new_name or (
                    state.profile_name if state.profile_name != "default" else ""
                )
                if not target_name:
                    ui.notify(
                        "Choose a profile name before saving (the built-in "
                        "default profile cannot be overwritten).",
                        type="warning",
                    )
                    return
                if target_name == state.profile_name:
                    # Updating the same profile this session opened --
                    # Step 11's "optimistic conflict protection" criterion:
                    # refuse rather than silently clobber a concurrent edit
                    # from another tab/session (mirrors the standalone
                    # profile editor's own `save_draft` staleness check).
                    current_hash = _profile_hash_or_none(profiles_dir, target_name)
                    if current_hash != state.profile_opened_hash:
                        ui.notify(
                            f"Profile {target_name!r} changed elsewhere "
                            "while this workspace was open -- reopen it to "
                            "see the latest version before saving.",
                            type="negative",
                        )
                        return
                elif profile_path(profiles_dir, target_name).exists():
                    ui.notify(
                        f"Profile {target_name!r} already exists -- choose "
                        "a different name.",
                        type="negative",
                    )
                    return
                saved_profile = profile.model_copy(
                    update={"name": target_name, "input_contract": contract}
                )
                try:
                    save_profile(saved_profile, profile_path(profiles_dir, target_name))
                except Exception as exc:
                    ui.notify(f"Profile could not be saved: {exc}", type="negative")
                    return
                profile = saved_profile
                workspace_state["value"] = dataclasses.replace(
                    state,
                    profile_name=target_name,
                    profile_opened_hash=_profile_hash_or_none(profiles_dir, target_name),
                )
                ui.notify(f"Profile {target_name!r} saved")
            if not run:
                return
            resolved = build_resolved_configuration(
                workspace_state["value"],
                profile=profile,
                profile_sha256=profile_sha256(profile),
                warnings_acknowledged=tuple(sorted(workspace_state["value"].warnings_acknowledged)),
                file_hashes=file_hashes,
            )
            try:
                manifest = PackageManifest.from_role_files(
                    {role: str(path) for role, path in files.items()}
                )
            except ValueError as exc:
                ui.notify(str(exc), type="negative")
                return
            from qc_tool.ui.app import SessionState, build_run_request

            live_state = workspace_state["value"]
            deck = live_state.deck_review
            available_slides = (
                [(s.slide_index, s.title) for s in deck.current_slides]
                if deck is not None
                else []
            )
            selected_slides = (
                {s.slide_index for s in deck.current_slides if s.included}
                if deck is not None
                else set()
            )
            run_state = SessionState(
                files=files,
                file_hashes=file_hashes,
                profile_name=profile.name,
                mode=mode,
                output_mode=_output_mode_from_choices(choices),
                allow_large_workbooks=live_state.allow_large_workbooks,
                allow_dependency_indexing=live_state.allow_dependency_indexing,
                acceptance_absolute=_coerce_float(choices.get("acceptance_absolute"), 0.0),
                acceptance_percent=_coerce_float(choices.get("acceptance_percent"), 0.0),
                rerun_of=_coerce_optional_int(choices.get("rerun_of")),
                available_slides=available_slides,
                selected_slides=selected_slides,
            )
            request = build_run_request(
                work_dir=work_dir,
                state=run_state,
                files=files,
                profile=profile,
                manifest=manifest,
                resolved_input_configuration=resolved.model_dump(mode="json"),
                resolved_input_digest=resolved.canonical_sha256(),
            )
            try:
                queue_manager.submit(request, {})
            except QueueBusyError as exc:
                ui.notify(str(exc), type="warning")
                return
            ui.notify("Run started")
            ui.navigate.to("/")

        def _set_allow_large_workbooks(value: bool) -> None:
            workspace_state["value"] = dataclasses.replace(
                workspace_state["value"], allow_large_workbooks=value
            )
            _persist_choices()

        def _set_allow_dependency_indexing(value: bool) -> None:
            workspace_state["value"] = dataclasses.replace(
                workspace_state["value"], allow_dependency_indexing=value
            )
            _persist_choices()

        def render_advanced_section() -> None:
            """Compact tolerance/materiality/waiver + availability summary
            (Step 9) plus run-only workload/dependency safety overrides.
            The full profile editor remains the place to actually CHANGE
            tolerance/waiver/availability policy -- this is read-only.
            """
            advanced_box.clear()
            state = workspace_state["value"]
            with advanced_box:
                if job.overall_status == "done":
                    profile = selected_profile["value"]
                    if profile is not None:
                        with ui.expansion(
                            "Advanced controls & policy summary", value=False
                        ).classes("w-full"):
                            ui.label(
                                f"Tolerance: ±{profile.tolerance.absolute:g} "
                                f"absolute, ±{profile.tolerance.relative:g} "
                                "relative"
                            ).classes("note")
                            ui.label(
                                f"{len(profile.waivers)} waiver(s), "
                                f"{len(profile.materiality_severity)} materiality "
                                "severity override(s)"
                            ).classes("note")
                            ui.label(
                                f"{len(profile.ppt.availability_rules)} PowerPoint "
                                "availability rule(s)"
                            ).classes("note")
                            ui.link(
                                "Edit tolerance, waivers, and availability rules "
                                "in Manage profiles",
                                "/",
                            ).classes("guide-jump")
                with ui.row().classes("items-center gap-2"):
                    ui.checkbox(
                        "Allow large workbooks (override the workload safety gate)",
                        value=state.allow_large_workbooks,
                        on_change=lambda event: _set_allow_large_workbooks(bool(event.value)),
                    )
                with ui.row().classes("items-center gap-2"):
                    ui.checkbox(
                        "Force full dependency indexing (override the size gate)",
                        value=state.allow_dependency_indexing,
                        on_change=lambda event: _set_allow_dependency_indexing(bool(event.value)),
                    )

        def refresh() -> None:
            render_scan_status()
            render_profile_section()
            render_regions_section()
            render_deck_section()
            render_preview_section()
            render_warnings_section()
            render_advanced_section()
            actions_box.clear()
            with actions_box:
                blockers = unresolved_blockers(workspace_state["value"], current_warnings())
                enabled = job.overall_status == "done" and not blockers
                ui.button("Run once", on_click=do_run_once).classes("runbtn").props(
                    "no-caps"
                ).set_enabled(enabled)
                ui.button("Save profile", on_click=do_save_profile).classes("ghostbtn").props(
                    "no-caps flat"
                ).set_enabled(job.overall_status == "done" and not blockers)
                ui.button(
                    "Save profile and run", on_click=do_save_profile_and_run
                ).classes("ghostbtn").props("no-caps flat").set_enabled(enabled)
                ui.button(
                    "Update profile and run", on_click=do_save_profile_and_run
                ).classes("ghostbtn").props("no-caps flat").set_enabled(
                    enabled and workspace_state["value"].profile_name != "default"
                )
                ui.button(
                    "Export configuration", on_click=do_export_configuration
                ).classes("ghostbtn").props("no-caps flat").set_enabled(
                    job.overall_status == "done"
                )
                if blockers:
                    ui.label("; ".join(blockers)).classes("notecard")

        if scan_members and job.overall_status == "idle":
            passwords = {}  # browser-supplied credentials mid-scan remain out of scope
            start_setup_scan(work_dir, job, scan_members, passwords)

        def poll() -> None:
            seeded = False
            if job.overall_status == "done" and not workspace_state["value"].member_reviews:
                _seed_member_reviews()
                seeded = True
            if job.overall_status == "done" and workspace_state["value"].deck_review is None:
                _seed_deck_review()
                seeded = True
            if seeded or job.overall_status == "running":
                refresh()

        refresh()
        ui.timer(0.5, poll)


def _output_mode_from_choices(choices: dict[str, object]):
    from qc_tool.coverage import FindingOutputMode

    try:
        return FindingOutputMode(str(choices.get("output_mode", "decision")))
    except ValueError:
        return FindingOutputMode.DECISION
