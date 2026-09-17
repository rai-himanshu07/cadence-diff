"""Complete, lossless reporting-contract editor behavior."""

import asyncio
import threading
from pathlib import Path
from typing import Any

import pytest

import qc_tool.ui.profile_editor as profile_editor_module
from qc_tool.config.editor import (
    PROFILE_SCHEMA,
    ProfileDraft,
    default_for_schema,
    profile_to_yaml,
    resolve_schema,
    save_draft,
    schema_leaf_paths,
    yaml_to_profile,
)
from qc_tool.config.profile import DeliverableProfile, load_profile, save_profile
from qc_tool.progress import CancellationToken, RunCancelled
from qc_tool.ui.profile_editor import ValidationTask, editor_leaf_paths


def _complete_profile(name: str = "contract") -> DeliverableProfile:
    return DeliverableProfile.model_validate(
        {
            "name": name,
            "description": "Quarterly board pack",
            "tolerance": {"absolute": 0.5, "relative": 0.01},
            "restatement_windows": {"week": 4, "month": 3, "quarter": 2},
            "excel": {
                "ignore_sheets": ["Scratch"],
                "sheets": {
                    "Dashboard": {
                        "ignore": False,
                        "ignore_ranges": ["Z:Z"],
                        "refresh_ranges": ["B2:B5"],
                        "acceptance_bands": [
                            {"range": "C2:C5", "absolute": 1, "relative": 0.02}
                        ],
                        "regions": [
                            {
                                "range": "A1:G20",
                                "orientation": "wide",
                                "header_row": 2,
                                "key_column": "A",
                            }
                        ],
                        "cadence_bands": [
                            {"range": "B1:G1", "kind": "month"}
                        ],
                        "availability_rules": [
                            {
                                "name": "Forecast horizon",
                                "range": "B5:G5",
                                "periods": "B1:G1",
                                "required_through": "Jun 2026",
                                "allow_blank_after": True,
                                "role": "forecast",
                            }
                        ],
                        "chart_windows": {"Revenue": "rolling"},
                    }
                },
                "controls": {
                    "required_ranges": [
                        {"name": "Revenue", "sheet": "Dashboard", "range": "B2"}
                    ],
                    "unique_ranges": [
                        {
                            "name": "IDs",
                            "sheet": "Data",
                            "range": "A:A",
                            "skip_header": True,
                        }
                    ],
                    "numeric_bounds": [
                        {
                            "name": "Margin",
                            "sheet": "Dashboard",
                            "range": "B4",
                            "minimum": 0,
                            "maximum": 1,
                        }
                    ],
                    "tie_outs": [
                        {
                            "name": "Revenue signed tie-out",
                            "target": "Dashboard!B2",
                            "components": [],
                            "terms": [
                                {"reference": "Data!B4", "operation": "subtract"}
                            ],
                            "absolute_tolerance": 0.1,
                            "relative_tolerance": 0.01,
                        },
                        {
                            "name": "Profit tie-out",
                            "target": "Dashboard!B3",
                            "components": ["Data!B5", "Data!B6"],
                            "terms": [],
                            "absolute_tolerance": 0,
                            "relative_tolerance": 0,
                        },
                    ],
                },
            },
            "ppt": {
                "slide_pins": {"Old summary": "Executive summary"},
                "match_threshold": 70,
                "chart_windows": {"Executive summary": "full"},
                "availability_rules": [
                    {
                        "name": "Forecast chart",
                        "slide": "Executive summary",
                        "scope": "chart",
                        "element": "Revenue",
                        "series": "Forecast",
                        "required_through": "Jun 2026",
                        "allow_blank_after": True,
                        "role": "forecast",
                    }
                ],
                "required_slides": ["Executive summary"],
                "draft_tokens": ["TBD", "DRAFT"],
            },
            "crosscheck": {
                "mappings": [
                    {
                        "slide": "Executive summary",
                        "line_skeleton": "Revenue $#M",
                        "figure_index": 0,
                        "label": "Revenue",
                        "source_sheet": "Dashboard",
                        "source_cell": "B2",
                    }
                ],
                "max_candidates": 7,
            },
            "waivers": [
                {
                    "finding_class": "style_changed",
                    "reason": "Approved branding",
                    "expires": "2026-12-31",
                    "sheet": "Dashboard",
                    "slide": None,
                    "location": "B2",
                    "element": "cell",
                }
            ],
            "severity": {"value_changed": "warning"},
            "materiality_severity": {"material": "critical"},
        }
    )


def _schema_at(*parts: str) -> dict[str, object]:
    schema: dict[str, object] = PROFILE_SCHEMA
    for part in parts:
        resolved, _ = resolve_schema(schema)
        if part == "[]":
            schema = resolved["items"]  # type: ignore[assignment]
        elif part == "*":
            schema = resolved["additionalProperties"]  # type: ignore[assignment]
        else:
            schema = resolved["properties"][part]  # type: ignore[index,assignment]
    return schema


def test_complete_profile_survives_form_yaml_form_roundtrip() -> None:
    profile = _complete_profile()
    draft = ProfileDraft.from_profile(profile)

    draft.apply_yaml(draft.yaml())

    assert draft.profile() == profile
    assert yaml_to_profile(profile_to_yaml(profile)) == profile
    paths = schema_leaf_paths()
    assert editor_leaf_paths() == paths
    assert {
        ("excel", "sheets", "*", "availability_rules", "[]", "role"),
        ("excel", "controls", "tie_outs", "[]", "terms", "[]", "operation"),
        ("ppt", "availability_rules", "[]", "series"),
        ("crosscheck", "mappings", "[]", "source_cell"),
        ("waivers", "[]", "expires"),
        ("severity", "*"),
        ("materiality_severity", "*"),
    } <= paths


def test_nested_repeatable_mutations_preserve_unrelated_advanced_fields() -> None:
    draft = ProfileDraft.from_profile(_complete_profile())
    required_schema = _schema_at(
        "excel", "controls", "required_ranges", "[]"
    )

    draft.append(("excel", "controls", "required_ranges"), required_schema)
    draft.set(("excel", "controls", "required_ranges", 1, "name"), "Profit")
    draft.set(("excel", "controls", "required_ranges", 1, "sheet"), "Dashboard")
    draft.set(("excel", "controls", "required_ranges", 1, "range"), "B3")
    draft.move(("excel", "controls", "required_ranges"), 1, -1)
    draft.remove(("ppt", "draft_tokens"), 0)
    draft.rename_mapping_key(
        ("ppt", "slide_pins"), "Old summary", "Prior summary"
    )
    draft.set_mapping(
        ("excel", "sheets", "Dashboard", "chart_windows"),
        "Profit",
        default_for_schema(
            _schema_at("excel", "sheets", "*", "chart_windows", "*")
        ),
    )

    profile = draft.profile()
    assert profile.excel.controls.required_ranges[0].name == "Profit"
    assert profile.ppt.draft_tokens == ["DRAFT"]
    assert profile.ppt.slide_pins == {"Prior summary": "Executive summary"}
    assert profile.excel.sheets["Dashboard"].chart_windows["Profit"] == "rolling"
    assert profile.excel.sheets["Dashboard"].availability_rules[0].role == "forecast"
    assert draft.dirty


def test_append_materializes_an_omitted_optional_array() -> None:
    draft = ProfileDraft.from_profile(DeliverableProfile(name="optional-array"))
    path = ("excel", "comparison_prerequisites")
    item_schema = _schema_at("excel", "comparison_prerequisites", "[]")

    assert "comparison_prerequisites" not in draft.payload["excel"]
    draft.append(path, item_schema)
    draft.set((*path, 0, "name"), "Forecast scenario")
    draft.set((*path, 0, "sheet"), "Dashboard")
    draft.set((*path, 0, "cell"), "B3")

    [prerequisite] = draft.profile().excel.comparison_prerequisites
    assert prerequisite.name == "Forecast scenario"
    assert prerequisite.sheet == "Dashboard"
    assert prerequisite.cell == "B3"


def test_default_is_immutable_and_named_profile_can_be_renamed(tmp_path: Path) -> None:
    default = ProfileDraft.from_profile(DeliverableProfile(name="default"))
    with pytest.raises(ValueError, match="default profile is immutable"):
        save_draft(default, tmp_path)
    assert not list(tmp_path.iterdir())

    original_path = tmp_path / "old-name.yaml"
    save_profile(DeliverableProfile(name="old-name"), original_path)
    draft = ProfileDraft.from_path(original_path)
    draft.set(("name",), "new-name")

    saved = save_draft(draft, tmp_path)

    assert saved.name == "new-name"
    assert not original_path.exists()
    assert load_profile(tmp_path / "new-name.yaml") == saved
    assert draft.selected_name == "new-name"


def test_stale_or_deleted_source_is_refused_without_overwrite(tmp_path: Path) -> None:
    path = tmp_path / "contract.yaml"
    original = _complete_profile()
    save_profile(original, path)
    stale = ProfileDraft.from_path(path)
    path.write_text(
        f"# external formatting edit\n{path.read_text(encoding='utf-8')}",
        encoding="utf-8",
    )

    stale.set(("description",), "analyst edit")
    with pytest.raises(ValueError, match="changed while this editor was open"):
        save_draft(stale, tmp_path)
    assert load_profile(path) == original

    fresh = ProfileDraft.from_path(path)
    replacement = original.model_copy(update={"description": "external edit"})
    save_profile(replacement, path)

    fresh.set(("description",), "analyst edit")
    with pytest.raises(ValueError, match="changed while this editor was open"):
        save_draft(fresh, tmp_path)
    assert load_profile(path) == replacement

    fresh = ProfileDraft.from_path(path)
    path.unlink()
    with pytest.raises(ValueError, match="changed while this editor was open"):
        save_draft(fresh, tmp_path)
    assert not path.exists()


def test_invalid_draft_cannot_corrupt_disk_and_valid_save_clears_dirty(
    tmp_path: Path,
) -> None:
    path = tmp_path / "contract.yaml"
    profile = _complete_profile()
    save_profile(profile, path)
    draft = ProfileDraft.from_path(path)
    draft.set(("restatement_windows", "week"), -1)

    with pytest.raises(ValueError, match="greater than or equal to 0"):
        save_draft(draft, tmp_path)
    assert load_profile(path) == profile

    draft.set(("restatement_windows", "week"), 5)
    saved = save_draft(draft, tmp_path)
    assert saved.restatement_windows.week == 5
    assert load_profile(path) == saved
    assert not draft.dirty


@pytest.mark.asyncio
async def test_file_validation_prefers_current_files_and_forwards_passwords(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls: list[tuple[str, Path, str | None, bool]] = []

    def load_workbook(
        path: Path,
        *,
        password: str | None,
        allow_large_workbook: bool,
        cancellation_token: CancellationToken,
    ) -> object:
        calls.append(("excel", path, password, allow_large_workbook))
        cancellation_token.check()
        return object()

    def load_deck(
        path: Path,
        *,
        password: str | None,
        cancellation_token: CancellationToken,
    ) -> object:
        calls.append(("ppt", path, password, False))
        cancellation_token.check()
        return object()

    monkeypatch.setattr(profile_editor_module, "load_workbook_snapshot", load_workbook)
    monkeypatch.setattr(profile_editor_module, "load_deck_snapshot", load_deck)
    monkeypatch.setattr(profile_editor_module, "lint_profile", lambda *args, **kwargs: [])
    completed = asyncio.Event()
    result: list[list[Any] | Exception] = []
    task = ValidationTask()

    def complete(value: list[Any] | Exception) -> None:
        result.append(value)
        completed.set()

    task.start(
        DeliverableProfile(name="contract"),
        {
            "baseline_excel": tmp_path / "baseline.xlsx",
            "current_excel": tmp_path / "current.xlsx",
            "baseline_ppt": tmp_path / "baseline.pptx",
            "current_ppt": tmp_path / "current.pptx",
        },
        {"current_excel": "excel-secret", "current_ppt": "ppt-secret"},
        lambda _message: None,
        complete,
    )
    await asyncio.wait_for(completed.wait(), timeout=1)

    assert result == [[]]
    assert calls == [
        ("excel", tmp_path / "current.xlsx", "excel-secret", False),
        ("ppt", tmp_path / "current.pptx", "ppt-secret", False),
    ]


@pytest.mark.asyncio
async def test_file_validation_cancels_through_the_loader_token(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    entered = threading.Event()
    release = threading.Event()
    completed = asyncio.Event()
    result: list[list[Any] | Exception] = []

    def load_workbook(
        path: Path,
        *,
        password: str | None,
        allow_large_workbook: bool,
        cancellation_token: CancellationToken,
    ) -> object:
        del path, password, allow_large_workbook
        entered.set()
        release.wait(timeout=1)
        cancellation_token.check()
        return object()

    monkeypatch.setattr(profile_editor_module, "load_workbook_snapshot", load_workbook)
    task = ValidationTask()

    def complete(value: list[Any] | Exception) -> None:
        result.append(value)
        completed.set()

    task.start(
        DeliverableProfile(name="contract"),
        {"current_excel": tmp_path / "current.xlsx"},
        {},
        lambda _message: None,
        complete,
    )
    assert await asyncio.to_thread(entered.wait, 1)
    assert task.cancel()
    release.set()
    await asyncio.wait_for(completed.wait(), timeout=1)

    assert len(result) == 1
    assert isinstance(result[0], RunCancelled)
