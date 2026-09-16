"""Step 5: cycle-comparison prerequisites (Acceptance Criterion 7).

A configured selector/scenario cell must be present, non-blank, and exactly
equal between baseline and current before any alignment, diffing, reports,
or history proceed. A mismatch raises one bounded, value-free
`RunBlockedError` instead of producing noisy or misleading findings.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from qc_tool.config.profile import (
    ComparisonPrerequisite,
    DeliverableProfile,
    ExcelMemberProfile,
    ExcelProfile,
    default_profile,
)
from qc_tool.coverage import QCRunMode
from qc_tool.engine import run_qc
from qc_tool.excel.prerequisites import evaluate_resolved_selectors
from qc_tool.io.model import CellRecord, SheetSnapshot, WorkbookSnapshot
from qc_tool.package import PackageArtifact, PackageManifest, PackageMember, PackageSide
from qc_tool.run_action import RunActionReason, RunBlockedError


def _book(path: Path, scenario: str | None) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Config"
    sheet["A1"] = "Scenario"
    sheet["B2"] = scenario
    workbook.save(path)


def _profile_with_prerequisite() -> DeliverableProfile:
    return DeliverableProfile(
        name="prereq",
        excel=ExcelProfile(
            comparison_prerequisites=[
                ComparisonPrerequisite(name="Scenario", sheet="Config", cell="B2")
            ]
        ),
    )


def test_matching_prerequisite_does_not_block(tmp_path: Path) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _book(baseline, "Base Case")
    _book(current, "Base Case")

    result = run_qc(
        baseline_excel=baseline,
        current_excel=current,
        profile=_profile_with_prerequisite(),
        mode=QCRunMode.CYCLE_COMPARISON,
    )

    assert result is not None


def test_mismatched_prerequisite_blocks_before_analysis(tmp_path: Path) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _book(baseline, "Base Case")
    _book(current, "Upside Case")

    with pytest.raises(RunBlockedError) as excinfo:
        run_qc(
            baseline_excel=baseline,
            current_excel=current,
            profile=_profile_with_prerequisite(),
            mode=QCRunMode.CYCLE_COMPARISON,
        )

    action = excinfo.value.action_required
    assert action.reason is RunActionReason.COMPARISON_PREREQUISITE_MISMATCH
    assert [item.sheet for item in action.items] == ["Config"]
    assert [item.cell for item in action.items] == ["B2"]
    assert action.items[0].label == "Scenario"
    assert "recalculate" in action.message.lower()
    # The bounded payload never carries the actual cell values.
    serialized = action.model_dump_json()
    assert "Base Case" not in serialized
    assert "Upside Case" not in serialized


def test_blank_prerequisite_cell_blocks(tmp_path: Path) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _book(baseline, "   ")  # whitespace-only: a real stored cell, still blank
    _book(current, "Base Case")

    with pytest.raises(RunBlockedError) as excinfo:
        run_qc(
            baseline_excel=baseline,
            current_excel=current,
            profile=_profile_with_prerequisite(),
            mode=QCRunMode.CYCLE_COMPARISON,
        )

    assert "blank" in excinfo.value.action_required.items[0].detail


def test_missing_sheet_blocks(tmp_path: Path) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    workbook = Workbook()
    active = workbook.active
    assert active is not None
    active.title = "Other"
    workbook.save(baseline)
    _book(current, "Base Case")

    with pytest.raises(RunBlockedError) as excinfo:
        run_qc(
            baseline_excel=baseline,
            current_excel=current,
            profile=_profile_with_prerequisite(),
            mode=QCRunMode.CYCLE_COMPARISON,
        )

    assert "missing" in excinfo.value.action_required.items[0].detail


def test_no_prerequisites_configured_never_blocks(tmp_path: Path) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _book(baseline, "Base Case")
    _book(current, "Upside Case")  # would mismatch if a prerequisite existed

    result = run_qc(
        baseline_excel=baseline,
        current_excel=current,
        profile=default_profile(),
        mode=QCRunMode.CYCLE_COMPARISON,
    )

    assert result is not None


def _package_member(
    member_id: str, side: PackageSide, display_name: str
) -> PackageMember:
    return PackageMember(
        member_id=member_id,
        side=side,
        artifact=PackageArtifact.EXCEL,
        display_name=display_name,
    )


def test_multi_member_package_attributes_the_blocked_member(tmp_path: Path) -> None:
    baseline_primary = tmp_path / "baseline_primary.xlsx"
    current_primary = tmp_path / "current_primary.xlsx"
    baseline_ops = tmp_path / "baseline_ops.xlsx"
    current_ops = tmp_path / "current_ops.xlsx"
    _book(baseline_primary, "Base Case")
    _book(current_primary, "Base Case")
    _book(baseline_ops, "Base Case")
    _book(current_ops, "Upside Case")  # the mismatch lives on member "ops"

    prerequisite = [
        ComparisonPrerequisite(name="Scenario", sheet="Config", cell="B2")
    ]
    profile = DeliverableProfile(
        name="prereq-pkg",
        excel=ExcelProfile(
            members={
                "primary": ExcelMemberProfile(comparison_prerequisites=prerequisite),
                "ops": ExcelMemberProfile(comparison_prerequisites=prerequisite),
            }
        ),
    )
    manifest = PackageManifest(
        members=(
            _package_member("primary", PackageSide.BASELINE, "baseline_primary.xlsx"),
            _package_member("primary", PackageSide.CURRENT, "current_primary.xlsx"),
            _package_member("ops", PackageSide.BASELINE, "baseline_ops.xlsx"),
            _package_member("ops", PackageSide.CURRENT, "current_ops.xlsx"),
        )
    )
    files = {
        "baseline_excel": baseline_primary,
        "current_excel": current_primary,
        "baseline_excel:ops": baseline_ops,
        "current_excel:ops": current_ops,
    }

    with pytest.raises(RunBlockedError) as excinfo:
        run_qc(
            package_manifest=manifest,
            package_files=files,
            profile=profile,
            mode=QCRunMode.CYCLE_COMPARISON,
        )

    assert excinfo.value.action_required.items[0].member_id == "ops"


def test_perform_run_writes_no_history_or_reports_when_blocked(tmp_path: Path) -> None:
    from qc_tool.history.store import RunHistory
    from qc_tool.run_service import perform_run

    work_dir = tmp_path / "work"
    work_dir.mkdir()
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _book(baseline, "Base Case")
    _book(current, "Upside Case")

    with pytest.raises(RunBlockedError):
        perform_run(
            work_dir,
            {"baseline_excel": baseline, "current_excel": current},
            {},
            _profile_with_prerequisite(),
            mode=QCRunMode.CYCLE_COMPARISON,
        )

    history_db = work_dir / "history.sqlite3"
    if history_db.exists():
        assert RunHistory(history_db).list_runs() == []
    runs_dir = work_dir / "runs"
    assert not runs_dir.exists() or not any(runs_dir.iterdir())


def _resolved_configuration_with_selector(
    member_id: str = "primary",
):
    from qc_tool.config.resolved_input import (
        ResolvedInputConfigurationV1,
        ResolvedMember,
        ResolvedSelector,
        ResolvedSheet,
    )

    return ResolvedInputConfigurationV1(
        mode=QCRunMode.CYCLE_COMPARISON,
        members=(
            ResolvedMember(
                member_id=member_id,
                sheets=(
                    ResolvedSheet(
                        sheet_id="config",
                        baseline_sheet_name="Config",
                        current_sheet_name="Config",
                        selectors=(
                            ResolvedSelector(
                                selector_id="scenario",
                                baseline_cell="B2",
                                current_cell="B2",
                            ),
                        ),
                    ),
                ),
            ),
        ),
    )


def test_matching_resolved_selector_does_not_block(tmp_path: Path) -> None:
    """Step 8: a selector prerequisite declared through the mode-aware
    configuration workspace (a ``ResolvedInputConfigurationV1``, not a
    legacy ``ComparisonPrerequisite``) gates the run identically.
    """
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _book(baseline, "Base Case")
    _book(current, "Base Case")

    result = run_qc(
        baseline_excel=baseline,
        current_excel=current,
        profile=default_profile(),
        mode=QCRunMode.CYCLE_COMPARISON,
        resolved_input_configuration=_resolved_configuration_with_selector(),
    )

    assert result.resolved_input_configuration is not None
    [member] = result.resolved_input_configuration.members
    [sheet] = member.sheets
    [selector] = sheet.selectors
    assert selector.equal is True
    assert selector.baseline_formula_backed is False
    assert selector.current_formula_backed is False
    assert result.resolved_input_digest == (
        result.resolved_input_configuration.canonical_sha256()
    )


def test_selector_outcome_records_formula_presence_without_values() -> None:
    def workbook(value: str, *, formula: bool) -> WorkbookSnapshot:
        cell = CellRecord(
            row=2,
            column=2,
            value=value,
            formula="=A1" if formula else None,
            is_formula=formula,
        )
        return WorkbookSnapshot(
            source_name="synthetic",
            file_format="xlsx",
            formulas_available=True,
            styles_available=True,
            sheets=[
                SheetSnapshot(
                    name="Config",
                    visibility="visible",
                    max_row=2,
                    max_column=2,
                    cells={(2, 2): cell},
                )
            ],
        )

    outcomes, mismatches = evaluate_resolved_selectors(
        workbook("Base Case", formula=True),
        workbook("Base Case", formula=False),
        _resolved_configuration_with_selector(),
    )

    assert mismatches == []
    [outcome] = outcomes
    assert outcome.equal is True
    assert outcome.baseline_formula_backed is True
    assert outcome.current_formula_backed is False
    assert "Base Case" not in repr(outcome)


def test_mismatched_resolved_selector_blocks_before_analysis(tmp_path: Path) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _book(baseline, "Base Case")
    _book(current, "Upside Case")

    with pytest.raises(RunBlockedError) as excinfo:
        run_qc(
            baseline_excel=baseline,
            current_excel=current,
            profile=default_profile(),
            mode=QCRunMode.CYCLE_COMPARISON,
            resolved_input_configuration=_resolved_configuration_with_selector(),
        )

    action = excinfo.value.action_required
    assert action.reason is RunActionReason.COMPARISON_PREREQUISITE_MISMATCH
    assert action.items[0].label == "scenario"
    serialized = action.model_dump_json()
    assert "Base Case" not in serialized
    assert "Upside Case" not in serialized


def test_multi_member_resolved_selector_is_enforced_for_its_owner(
    tmp_path: Path,
) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _book(baseline, "Base Case")
    _book(current, "Upside Case")
    manifest = PackageManifest(
        members=(
            _package_member("ops", PackageSide.BASELINE, "baseline.xlsx"),
            _package_member("ops", PackageSide.CURRENT, "current.xlsx"),
        )
    )

    with pytest.raises(RunBlockedError) as excinfo:
        run_qc(
            package_manifest=manifest,
            package_files={
                "baseline_excel:ops": baseline,
                "current_excel:ops": current,
            },
            profile=DeliverableProfile(
                name="package",
                excel=ExcelProfile(members={"ops": ExcelMemberProfile()}),
            ),
            mode=QCRunMode.CYCLE_COMPARISON,
            resolved_input_configuration=_resolved_configuration_with_selector("ops"),
        )

    assert (
        excinfo.value.action_required.reason
        is RunActionReason.COMPARISON_PREREQUISITE_MISMATCH
    )
    assert excinfo.value.action_required.items[0].member_id == "ops"
