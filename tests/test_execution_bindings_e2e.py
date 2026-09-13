"""End-to-end proof that engine-native logical bindings drive a real
``run_qc()`` cycle comparison (plan-20260913, Step 3): a confirmed sheet
rename plus an execution-confirmed keyed region.

Deliberately exercises the full pipeline (load -> align -> structure ->
formulas -> values) rather than calling individual producers directly, so
this is real proof the private ``_resolved_input_configuration`` parameter
is wired correctly end to end, not just at the ``align_workbooks()`` layer.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from qc_tool.config.resolved_input import (
    ResolvedColumn,
    ResolvedInputConfigurationV1,
    ResolvedMember,
    ResolvedRegion,
    ResolvedSheet,
)
from qc_tool.coverage import QCRunMode
from qc_tool.engine import run_qc
from qc_tool.findings import FindingClass


def _write_pair(tmp_path: Path) -> tuple[Path, Path]:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"

    base_wb = Workbook()
    base_ws = base_wb.active
    assert base_ws is not None
    base_ws.title = "Ledger2025"
    base_ws.append(["ID", "Amount"])
    base_ws.append(["A1", 10])
    base_ws.append(["A2", 20])
    base_ws.append(["A3", 30])
    base_wb.save(baseline_path)

    curr_wb = Workbook()
    curr_ws = curr_wb.active
    assert curr_ws is not None
    curr_ws.title = "Ledger2026"
    curr_ws.append(["ID", "Amount"])
    curr_ws.append(["A3", 30])
    curr_ws.append(["A1", 10])
    curr_ws.append(["A2", 25])  # genuine value change
    curr_wb.save(current_path)
    return baseline_path, current_path


def _resolved_configuration() -> ResolvedInputConfigurationV1:
    return ResolvedInputConfigurationV1(
        mode=QCRunMode.CYCLE_COMPARISON,
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(
                    ResolvedSheet(
                        sheet_id="ledger",
                        baseline_sheet_name="Ledger2025",
                        current_sheet_name="Ledger2026",
                        regions=(
                            ResolvedRegion(
                                region_id="r1",
                                mode="keyed",
                                header_intent="first_data_row",
                                current_data_range="A2:B4",
                                current_first_data_row=2,
                                columns=(
                                    ResolvedColumn(
                                        column_id="id",
                                        current_letter="A",
                                        alignment_role="identity",
                                    ),
                                ),
                            ),
                        ),
                    ),
                ),
            ),
        ),
    )


def test_run_qc_with_resolved_configuration_reports_rename_and_real_change(
    tmp_path: Path,
) -> None:
    baseline_path, current_path = _write_pair(tmp_path)

    result = run_qc(
        baseline_excel=baseline_path,
        current_excel=current_path,
        mode=QCRunMode.CYCLE_COMPARISON,
        _resolved_input_configuration=_resolved_configuration(),
    )

    classes = [f.finding_class for f in result.findings]
    assert classes.count(FindingClass.SHEET_RENAMED) == 1
    assert classes.count(FindingClass.SHEET_ADDED) == 0
    assert classes.count(FindingClass.SHEET_REMOVED) == 0
    # Pure reorder produces no row insertion/deletion noise.
    assert classes.count(FindingClass.ROW_INSERTED) == 0
    assert classes.count(FindingClass.ROW_DELETED) == 0
    # The genuine value edit (A2: 20 -> 25) still reports.
    value_changes = [f for f in result.findings if f.finding_class is FindingClass.VALUE_CHANGED]
    assert len(value_changes) == 1
    assert value_changes[0].sheet == "Ledger2026"


def test_structural_and_formula_findings_carry_a_bounded_logical_address(
    tmp_path: Path,
) -> None:
    baseline_path, current_path = _write_pair(tmp_path)

    result = run_qc(
        baseline_excel=baseline_path,
        current_excel=current_path,
        mode=QCRunMode.CYCLE_COMPARISON,
        _resolved_input_configuration=_resolved_configuration(),
    )

    renamed = next(
        f for f in result.findings if f.finding_class is FindingClass.SHEET_RENAMED
    )
    assert renamed.logical_address is not None
    assert renamed.logical_address.member_id == "primary"
    assert renamed.logical_address.sheet_id == "ledger"
    # Never a sheet name, range, or value -- content-free ids only.
    assert "Ledger2025" not in renamed.logical_address.model_dump_json()
    assert "Ledger2026" not in renamed.logical_address.model_dump_json()
    # The private sidecar never reaches the public payload.
    assert "logical_address" not in renamed.model_dump(mode="json")


def test_run_qc_without_resolved_configuration_stays_legacy(tmp_path: Path) -> None:
    """Guard: omitting ``_resolved_input_configuration`` (the default) keeps
    a physical rename as ordinary sheet_removed + sheet_added.
    """
    baseline_path, current_path = _write_pair(tmp_path)

    result = run_qc(
        baseline_excel=baseline_path,
        current_excel=current_path,
        mode=QCRunMode.CYCLE_COMPARISON,
    )

    classes = [f.finding_class for f in result.findings]
    assert FindingClass.SHEET_RENAMED not in classes
    assert classes.count(FindingClass.SHEET_ADDED) == 1
    assert classes.count(FindingClass.SHEET_REMOVED) == 1
