"""Step 10: dependency-aware safeguards and phase telemetry."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from openpyxl import Workbook

from qc_tool.coverage import CoverageState
from qc_tool.engine import _workload_coverage, run_qc
from qc_tool.excel import complexity as complexity_module
from qc_tool.excel.complexity import (
    WorkbookComplexityError,
    assess_workbook_complexity,
)
from qc_tool.io.model import (
    CellRecord,
    ConditionalFormatDescriptor,
    SheetSnapshot,
    WorkbookSnapshot,
    WorkbookWorkload,
)
from qc_tool.progress import PhaseTelemetry, ProgressEvent, RunPhase


def _snapshot(
    cells: dict[tuple[int, int], CellRecord],
    *,
    rules: int = 0,
) -> WorkbookSnapshot:
    return WorkbookSnapshot(
        "synthetic.xlsx",
        "xlsx",
        True,
        True,
        formula_presence_available=True,
        formula_source="openpyxl",
        sheets=[SheetSnapshot("Data", "visible", 100, 10, cells)],
        conditional_formats=[
            ConditionalFormatDescriptor(
                sheet="Data",
                source_index=index,
                source_id=f"rule-{index}",
                target_ranges=("A1:A2",),
                rule_type="expression",
                formulas=("A1>0",),
                priority=index + 1,
                stop_if_true=False,
                style_key="synthetic",
            )
            for index in range(rules)
        ],
    )


def _cycle_workbook(path: Path, *, offset: int) -> None:
    workbook = Workbook()
    sheet = workbook.active
    if sheet is None:  # pragma: no cover - openpyxl always creates one sheet
        raise RuntimeError("openpyxl did not create a default worksheet")
    sheet.title = "Data"
    sheet.append(["Key", "Amount", "Derived"])
    for index in range(1, 6):
        sheet.append([f"k{index}", index + offset, "=SUM(B2:B6)"])
    workbook.save(path)


def test_step10_checkpoint_reconciles_with_the_delta_ledger() -> None:
    oracles = Path(__file__).parent / "oracles"
    step9 = json.loads((oracles / "real_workload_step9.json").read_text(encoding="utf-8"))
    step10 = json.loads(
        (oracles / "real_workload_step10.json").read_text(encoding="utf-8")
    )
    ledger = json.loads(
        (oracles / "step_delta_ledger.json").read_text(encoding="utf-8")
    )["steps"]["10"]

    assert step10["atomic_findings"] - step9["atomic_findings"] == (
        ledger["expected_atomic_delta"]
    )
    assert step10["severity"] == step9["severity"]
    assert step10["review_counts"] == step9["review_counts"]
    assert step10["pattern_review_counts"] == step9["pattern_review_counts"]
    assert (
        step10["coverage_states"]["checked"] - step9["coverage_states"]["checked"]
        == ledger["expected_checked_coverage_delta"]
    )
    assert step10["coverage_states"]["degraded"] == step9["coverage_states"]["degraded"]
    assert (
        step10["coverage_states"]["unavailable"]
        == step9["coverage_states"]["unavailable"]
    )
    assert set(ledger["expected_phases"]) <= set(step10["phases"])
    assert step10["source_hashes_unchanged"] is True


def test_complexity_reports_every_planned_cost_driver() -> None:
    cells = {
        (1, 1): CellRecord(1, 1, 1),
        (1, 2): CellRecord(1, 2, 1, formula="=SUM(A1:A50)"),
        (2, 2): CellRecord(2, 2, 1, formula="=LET(x,A1,x*2)"),
        (3, 2): CellRecord(3, 2, 1, formula="=A1"),
    }

    complexity = assess_workbook_complexity(_snapshot(cells, rules=3))

    assert complexity.formula_count == 3
    assert complexity.lexical_formula_count == 1
    assert complexity.reference_operands == 3
    assert complexity.resolved_range_cells == 52
    assert complexity.projected_concrete_edges == 49
    assert complexity.interaction_rule_count == 3
    assert complexity.warning_reasons == ()
    assert not complexity.degraded


def test_warning_thresholds_degrade_without_refusing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        complexity_module,
        "_COMPLEXITY_LIMITS",
        (("formula_count", 2, 100, "formula cells"),),
    )
    cells = {
        (row, 2): CellRecord(row, 2, 1, formula="=A1") for row in range(1, 5)
    }

    complexity = assess_workbook_complexity(_snapshot(cells))

    assert complexity.degraded
    assert not complexity.override_used
    assert complexity.warning_reasons == (
        "formula cells 4 >= warning limit 2",
    )


def test_refusal_limits_require_an_explicit_local_override(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        complexity_module,
        "_COMPLEXITY_LIMITS",
        (("formula_count", 2, 3, "formula cells"),),
    )
    cells = {
        (row, 2): CellRecord(row, 2, 1, formula="=A1") for row in range(1, 5)
    }
    workbook = _snapshot(cells)

    with pytest.raises(WorkbookComplexityError, match="allow_complex_workbook"):
        assess_workbook_complexity(workbook)

    overridden = assess_workbook_complexity(workbook, allow_complex_workbook=True)

    assert overridden.override_used
    assert overridden.degraded
    assert any("override accepted" in reason for reason in overridden.warning_reasons)


def test_progress_subdivides_every_planned_excel_phase(tmp_path: Path) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _cycle_workbook(baseline, offset=0)
    _cycle_workbook(current, offset=1)
    telemetry = PhaseTelemetry()
    before = {path: path.read_bytes() for path in (baseline, current)}

    run_qc(baseline_excel=baseline, current_excel=current, on_progress=telemetry)

    assert {
        RunPhase.PREPARING,
        RunPhase.LOADING_BASELINE_EXCEL,
        RunPhase.LOADING_CURRENT_EXCEL,
        RunPhase.DIFFING_EXCEL,
        RunPhase.COMPARING_FORMULAS,
        RunPhase.INDEXING_DEPENDENCIES,
        RunPhase.QUERYING_IMPACTS,
        RunPhase.BUILDING_REVIEW,
    } <= set(telemetry.records)
    assert all(path.read_bytes() == data for path, data in before.items())


def test_phase_telemetry_records_timings_units_and_sanitized_failures() -> None:
    telemetry = PhaseTelemetry()

    telemetry(ProgressEvent(RunPhase.INDEXING_DEPENDENCIES, total=10))
    telemetry(ProgressEvent(RunPhase.INDEXING_DEPENDENCIES, processed=10, total=10))
    telemetry.fail(RunPhase.INDEXING_DEPENDENCIES, "worker exited with code 1")

    payload = telemetry.as_payload()

    assert len(payload) == 1
    record = payload[0]
    assert record["phase"] == "indexing_dependencies"
    assert record["processed"] == 10
    assert record["total"] == 10
    assert record["started_at"] and record["finished_at"]
    assert isinstance(record["elapsed_seconds"], float)
    assert record["error"] == "worker exited with code 1"


def test_dependency_workload_coverage_is_reported_for_a_cycle_run(
    tmp_path: Path,
) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _cycle_workbook(baseline, offset=0)
    _cycle_workbook(current, offset=1)

    result = run_qc(baseline_excel=baseline, current_excel=current)
    item = next(
        item
        for item in result.coverage
        if item.check_id == "excel-dependency-workload"
    )

    assert item.state is CoverageState.CHECKED
    assert "projected cell dependencies" in item.detail


def test_xlsb_physical_workload_coverage_is_reported_when_scanned() -> None:
    workbook = WorkbookSnapshot(
        source_name="synthetic.xlsb",
        file_format="xlsb",
        formulas_available=False,
        styles_available=False,
        workload=WorkbookWorkload(
            format="xlsb",
            cell_count=10,
            formula_count=3,
            worksheet_binary_bytes=1_024,
            sheet_count=1,
            largest_sheet_area=20,
        ),
    )

    item = _workload_coverage(workbook)

    assert item.state is CoverageState.CHECKED
    assert "BIFF12 worksheets" in item.detail


def test_unavailable_xlsb_metrics_remain_explicit() -> None:
    workbook = WorkbookSnapshot(
        source_name="synthetic.xlsb",
        file_format="xlsb",
        formulas_available=False,
        styles_available=False,
        workload=WorkbookWorkload(format="xlsb", metrics_available=False),
    )

    item = _workload_coverage(workbook)

    assert item.state is CoverageState.UNAVAILABLE
    assert item.detail == "Workbook workload metrics are unavailable"
