"""Step 4: cycle provenance and accurate finding subtypes."""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Any

from openpyxl import Workbook

from qc_tool.config.profile import NumericTolerance
from qc_tool.engine import QCRunResult, compare_findings, run_qc
from qc_tool.excel.align import AxisAlignment, RegionAlignment, WorkbookAlignment
from qc_tool.excel.diff_values import _RangeSet, diff_region_values
from qc_tool.excel.formulas import diff_workbook_formulas
from qc_tool.excel.regions import TableRegion
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingEvidenceTag,
    FindingExpectedReason,
    FindingProvenance,
    FindingSubtype,
    FindingTemporalContext,
    Materiality,
    Severity,
)
from qc_tool.history.store import RunHistory
from qc_tool.io.model import CellRecord, SheetSnapshot, WorkbookSnapshot
from qc_tool.review import finding_identity_key
from qc_tool.triage.rules import triage

_ORACLE_DIR = Path(__file__).parent / "oracles"


def _oracle(name: str) -> dict[str, Any]:
    return json.loads((_ORACLE_DIR / name).read_text(encoding="utf-8"))


def _workbook(cells: dict[tuple[int, int], CellRecord], *, rows: int) -> WorkbookSnapshot:
    return WorkbookSnapshot(
        "synthetic.xlsx",
        "xlsx",
        True,
        True,
        formula_presence_available=True,
        formula_source="openpyxl",
        sheets=[SheetSnapshot("Synthetic", "visible", rows, 2, cells)],
    )


def _long_alignment(
    *,
    baseline_rows: int,
    current_rows: int,
    growth_rows: tuple[int, ...] = (),
) -> WorkbookAlignment:
    """One long region keyed on column A with formulas in column B."""
    baseline = TableRegion("Synthetic", 1, 1, baseline_rows, 2, "long", 1, 1, "rows")
    current = TableRegion("Synthetic", 1, 1, current_rows, 2, "long", 1, 1, "rows")
    return WorkbookAlignment(
        common_sheets=["Synthetic"],
        regions={
            "Synthetic": [
                RegionAlignment(
                    baseline,
                    current,
                    AxisAlignment(
                        pairs=[(row, row) for row in range(1, baseline_rows + 1)],
                        growth=list(growth_rows),
                    ),
                    AxisAlignment(pairs=[(1, 1), (2, 2)]),
                )
            ]
        },
    )


def _errors(findings: list[Finding]) -> dict[str, Finding]:
    return {
        finding.location or "": finding
        for finding in findings
        if finding.finding_class is FindingClass.FORMULA_ERROR
    }


def _outliers(findings: list[Finding]) -> dict[str, Finding]:
    return {
        finding.location or "": finding
        for finding in findings
        if finding.finding_class is FindingClass.FORMULA_INCONSISTENT
    }


def test_step4_scenarios_are_owned_by_this_slice() -> None:
    manifest = _oracle("step1_scenarios.json")
    owned = {
        scenario["id"]
        for scenario in manifest["scenarios"]
        if scenario["owner_step"] == 4
    }

    assert owned == {
        "formula-error-inherited",
        "formula-error-new",
        "consistency-historical",
        "consistency-novel",
        "value-added-population",
        "value-replacement",
    }


def test_step4_checkpoint_reconciles_with_the_delta_ledger() -> None:
    """The checkpoint restates the harness ``<finding_class>:<state>`` counters.

    ``real_workload_step4.json`` splits the harness ``finding_provenance`` and
    ``finding_subtypes`` maps into ``error_provenance`` (``formula_error:*``),
    ``outlier_provenance`` (``formula_inconsistent:*``), and ``value_subtypes``
    (``value_changed:*``); ``unset`` counters are dropped.
    """
    before = _oracle("real_workload_before_state.json")["workloads"]["cycle_pair"]
    step3 = _oracle("real_workload_step3.json")
    step4 = _oracle("real_workload_step4.json")
    ledger = _oracle("step_delta_ledger.json")["steps"]["4"]

    assert step4["atomic_findings"] - step3["atomic_findings"] == (
        ledger["expected_atomic_delta"]
    )
    assert step4["coverage_states"] == step3["coverage_states"]
    assert step4["severity"] == step3["severity"]
    assert step4["projected_concrete_edges"] == step3["projected_concrete_edges"]
    assert step4["source_hashes_unchanged"] is True
    assert {
        status: step4["reference_statuses"][status] - step3["reference_statuses"][status]
        for status in step3["reference_statuses"]
    } == ledger["expected_reference_status_delta"]
    assert sum(step4["review_counts"].values()) - sum(step3["review_counts"].values()) == (
        ledger["expected_review_count_delta"]
    )

    errors = step4["error_provenance"]
    outliers = step4["outlier_provenance"]
    subtypes = step4["value_subtypes"]
    assert errors == ledger["expected_error_provenance"]
    assert errors["inherited"] == before["formula_findings"]["errors"]
    assert outliers["new"] == ledger["expected_outlier_provenance"]["novel"]
    assert (
        outliers["inherited"] + outliers["historical_pattern"]
        == ledger["expected_outlier_provenance"]["historical"]
        == before["formula_findings"]["consistency"]
    )
    assert subtypes == ledger["expected_value_subtypes"]
    assert subtypes["added_population"] == before["value_findings"]["added_population"]
    assert sum(subtypes.values()) == before["value_findings"]["total"]


def test_error_provenance_separates_inherited_changed_and_new_states() -> None:
    baseline = _workbook(
        {
            (1, 1): CellRecord(1, 1, "k1"),
            (1, 2): CellRecord(1, 2, "#DIV/0!", formula="=1/0"),
            (2, 1): CellRecord(2, 1, "k2"),
            (2, 2): CellRecord(2, 2, "#N/A", formula="=NA()"),
            (3, 1): CellRecord(3, 1, "k3"),
            (3, 2): CellRecord(3, 2, 7, formula="=7"),
        },
        rows=3,
    )
    current = _workbook(
        {
            (1, 1): CellRecord(1, 1, "k1"),
            (1, 2): CellRecord(1, 2, "#DIV/0!", formula="=1/0"),
            (2, 1): CellRecord(2, 1, "k2"),
            (2, 2): CellRecord(2, 2, "#REF!", formula="=NA()"),
            (3, 1): CellRecord(3, 1, "k3"),
            (3, 2): CellRecord(3, 2, "#VALUE!", formula="=7"),
        },
        rows=3,
    )

    errors = _errors(
        diff_workbook_formulas(
            baseline,
            current,
            _long_alignment(baseline_rows=3, current_rows=3),
        )
    )

    assert errors["B1"].provenance is FindingProvenance.INHERITED
    assert errors["B2"].provenance is FindingProvenance.CHANGED
    assert errors["B3"].provenance is FindingProvenance.NEW
    assert "inherited from the baseline" in errors["B1"].message
    assert "new in this cycle" in errors["B3"].message
    assert all(
        finding.finding_class is FindingClass.FORMULA_ERROR for finding in errors.values()
    )


def test_error_literal_inside_formula_text_carries_its_own_provenance() -> None:
    cells = {
        (1, 1): CellRecord(1, 1, "k1"),
        (2, 1): CellRecord(2, 1, "k2"),
        (2, 2): CellRecord(2, 2, 0, formula="=SUM(#REF!)"),
    }
    baseline = _workbook({**cells, (1, 2): CellRecord(1, 2, 1, formula="=A1")}, rows=2)
    current = _workbook(
        {**cells, (1, 2): CellRecord(1, 2, 1, formula="=SUM(#REF!)")}, rows=2
    )

    errors = _errors(
        diff_workbook_formulas(
            baseline,
            current,
            _long_alignment(baseline_rows=2, current_rows=2),
        )
    )

    assert errors["B1"].provenance is FindingProvenance.NEW
    assert errors["B2"].provenance is FindingProvenance.INHERITED


def test_error_provenance_is_unset_without_a_confidently_aligned_baseline() -> None:
    baseline = _workbook({(1, 1): CellRecord(1, 1, "k1")}, rows=1)
    current = _workbook(
        {
            (1, 1): CellRecord(1, 1, "k1"),
            (5, 5): CellRecord(5, 5, "#REF!", formula="=1/0"),
        },
        rows=5,
    )

    errors = _errors(
        diff_workbook_formulas(
            baseline,
            current,
            _long_alignment(baseline_rows=1, current_rows=1),
        )
    )

    assert errors["E5"].provenance is None
    assert errors["E5"].message.endswith("error value #REF!")


def test_standalone_self_comparison_never_claims_cycle_provenance() -> None:
    workbook = _workbook(
        {
            (1, 1): CellRecord(1, 1, "k1"),
            (1, 2): CellRecord(1, 2, "#DIV/0!", formula="=1/0"),
        },
        rows=1,
    )

    findings = diff_workbook_formulas(
        workbook,
        workbook,
        _long_alignment(baseline_rows=1, current_rows=1),
        cycle=False,
    )

    assert [finding.provenance for finding in findings] == [None]


def test_consistency_provenance_splits_inherited_historical_and_novel_outliers() -> None:
    def row(index: int, formula: str) -> dict[tuple[int, int], CellRecord]:
        return {
            (index, 1): CellRecord(index, 1, f"k{index}"),
            (index, 2): CellRecord(index, 2, index, formula=formula),
        }

    baseline_cells: dict[tuple[int, int], CellRecord] = {}
    for index in range(1, 9):
        baseline_cells.update(row(index, "=A2*2" if index == 2 else f"=A{index}"))
    current_cells = dict(baseline_cells)
    current_cells.update(row(4, "=A4*3"))  # novel deviation on a paired cell
    current_cells.update(row(9, "=A9*2"))  # growth cell reusing a historical shape
    current_cells.update(row(10, "=A10/0"))  # growth cell with an unseen shape

    outliers = _outliers(
        diff_workbook_formulas(
            _workbook(baseline_cells, rows=8),
            _workbook(current_cells, rows=10),
            _long_alignment(baseline_rows=8, current_rows=10, growth_rows=(9, 10)),
        )
    )

    assert outliers["B2"].provenance is FindingProvenance.INHERITED
    assert outliers["B4"].provenance is FindingProvenance.NEW
    assert outliers["B9"].provenance is FindingProvenance.HISTORICAL_PATTERN
    assert outliers["B10"].provenance is FindingProvenance.NEW
    assert "already deviated this way" in outliers["B2"].message
    assert "already present in the historical range" in outliers["B9"].message
    assert "novel deviation" in outliers["B4"].message


def test_value_subtypes_cover_addition_clearance_and_replacement() -> None:
    baseline_sheet = SheetSnapshot(
        "Synthetic",
        "visible",
        3,
        1,
        {(2, 1): CellRecord(2, 1, 10), (3, 1): CellRecord(3, 1, 12)},
    )
    current_sheet = SheetSnapshot(
        "Synthetic",
        "visible",
        3,
        1,
        {(1, 1): CellRecord(1, 1, 5), (2, 1): CellRecord(2, 1, 11)},
    )
    region = TableRegion("Synthetic", 1, 1, 3, 1, "block", None, 1, "none")
    alignment = RegionAlignment(
        region,
        region,
        AxisAlignment(pairs=[(1, 1), (2, 2), (3, 3)]),
        AxisAlignment(pairs=[(1, 1)], method="positional"),
    )

    findings = diff_region_values(
        baseline_sheet,
        current_sheet,
        alignment,
        NumericTolerance(),
        ignore=_RangeSet([]),
        refresh=_RangeSet([]),
        sheet_profile=None,
    )
    by_location = {finding.location: finding for finding in findings}

    assert by_location["A1"].subtype is FindingSubtype.VALUE_ADDED_POPULATION
    assert by_location["A2"].subtype is FindingSubtype.VALUE_REPLACEMENT
    assert by_location["A3"].subtype is FindingSubtype.VALUE_CLEARED_POPULATION
    assert by_location["A3"].message.endswith(
        ": historical value cleared (required value is blank)"
    )
    assert not by_location["A3"].expected_growth
    assert all(finding.provenance is None for finding in findings)


def _save_cycle_workbook(path: Path, *, rows: list[list[object]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    if sheet is None:  # pragma: no cover - openpyxl always creates one sheet
        raise RuntimeError("openpyxl did not create a default worksheet")
    sheet.title = "Data"
    sheet.append(["Period", "Amount", "Check"])
    for row in rows:
        sheet.append(row)
    workbook.save(path)


def test_full_pipeline_carries_provenance_and_subtypes_into_history(
    tmp_path: Path,
) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _save_cycle_workbook(
        baseline,
        rows=[
            ["Jan-26", 10, "#REF!"],
            ["Feb-26", 20, 2],
            ["Mar-26", None, 3],
        ],
    )
    _save_cycle_workbook(
        current,
        rows=[
            ["Jan-26", 10, "#REF!"],
            ["Feb-26", 21, 2],
            ["Mar-26", 5, "#DIV/0!"],
        ],
    )

    result = run_qc(baseline_excel=baseline, current_excel=current)
    history = RunHistory(tmp_path / "history.sqlite3")
    run_id = history.record_run(result, file_hashes={}, report_paths={})
    stored = {finding.finding_id: finding for finding in history.get_run(run_id).findings}

    errors = {
        finding.location: finding
        for finding in result.findings
        if finding.finding_class is FindingClass.FORMULA_ERROR
    }
    values = {
        finding.location: finding
        for finding in result.findings
        if finding.finding_class is FindingClass.VALUE_CHANGED
    }

    assert errors["C2"].provenance is FindingProvenance.INHERITED
    assert errors["C4"].provenance is FindingProvenance.NEW
    assert values["B3"].subtype is FindingSubtype.VALUE_REPLACEMENT
    assert values["B4"].subtype is FindingSubtype.VALUE_ADDED_POPULATION
    assert all(
        finding.provenance is None
        for finding in result.findings
        if finding.finding_class
        not in {FindingClass.FORMULA_ERROR, FindingClass.FORMULA_INCONSISTENT}
    )
    assert all(
        (stored[finding.finding_id].provenance, stored[finding.finding_id].subtype)
        == (finding.provenance, finding.subtype)
        for finding in result.findings
    )


def test_additive_fields_stay_out_of_the_cross_run_identity_tuple() -> None:
    base = Finding(
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        sheet="Synthetic",
        location="B4",
        message="synthetic value change",
    )
    enriched = base.model_copy(
        update={
            "provenance": FindingProvenance.INHERITED,
            "subtype": FindingSubtype.VALUE_ADDED_POPULATION,
            "materiality": Materiality.MATERIAL,
            "temporal_context": FindingTemporalContext.RECENT_WINDOW,
            "evidence_tags": {FindingEvidenceTag.RESOLVED_DRIVER},
            "message": "synthetic value addition",
        }
    )
    expected = enriched.model_copy(
        update={"expected_reason": FindingExpectedReason.CADENCE_EXTENSION}
    )
    delta = compare_findings(triage([base]), triage([enriched]))

    assert finding_identity_key(base) == finding_identity_key(enriched)
    assert finding_identity_key(base) == finding_identity_key(expected)
    assert (delta.resolved, delta.new, delta.persisting) == (0, 0, 1)


def test_legacy_history_rows_rehydrate_without_provenance_or_subtype(
    tmp_path: Path,
) -> None:
    database = tmp_path / "history.sqlite3"
    history = RunHistory(database)
    findings = triage(
        [
            Finding(
                artifact="excel",
                finding_class=FindingClass.FORMULA_ERROR,
                sheet="Synthetic",
                location="B2",
                message="synthetic legacy error",
            )
        ]
    )
    run_id = history.record_run(
        QCRunResult(profile_name="legacy", findings=findings),
        file_hashes={},
        report_paths={},
    )
    history.set_annotation(run_id, findings[0].finding_id, severity="info", comment="ok")
    legacy = [
        {
            key: value
            for key, value in finding.model_dump(mode="json").items()
            if key not in {"provenance", "subtype"}
        }
        for finding in findings
    ]
    with sqlite3.connect(database) as connection:
        connection.execute(
            "UPDATE runs SET findings = ? WHERE id = ?",
            (json.dumps(legacy), run_id),
        )

    stored = history.get_run(run_id)

    assert [finding.provenance for finding in stored.findings] == [None]
    assert [finding.subtype for finding in stored.findings] == [None]
    assert stored.findings[0].severity is Severity.INFO
    assert stored.findings[0].analyst_comment == "ok"
    assert finding_identity_key(stored.findings[0]) == (
        "excel",
        "formula_error",
        "Synthetic",
        "",
        "B2",
        "",
    )
