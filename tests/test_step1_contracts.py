"""Frozen contracts for the real-workload optimization plan."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from qc_tool.config.profile import NumericTolerance
from qc_tool.engine import QCRunResult, compare_findings
from qc_tool.excel.align import AxisAlignment, RegionAlignment, _align_axis
from qc_tool.excel.diff_values import _axis_findings, _RangeSet, diff_region_values
from qc_tool.excel.interaction import diff_interaction_rules
from qc_tool.excel.periods import parse_period
from qc_tool.excel.regions import TableRegion
from qc_tool.findings import Finding, FindingClass, FindingSubtype, Severity
from qc_tool.history.store import RunHistory
from qc_tool.io.model import (
    CellRecord,
    ConditionalFormatDescriptor,
    SheetSnapshot,
    WorkbookSnapshot,
)
from qc_tool.review import build_review_groups, finding_identity_key
from qc_tool.triage.rules import triage

_ORACLE_DIR = Path(__file__).parent / "oracles"


def _oracle(name: str) -> dict[str, Any]:
    return json.loads((_ORACLE_DIR / name).read_text(encoding="utf-8"))


def _format(
    formula: str,
    *,
    priority: int,
    source_index: int,
) -> ConditionalFormatDescriptor:
    return ConditionalFormatDescriptor(
        sheet="Synthetic",
        source_index=source_index,
        source_id=f"synthetic-rule-{source_index}",
        target_ranges=("B2:B5",),
        rule_type="expression",
        formulas=(formula,),
        priority=priority,
        stop_if_true=False,
        style_key="synthetic-style",
    )


def _rules(formats: list[ConditionalFormatDescriptor]) -> WorkbookSnapshot:
    return WorkbookSnapshot(
        "synthetic.xlsx",
        "xlsx",
        True,
        True,
        sheets=[SheetSnapshot("Synthetic", "visible", 10, 5)],
        conditional_formats=formats,
    )


def test_step1_manifest_covers_every_planned_semantic_boundary() -> None:
    manifest = _oracle("step1_scenarios.json")
    scenarios = manifest["scenarios"]
    assert isinstance(scenarios, list)
    identifiers = {scenario["id"] for scenario in scenarios}

    assert identifiers == {
        "lexical-let",
        "lexical-lambda",
        "lexical-nested-scope",
        "lexical-shadowing",
        "lexical-higher-order",
        "lexical-unsupported-syntax",
        "formula-error-inherited",
        "formula-error-new",
        "consistency-historical",
        "consistency-novel",
        "value-added-population",
        "value-replacement",
        "conditional-rule-insertion",
        "conditional-rule-reorder",
        "axis-rolling-turnover",
        "axis-key-replacement",
        "axis-physical-insertion",
        "axis-physical-deletion",
        "repeated-transformation-patterns",
        "invalid-heavy-dependencies",
        "range-heavy-dependencies",
        "xlsb-customxml-relationship",
        "worker-crash",
        "queue-refresh",
    }
    serialized = json.dumps(manifest, sort_keys=True).casefold()
    assert all(
        forbidden not in serialized
        for forbidden in ("password", "formula_text", "source_path", "workbook_name")
    )


def test_anonymous_before_state_and_delta_ledger_reconcile() -> None:
    before = _oracle("real_workload_before_state.json")
    step2 = _oracle("real_workload_step2.json")
    step3 = _oracle("real_workload_step3.json")
    ledger = _oracle("step_delta_ledger.json")
    workloads = before["workloads"]
    assert isinstance(workloads, dict)
    cycle = workloads["cycle_pair"]
    assert cycle["atomic_findings"] == sum(cycle["severity"].values())
    assert cycle["reference_operands"] == sum(cycle["reference_statuses"].values())
    assert cycle["value_findings"]["total"] == (
        cycle["value_findings"]["added_population"]
        + cycle["value_findings"]["replacements"]
    )
    assert ledger["steps"]["2"]["expected_atomic_delta"] == 0
    assert ledger["steps"]["3"]["expected_atomic_delta"] == 0
    before_statuses = cycle["reference_statuses"]
    step2_statuses = step2["reference_statuses"]
    step3_statuses = step3["reference_statuses"]
    step2_delta = {
        status: step2_statuses[status] - before_statuses[status]
        for status in ("invalid", "resolved", "unsupported")
    }
    assert step2_delta == ledger["steps"]["2"]["expected_reference_status_delta"]
    observed_delta = {
        "invalid": step3_statuses["invalid"] - step2_statuses["invalid"],
        "local_symbol": step3_statuses["local_symbol"],
        "resolved": step3_statuses["resolved"] - step2_statuses["resolved"],
        "unsupported": (
            step3_statuses["unsupported"] - step2_statuses["unsupported"]
        ),
    }
    assert observed_delta == ledger["steps"]["3"]["anonymous_reference_status_delta"]
    assert ledger["steps"]["3"]["synthetic_reference_status_delta"] == {
        "invalid": -6,
        "local_symbol": 6,
        "resolved": 0,
        "unsupported": 0,
    }
    assert step2["operand_count"] == sum(step2_statuses.values())
    assert step3["operand_count"] == sum(step3_statuses.values())
    for checkpoint in (step2, step3):
        assert checkpoint["atomic_findings"] == cycle["atomic_findings"]
        assert checkpoint["coverage_states"] == cycle["coverage_states"]
        assert checkpoint["projected_concrete_edges"] == cycle["projected_concrete_edges"]
        assert checkpoint["source_hashes_unchanged"] is True


def test_final_checkpoint_matches_the_last_semantic_checkpoint() -> None:
    """Steps 11-13 change execution and validation only, never findings."""
    step10 = _oracle("real_workload_step10.json")
    step13 = _oracle("real_workload_step13.json")
    ledger = _oracle("step_delta_ledger.json")["steps"]

    for step in ("11", "12", "13"):
        assert ledger[step]["expected_atomic_delta"] == 0
        assert ledger[step]["expected_review_count_delta"] == 0
        assert ledger[step]["expected_pattern_review_delta"] == 0
        assert ledger[step]["expected_coverage_transitions"] == []
    assert ledger["13"]["expected_equal_to_checkpoint"] == "real_workload_step10.json"

    for key in (
        "atomic_findings",
        "coverage_states",
        "mixed_pattern_groups",
        "pattern_review_counts",
        "phases",
        "review_counts",
        "severity",
    ):
        assert step13[key] == step10[key], key
    assert step13["source_hashes_unchanged"] is True
    assert step13["trials"] >= 3
    # Plan targets G1/G2 on the anonymous cycle pair.
    assert step13["median_elapsed_seconds"] <= 60
    assert step13["median_peak_rss_mib"] <= 800


def test_value_additions_and_replacements_use_distinct_subtypes_and_wording() -> None:
    baseline_sheet = SheetSnapshot(
        "Synthetic",
        "visible",
        2,
        1,
        {(2, 1): CellRecord(2, 1, 10)},
    )
    current_sheet = SheetSnapshot(
        "Synthetic",
        "visible",
        2,
        1,
        {
            (1, 1): CellRecord(1, 1, 5),
            (2, 1): CellRecord(2, 1, 11),
        },
    )
    region = TableRegion("Synthetic", 1, 1, 2, 1, "block", None, 1, "none")
    alignment = RegionAlignment(
        region,
        region,
        AxisAlignment(pairs=[(1, 1), (2, 2)]),
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
    assert by_location["A1"].baseline_value == ""
    assert by_location["A1"].current_value == "5"
    assert by_location["A2"].baseline_value == "10"
    assert by_location["A2"].current_value == "11"
    assert by_location["A1"].subtype is FindingSubtype.VALUE_ADDED_POPULATION
    assert by_location["A2"].subtype is FindingSubtype.VALUE_REPLACEMENT
    assert by_location["A1"].message.endswith(": value added to a previously blank cell")
    assert by_location["A2"].message.endswith(": historical value changed")
    assert {
        finding_identity_key(finding) for finding in findings
    } == {
        ("excel", "value_changed", "Synthetic", "", "A1", ""),
        ("excel", "value_changed", "Synthetic", "", "A2", ""),
    }


def test_conditional_rule_insertion_no_longer_fans_out_priority_renumbering() -> None:
    baseline = _rules(
        [
            _format("B2>10", priority=1, source_index=0),
            _format("B2<0", priority=2, source_index=1),
        ]
    )
    current = _rules(
        [
            _format("B2=0", priority=1, source_index=0),
            _format("B2>10", priority=2, source_index=1),
            _format("B2<0", priority=3, source_index=2),
        ]
    )

    findings = diff_interaction_rules(baseline, current)

    assert sum("evaluation order" in finding.message for finding in findings) == 0
    assert sum("added" in finding.message for finding in findings) == 1


def test_conditional_rule_reorder_still_records_both_inverted_rules() -> None:
    baseline = _rules(
        [
            _format("B2>10", priority=1, source_index=0),
            _format("B2<0", priority=2, source_index=1),
        ]
    )
    current = _rules(
        [
            _format("B2<0", priority=1, source_index=0),
            _format("B2>10", priority=2, source_index=1),
        ]
    )

    findings = diff_interaction_rules(baseline, current)

    assert len(findings) == 2
    assert all("evaluation order" in finding.message for finding in findings)


def test_axis_before_state_distinguishes_three_aggregate_shapes() -> None:
    rolling = _align_axis(
        [
            _axis_entry(1, "W01"),
            _axis_entry(2, "W02"),
            _axis_entry(3, "W03"),
        ],
        [
            _axis_entry(1, "W02"),
            _axis_entry(2, "W03"),
            _axis_entry(3, "W04"),
        ],
    )
    replacement = _align_axis(
        [_axis_entry(index, key) for index, key in enumerate("ABCD", start=1)],
        [_axis_entry(index, key) for index, key in enumerate("AXCD", start=1)],
    )
    insertion = _align_axis(
        [_axis_entry(index, key) for index, key in enumerate("ABC", start=1)],
        [_axis_entry(index, key) for index, key in enumerate("AXBC", start=1)],
    )

    assert (rolling.deleted, rolling.growth) == ([1], [3])
    assert (replacement.deleted, replacement.inserted) == ([2], [2])
    assert (insertion.deleted, insertion.inserted) == ([], [2])

    def axis_findings(axis: AxisAlignment, region_id: str) -> list[Finding]:
        region = TableRegion("Synthetic", 1, 1, 20, 20, "wide", 1, 1, "columns")
        alignment = RegionAlignment(
            region, region, AxisAlignment(pairs=[(1, 1)]), axis
        )
        empty = SheetSnapshot("Synthetic", "visible", 20, 20, {})
        return _axis_findings(
            empty, empty, alignment, axis, is_rows=False, region_id=region_id
        )

    assert len(axis_findings(rolling, "R1")) == 2
    # The in-place pair is one merged key-change finding now.
    assert len(axis_findings(replacement, "R2")) == 1
    assert len(axis_findings(insertion, "R3")) == 1


def _axis_entry(index: int, key: str):
    from qc_tool.excel.align import AxisEntry

    return AxisEntry(index=index, key=(key, 0), periods=(parse_period(key),))


def test_spatial_review_before_state_can_merge_distinct_formula_transformations() -> None:
    findings = triage(
        [
            Finding(
                artifact="excel",
                finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
                sheet="Synthetic",
                location="B2",
                baseline_value="=A2",
                current_value="=IFERROR(A2,0)",
                message="synthetic formula change one",
            ),
            Finding(
                artifact="excel",
                finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
                sheet="Synthetic",
                location="B3",
                baseline_value="=A3",
                current_value="=A3*2",
                message="synthetic formula change two",
            ),
        ]
    )

    groups = build_review_groups(findings)

    assert len(groups) == 1
    assert groups[0].member_count == 2
    assert len(
        {(member.baseline_value, member.current_value) for member in groups[0].members}
    ) == 2


def test_atomic_identity_annotations_and_rerun_matching_are_frozen(tmp_path: Path) -> None:
    matrix = _oracle("identity_matrix.json")
    expected = {tuple(identity) for identity in matrix["identities"]}
    findings = triage(
        [
            Finding(
                artifact=identity[0],
                finding_class=FindingClass(identity[1]),
                sheet=identity[2] or None,
                slide=identity[3] or None,
                location=identity[4] or None,
                element=identity[5] or None,
                message=f"synthetic {identity[1]}",
            )
            for identity in matrix["identities"]
        ]
    )
    result = QCRunResult(profile_name="identity-oracle", findings=findings)
    history = RunHistory(tmp_path / "history.sqlite3")
    run_id = history.record_run(result, file_hashes={}, report_paths={})
    annotated_id = next(
        finding.finding_id
        for finding in findings
        if finding.finding_class is FindingClass.FORMULA_ERROR
    )
    history.set_annotation(
        run_id,
        annotated_id,
        severity="info",
        comment="synthetic legacy annotation",
    )

    stored = history.get_run(run_id)
    regenerated = [
        finding.model_copy(
            update={
                "message": f"reworded {finding.finding_class.value}",
                "root_cause_key": "synthetic-additive-field",
            }
        )
        for finding in findings
    ]
    delta = compare_findings(findings, regenerated)

    assert {finding_identity_key(finding) for finding in stored.findings} == expected
    assert {finding_identity_key(finding) for finding in regenerated} == expected
    assert (delta.resolved, delta.new, delta.persisting) == (0, 0, len(findings))
    annotated = next(
        finding for finding in stored.findings if finding.finding_id == annotated_id
    )
    assert annotated.severity is Severity.INFO
    assert annotated.analyst_comment == "synthetic legacy annotation"
