"""Step 5: consolidated structural event semantics at the producing layer."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from qc_tool.engine import compare_findings
from qc_tool.excel.align import AxisAlignment, RegionAlignment
from qc_tool.excel.diff_values import _axis_findings
from qc_tool.excel.interaction import diff_interaction_rules
from qc_tool.excel.regions import TableRegion
from qc_tool.findings import Finding, FindingClass, FindingSubtype, Severity
from qc_tool.io.model import (
    CellRecord,
    CellValue,
    ConditionalFormatDescriptor,
    SheetSnapshot,
    WorkbookSnapshot,
)
from qc_tool.review import finding_identity_key
from qc_tool.triage.rules import assign_severity, triage

_ORACLE_DIR = Path(__file__).parent / "oracles"


def _oracle(name: str) -> dict[str, Any]:
    return json.loads((_ORACLE_DIR / name).read_text(encoding="utf-8"))


def _format(
    formula: str,
    *,
    priority: int,
    source_index: int,
    style_key: str = "synthetic-style",
) -> ConditionalFormatDescriptor:
    return ConditionalFormatDescriptor(
        sheet="Synthetic",
        source_index=source_index,
        source_id=f"synthetic-rule-{formula}",
        target_ranges=("B2:B5",),
        rule_type="expression",
        formulas=(formula,),
        priority=priority,
        stop_if_true=False,
        style_key=style_key,
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


def _events(findings: list[Finding]) -> set[str]:
    return {finding.event_key for finding in findings}


def _axis_run(
    axis: AxisAlignment,
    *,
    is_rows: bool = True,
    region_id: str,
    base_keys: dict[int, CellValue] | None = None,
    curr_keys: dict[int, CellValue] | None = None,
    base_formulas: dict[int, str] | None = None,
    curr_formulas: dict[int, str] | None = None,
) -> list[Finding]:
    """Run `_axis_findings` against minimal synthetic sheets.

    Key cells live in column 1 for row axes and in header row 1 for column
    axes, matching the detected-region conventions.
    """

    def sheet(
        keys: dict[int, CellValue] | None, formulas: dict[int, str] | None
    ) -> SheetSnapshot:
        cells: dict[tuple[int, int], CellRecord] = {}
        for index, value in (keys or {}).items():
            coordinate = (index, 1) if is_rows else (1, index)
            formula = (formulas or {}).get(index)
            cells[coordinate] = CellRecord(
                coordinate[0],
                coordinate[1],
                value,
                formula=formula,
                is_formula=formula is not None,
            )
        return SheetSnapshot("Synthetic", "visible", 20, 20, cells)

    orientation = "long" if is_rows else "wide"
    period_axis = "rows" if is_rows else "columns"
    region = TableRegion("Synthetic", 1, 1, 20, 20, orientation, 1, 1, period_axis)
    event_axis = axis
    other_axis = AxisAlignment(pairs=[(1, 1)])
    alignment = RegionAlignment(
        region,
        region,
        event_axis if is_rows else other_axis,
        other_axis if is_rows else event_axis,
    )
    return _axis_findings(
        sheet(base_keys, base_formulas),
        sheet(curr_keys, curr_formulas),
        alignment,
        axis,
        is_rows=is_rows,
        region_id=region_id,
    )


def test_step5_scenarios_are_owned_by_this_slice() -> None:
    manifest = _oracle("step1_scenarios.json")
    owned = {
        scenario["id"]
        for scenario in manifest["scenarios"]
        if scenario["owner_step"] == 5
    }

    assert owned == {
        "conditional-rule-insertion",
        "conditional-rule-reorder",
        "axis-rolling-turnover",
        "axis-key-replacement",
        "axis-physical-insertion",
        "axis-physical-deletion",
    }


def test_step5_checkpoint_reconciles_with_the_delta_ledger() -> None:
    step4 = _oracle("real_workload_step4.json")
    step5 = _oracle("real_workload_step5.json")
    ledger = _oracle("step_delta_ledger.json")["steps"]["5"]

    assert step5["atomic_findings"] - step4["atomic_findings"] == (
        ledger["expected_atomic_delta"]
    )
    assert step5["conditional_format_findings"] - 123 == (
        ledger["expected_conditional_format_delta"]
    )
    assert step5["conditional_order_findings"] == (
        ledger["expected_conditional_order_findings"]
    )
    assert step5["axis_event_groups"] == ledger["expected_axis_event_groups"]
    assert sum(step5["axis_subtypes"].values()) == ledger["expected_axis_atomics"]
    assert {
        severity: step5["severity"][severity] - step4["severity"][severity]
        for severity in step4["severity"]
    } == ledger["expected_severity_delta"]
    assert sum(step5["review_counts"].values()) - sum(step4["review_counts"].values()) == (
        ledger["expected_review_count_delta"]
    )
    assert {
        status: step5["reference_statuses"][status] - step4["reference_statuses"][status]
        for status in step4["reference_statuses"]
    } == ledger["expected_reference_status_delta"]
    assert step5["coverage_states"] == step4["coverage_states"]
    assert step5["projected_concrete_edges"] == step4["projected_concrete_edges"]
    assert step5["error_provenance"] == step4["error_provenance"]
    assert step5["outlier_provenance"] == step4["outlier_provenance"]
    assert step5["value_subtypes"] == step4["value_subtypes"]
    assert step5["unchanged_chart_findings"] == {
        "chart_geometry_changed": 1,
        "chart_labels_changed": 8,
    }
    assert step5["source_hashes_unchanged"] is True


def test_suppressed_fanout_resolves_naturally_in_a_rerun() -> None:
    def fanout(index: int) -> Finding:
        return Finding(
            artifact="excel",
            finding_class=FindingClass.CONDITIONAL_FORMAT_CHANGED,
            sheet="Synthetic",
            element=f"expression rule {index}",
            message=f"conditional-format expression rule {index} priority changed",
        )

    previous = triage([fanout(1), fanout(2), fanout(3)])
    current = triage([fanout(1)])

    delta = compare_findings(previous, current)

    assert (delta.resolved, delta.new, delta.persisting) == (2, 0, 1)
    assert {finding_identity_key(finding) for finding in current} <= {
        finding_identity_key(finding) for finding in previous
    }


def test_pure_renumbering_from_additions_and_removals_emits_no_order_findings() -> None:
    baseline = _rules(
        [
            _format("B2>10", priority=1, source_index=0),
            _format("B2<0", priority=2, source_index=1),
            _format("B2=1", priority=3, source_index=2),
        ]
    )
    current = _rules(
        [
            _format("B2=0", priority=1, source_index=0),
            _format("B2>10", priority=2, source_index=1),
            _format("B2<0", priority=3, source_index=2),
            _format("B2=1", priority=4, source_index=3),
        ]
    )

    findings = diff_interaction_rules(baseline, current)

    assert not any(
        finding.subtype is FindingSubtype.OBJECT_ORDER_CHANGED for finding in findings
    )
    assert [finding.subtype for finding in findings] == [FindingSubtype.OBJECT_ADDED]


def test_only_the_inverted_pair_reports_an_order_change() -> None:
    baseline = _rules(
        [
            _format("B2>10", priority=1, source_index=0),
            _format("B2<0", priority=2, source_index=1),
            _format("B2=1", priority=3, source_index=2),
        ]
    )
    current = _rules(
        [
            _format("B2<0", priority=1, source_index=0),
            _format("B2>10", priority=2, source_index=1),
            _format("B2=1", priority=3, source_index=2),
        ]
    )

    order_findings = [
        finding
        for finding in diff_interaction_rules(baseline, current)
        if finding.subtype is FindingSubtype.OBJECT_ORDER_CHANGED
    ]

    assert len(order_findings) == 2
    assert {finding.baseline_value for finding in order_findings} == {"1", "2"}
    assert all("evaluation order" in finding.message for finding in order_findings)


def test_one_rule_edit_shares_one_stable_event_key() -> None:
    baseline = _rules([_format("B2>10", priority=1, source_index=0)])
    current = _rules(
        [_format("B2>20", priority=1, source_index=0, style_key="rgb-green")]
    )

    findings = diff_interaction_rules(baseline, current)

    assert len(findings) == 2
    assert len(_events(findings)) == 1
    assert _events(findings) == {"excel:Synthetic:conditional-format:synthetic-rule-B2>20"}
    assert {finding.subtype for finding in findings} == {
        FindingSubtype.OBJECT_CONDITION_CHANGED,
        FindingSubtype.OBJECT_STYLE_CHANGED,
    }


def test_axis_events_are_typed_per_region_axis() -> None:
    rolling = AxisAlignment(pairs=[(2, 1), (3, 2)], deleted=[1], growth=[3])
    replacement = AxisAlignment(pairs=[(1, 1), (3, 3)], deleted=[2], inserted=[2])
    insertion = AxisAlignment(pairs=[(1, 1), (2, 3)], inserted=[2])
    deletion = AxisAlignment(pairs=[(1, 1), (3, 2)], deleted=[2])
    growth = AxisAlignment(pairs=[(1, 1)], growth=[2, 3])

    findings = {
        name: _axis_run(axis, region_id=f"Synthetic!{name}")
        for name, axis in (
            ("rolling", rolling),
            ("replacement", replacement),
            ("insertion", insertion),
            ("deletion", deletion),
            ("growth", growth),
        )
    }

    assert {finding.subtype for finding in findings["rolling"]} == {
        FindingSubtype.AXIS_ROLLING_TURNOVER
    }
    assert {finding.subtype for finding in findings["replacement"]} == {
        FindingSubtype.AXIS_KEY_REPLACEMENT
    }
    assert {finding.finding_class for finding in findings["replacement"]} == {
        FindingClass.ROW_KEY_CHANGED
    }
    assert len(findings["replacement"]) == 1
    assert {finding.subtype for finding in findings["insertion"]} == {
        FindingSubtype.AXIS_PHYSICAL_INSERTION
    }
    assert {finding.subtype for finding in findings["deletion"]} == {
        FindingSubtype.AXIS_PHYSICAL_DELETION
    }
    assert {finding.subtype for finding in findings["growth"]} == {
        FindingSubtype.AXIS_EXTENT_GROWTH
    }
    assert all(
        len(_events(group)) == 1 for group in findings.values()
    )
    assert {
        finding.finding_class
        for group in findings.values()
        for finding in group
    } <= {
        FindingClass.ROW_DELETED,
        FindingClass.ROW_INSERTED,
        FindingClass.ROW_GROWTH,
        FindingClass.ROW_KEY_CHANGED,
    }


def test_key_replacement_wording_never_claims_a_physical_row_edit() -> None:
    replacement = AxisAlignment(pairs=[(1, 1), (3, 3)], deleted=[2], inserted=[2])
    rolling = AxisAlignment(pairs=[(2, 1), (3, 2)], deleted=[1], growth=[3])

    replaced = _axis_run(
        replacement,
        region_id="Synthetic!A1:B3",
        base_keys={2: "Old Key"},
        curr_keys={2: "New Key"},
    )
    turned_over = _axis_run(rolling, region_id="Synthetic!A1:B3")

    assert all("deleted" not in finding.message for finding in replaced)
    assert all("inserted" not in finding.message for finding in replaced)
    assert len(replaced) == 1
    assert "replaced in place" in replaced[0].message
    assert "'Old Key' -> 'New Key'" in replaced[0].message
    assert replaced[0].baseline_value == "Old Key"
    assert replaced[0].current_value == "New Key"
    assert any("left the rolling window" in finding.message for finding in turned_over)
    assert {finding_identity_key(finding) for finding in replaced} == {
        ("excel", "row_key_changed", "Synthetic", "", "row 2", ""),
    }


def test_constant_key_change_is_critical_and_derived_label_is_warning() -> None:
    axis = AxisAlignment(pairs=[(1, 1)], deleted=[2], inserted=[2])

    constant = _axis_run(
        axis,
        region_id="Synthetic!A1:B3",
        base_keys={2: "Old"},
        curr_keys={2: "New"},
    )[0]
    derived = _axis_run(
        axis,
        region_id="Synthetic!A1:B3",
        base_keys={2: "Old Brand"},
        curr_keys={2: "New Brand"},
        base_formulas={2: "=INDEX(Setup!A:A,MATCH(1,Setup!B:B,0))"},
        curr_formulas={2: "=INDEX(Setup!A:A,MATCH(1,Setup!B:B,0))"},
    )[0]
    half_derived = _axis_run(
        axis,
        region_id="Synthetic!A1:B3",
        base_keys={2: "Old"},
        curr_keys={2: "New"},
        curr_formulas={2: "=INDEX(Setup!A:A,1)"},
    )[0]

    assert constant.subtype is FindingSubtype.AXIS_KEY_REPLACEMENT
    assert assign_severity(constant) is Severity.CRITICAL

    assert derived.subtype is FindingSubtype.AXIS_KEY_DERIVED_LABEL
    assert assign_severity(derived) is Severity.WARNING
    assert "formula-derived" in derived.message
    assert "upstream driver changed" in derived.message

    assert half_derived.subtype is FindingSubtype.AXIS_KEY_REPLACEMENT
    assert assign_severity(half_derived) is Severity.CRITICAL


def test_rolling_turnover_is_cadence_design_not_a_deletion() -> None:
    rolling = AxisAlignment(pairs=[(2, 1), (3, 2)], deleted=[1], growth=[3])
    turned_over = _axis_run(rolling, region_id="Synthetic!A1:B3")

    deletions = [
        finding
        for finding in turned_over
        if finding.subtype is FindingSubtype.AXIS_ROLLING_TURNOVER
        and finding.finding_class is FindingClass.ROW_DELETED
    ]
    assert deletions
    assert all(assign_severity(finding) is Severity.WARNING for finding in deletions)


def test_advanced_period_key_is_expected_cadence() -> None:
    axis = AxisAlignment(pairs=[(1, 1)], deleted=[2], inserted=[2])
    advanced = _axis_run(
        axis,
        region_id="Synthetic!A1:B3",
        base_keys={2: "Jun-26"},
        curr_keys={2: "Jul-26"},
    )[0]
    regressed = _axis_run(
        axis,
        region_id="Synthetic!A1:B3",
        base_keys={2: "Jul-26"},
        curr_keys={2: "Jun-26"},
    )[0]

    assert advanced.expected_growth
    assert assign_severity(advanced) is Severity.EXPECTED
    assert "tracking period advanced" in advanced.message
    assert not regressed.expected_growth
    assert assign_severity(regressed) is Severity.CRITICAL


def test_partially_reused_positions_are_typed_per_index_inside_one_event() -> None:
    axis = AxisAlignment(pairs=[(1, 1)], deleted=[2, 9], inserted=[2, 5], growth=[12])

    findings = _axis_run(axis, region_id="Synthetic!A1:B12")
    by_key = {
        (finding.finding_class, finding.baseline_location or finding.location): finding
        for finding in findings
    }

    assert len(_events(findings)) == 1
    assert len(findings) == 4
    assert by_key[(FindingClass.ROW_KEY_CHANGED, "row 2")].subtype is (
        FindingSubtype.AXIS_KEY_REPLACEMENT
    )
    assert by_key[(FindingClass.ROW_DELETED, "row 9")].subtype is (
        FindingSubtype.AXIS_PHYSICAL_DELETION
    )
    assert by_key[(FindingClass.ROW_INSERTED, "row 5")].subtype is (
        FindingSubtype.AXIS_PHYSICAL_INSERTION
    )
    assert by_key[(FindingClass.ROW_GROWTH, "row 12")].subtype is (
        FindingSubtype.AXIS_EXTENT_GROWTH
    )
    assert "deleted" not in by_key[(FindingClass.ROW_KEY_CHANGED, "row 2")].message


def test_axis_events_do_not_enter_finding_identity_or_review_partitioning() -> None:
    axis = AxisAlignment(pairs=[(1, 1)], deleted=[2], inserted=[2])

    findings = _axis_run(axis, is_rows=False, region_id="Synthetic!A1:C3")

    assert all(finding.event_key == "excel:Synthetic!A1:C3:columns" for finding in findings)
    assert {finding_identity_key(finding) for finding in findings} == {
        ("excel", "column_key_changed", "Synthetic", "", "column B", ""),
    }
