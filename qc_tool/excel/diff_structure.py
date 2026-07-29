"""Workbook structure diff: sheets, visibility, named ranges, charts, pivots.

Named-range changes that are pure range *extensions* consistent with
cadence growth (same sheet and anchor, larger end row/column) are flagged
as expected rather than errors.
"""

import logging
from collections.abc import Callable

from qc_tool.config.profile import DeliverableProfile
from qc_tool.excel.align import WorkbookAlignment
from qc_tool.excel.charts import diff_charts
from qc_tool.excel.interaction import diff_interaction_rules
from qc_tool.excel.references import is_pure_range_extension
from qc_tool.findings import Finding, FindingClass
from qc_tool.io.model import TableDescriptor, WorkbookSnapshot

logger = logging.getLogger(__name__)

def _sheet_findings(alignment: WorkbookAlignment) -> list[Finding]:
    findings = [
        Finding(
            artifact="excel",
            finding_class=FindingClass.SHEET_ADDED,
            sheet=name,
            message=f"sheet {name!r} added in current cycle",
        )
        for name in alignment.added_sheets
    ]
    findings.extend(
        Finding(
            artifact="excel",
            finding_class=FindingClass.SHEET_REMOVED,
            sheet=name,
            message=f"sheet {name!r} removed in current cycle",
        )
        for name in alignment.removed_sheets
    )
    return findings


def _visibility_findings(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot, alignment: WorkbookAlignment
) -> list[Finding]:
    findings = []
    for name in alignment.common_sheets:
        base_state = baseline.sheet(name).visibility
        curr_state = current.sheet(name).visibility
        if base_state != curr_state:
            findings.append(
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.HIDDEN_CHANGED,
                    sheet=name,
                    baseline_value=base_state,
                    current_value=curr_state,
                    message=f"sheet {name!r} visibility changed: {base_state} -> {curr_state}",
                )
            )
    return findings


def _named_range_findings(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot
) -> list[Finding]:
    base_names = {n.name: n.target for n in baseline.named_ranges}
    curr_names = {n.name: n.target for n in current.named_ranges}
    findings = []
    for name in sorted(base_names.keys() | curr_names.keys()):
        base_target = base_names.get(name)
        curr_target = curr_names.get(name)
        if base_target == curr_target:
            continue
        expected = (
            base_target is not None
            and curr_target is not None
            and is_pure_range_extension(base_target, curr_target)
        )
        if base_target is None:
            message = f"named range {name!r} added -> {curr_target}"
        elif curr_target is None:
            message = f"named range {name!r} removed (was {base_target})"
        elif expected:
            message = f"named range {name!r} extended with new-cycle data"
        else:
            message = f"named range {name!r} repointed"
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.NAMED_RANGE_CHANGED,
                expected_growth=expected,
                element=name,
                baseline_value=base_target,
                current_value=curr_target,
                message=message,
            )
        )
    return findings


def _table_pairs(
    baseline: list[TableDescriptor], current: list[TableDescriptor]
) -> tuple[
    list[tuple[TableDescriptor, TableDescriptor]],
    list[TableDescriptor],
    list[TableDescriptor],
]:
    pairs: list[tuple[TableDescriptor, TableDescriptor]] = []
    unmatched_base = list(baseline)
    unmatched_current = list(current)

    def consume_matches(
        matcher: Callable[[TableDescriptor, TableDescriptor], bool],
    ) -> None:
        for base in list(unmatched_base):
            candidates = [
                item
                for item in unmatched_current
                if matcher(base, item)
            ]
            if len(candidates) != 1:
                continue
            curr = candidates[0]
            pairs.append((base, curr))
            unmatched_base.remove(base)
            unmatched_current.remove(curr)

    consume_matches(
        lambda base, curr: base.name.casefold() == curr.name.casefold()
    )
    consume_matches(
        lambda base, curr: base.source_id is not None
        and base.source_id == curr.source_id
        and base.sheet == curr.sheet
    )
    consume_matches(
        lambda base, curr: base.sheet == curr.sheet
        and base.cell_range == curr.cell_range
        and base.columns == curr.columns
    )
    return pairs, unmatched_base, unmatched_current


def _table_findings(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot
) -> list[Finding]:
    if not baseline.tables_available or not current.tables_available:
        return []
    pairs, removed, added = _table_pairs(baseline.tables, current.tables)
    findings = [
        Finding(
            artifact="excel",
            finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
            sheet=table.sheet,
            element=table.name,
            baseline_value=table.cell_range,
            message=f"Excel table {table.name!r} removed",
        )
        for table in removed
    ]
    findings.extend(
        Finding(
            artifact="excel",
            finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
            sheet=table.sheet,
            element=table.name,
            current_value=table.cell_range,
            message=f"Excel table {table.name!r} added",
        )
        for table in added
    )
    for base, curr in pairs:
        if (base.name, base.display_name) != (curr.name, curr.display_name):
            findings.append(
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
                    sheet=curr.sheet,
                    element=curr.name,
                    baseline_value=base.display_name,
                    current_value=curr.display_name,
                    message=(
                        f"Excel table {base.display_name!r} renamed to "
                        f"{curr.display_name!r}"
                    ),
                )
            )
        if base.cell_range != curr.cell_range or base.sheet != curr.sheet:
            base_ref = f"{base.sheet}!{base.cell_range}"
            curr_ref = f"{curr.sheet}!{curr.cell_range}"
            expected = is_pure_range_extension(base_ref, curr_ref)
            findings.append(
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
                    expected_growth=expected,
                    sheet=curr.sheet,
                    element=curr.name,
                    baseline_value=base_ref,
                    current_value=curr_ref,
                    message=(
                        f"Excel table {curr.name!r} range "
                        + (
                            "extended with new-cycle data"
                            if expected
                            else "changed"
                        )
                    ),
                )
            )
        if base.columns != curr.columns:
            base_folded = [item.casefold() for item in base.columns]
            curr_folded = [item.casefold() for item in curr.columns]
            change = (
                "reordered"
                if sorted(base_folded) == sorted(curr_folded)
                else "added, removed, reordered, or renamed"
            )
            findings.append(
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
                    sheet=curr.sheet,
                    element=curr.name,
                    baseline_value=" | ".join(base.columns),
                    current_value=" | ".join(curr.columns),
                    message=f"Excel table {curr.name!r} columns {change}",
                )
            )
        base_rows = (base.header_row_count, base.totals_row_count)
        curr_rows = (curr.header_row_count, curr.totals_row_count)
        if base_rows != curr_rows:
            findings.append(
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
                    sheet=curr.sheet,
                    element=curr.name,
                    baseline_value=f"header={base_rows[0]}, totals={base_rows[1]}",
                    current_value=f"header={curr_rows[0]}, totals={curr_rows[1]}",
                    message=f"Excel table {curr.name!r} header/totals settings changed",
                )
            )
    return findings


def _pivot_findings(baseline: WorkbookSnapshot, current: WorkbookSnapshot) -> list[Finding]:
    base_pivots = {p.name: p for p in baseline.pivots}
    curr_pivots = {p.name: p for p in current.pivots}
    findings = []
    for name in sorted(base_pivots.keys() | curr_pivots.keys()):
        base = base_pivots.get(name)
        curr = curr_pivots.get(name)
        if base is None or curr is None:
            state = "added" if base is None else "removed"
            findings.append(
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.PIVOT_SOURCE_CHANGED,
                    element=name,
                    message=f"pivot table {name!r} {state}",
                )
            )
            continue
        if (base.source_sheet, base.source_ref) != (curr.source_sheet, curr.source_ref):
            base_source = f"{base.source_sheet}!{base.source_ref}"
            curr_source = f"{curr.source_sheet}!{curr.source_ref}"
            expected = is_pure_range_extension(base_source, curr_source)
            findings.append(
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.PIVOT_SOURCE_CHANGED,
                    expected_growth=expected,
                    element=name,
                    baseline_value=base_source,
                    current_value=curr_source,
                    message=(
                        f"pivot table {name!r} source "
                        + ("extended with new-cycle data" if expected else "changed")
                    ),
                )
            )
    return findings


def _unpaired_region_findings(alignment: WorkbookAlignment) -> list[Finding]:
    findings = []
    for region in alignment.unpaired_baseline_regions:
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.REGION_UNPAIRED,
                sheet=region.sheet,
                baseline_location=region.cell_range,
                message=f"baseline region {region.region_id} has no counterpart in current",
            )
        )
    for region in alignment.unpaired_current_regions:
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.REGION_UNPAIRED,
                sheet=region.sheet,
                location=region.cell_range,
                message=f"current region {region.region_id} has no counterpart in baseline",
            )
        )
    return findings


def diff_workbook_structure(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    alignment: WorkbookAlignment,
    profile: DeliverableProfile | None = None,
) -> list[Finding]:
    findings = _sheet_findings(alignment)
    findings.extend(_visibility_findings(baseline, current, alignment))
    findings.extend(_named_range_findings(baseline, current))
    findings.extend(_table_findings(baseline, current))
    findings.extend(diff_interaction_rules(baseline, current))
    findings.extend(diff_charts(baseline, current, profile))
    findings.extend(_pivot_findings(baseline, current))
    findings.extend(_unpaired_region_findings(alignment))
    return findings
