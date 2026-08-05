"""Workbook structure diff: sheets, visibility, named ranges, charts, pivots.

Named-range changes that are pure range *extensions* consistent with
cadence growth (same sheet and anchor, larger end row/column) are flagged
as expected rather than errors.
"""

import logging
from collections.abc import Callable

from openpyxl.utils.cell import range_boundaries

from qc_tool.config.profile import DeliverableProfile
from qc_tool.excel.align import WorkbookAlignment
from qc_tool.excel.charts import diff_charts
from qc_tool.excel.interaction import diff_interaction_rules
from qc_tool.excel.periods import is_period_after, parse_period
from qc_tool.excel.references import is_pure_range_extension
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingExpectedReason,
    FindingSubtype,
)
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
    base_names = {(n.sheet, n.name): n.target for n in baseline.named_ranges}
    curr_names = {(n.sheet, n.name): n.target for n in current.named_ranges}
    findings = []
    for scope, name in sorted(
        base_names.keys() | curr_names.keys(),
        key=lambda key: (key[0] or "", key[1]),
    ):
        key = (scope, name)
        base_target = base_names.get(key)
        curr_target = curr_names.get(key)
        if base_target == curr_target:
            continue
        expected = (
            base_target is not None
            and curr_target is not None
            and is_pure_range_extension(base_target, curr_target)
        )
        label = name if scope is None else f"{scope}!{name}"
        if base_target is None:
            message = f"named range {label!r} added -> {curr_target}"
        elif curr_target is None:
            message = f"named range {label!r} removed (was {base_target})"
        elif expected:
            message = f"named range {label!r} extended with new-cycle data"
        else:
            message = f"named range {label!r} repointed"
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.NAMED_RANGE_CHANGED,
                expected_reason=(
                    FindingExpectedReason.CADENCE_EXTENSION if expected else None
                ),
                sheet=scope,
                element=label,
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


def _table_event_key(table: TableDescriptor) -> str:
    return f"excel:{table.sheet}:table:{table.name}"


def _is_row_only_extension(base_range: str, curr_range: str) -> bool:
    """Same anchor and column span; only the trailing row edge grew."""
    try:
        base_bounds = range_boundaries(base_range)
        curr_bounds = range_boundaries(curr_range)
    except ValueError:
        return False
    return (
        base_bounds[0] == curr_bounds[0]
        and base_bounds[1] == curr_bounds[1]
        and base_bounds[2] == curr_bounds[2]
        and (curr_bounds[3] or 0) >= (base_bounds[3] or 0)
    )


def _table_data_rows(
    table: TableDescriptor,
) -> tuple[int, int, int, int] | None:
    try:
        min_col, min_row, max_col, max_row = range_boundaries(table.cell_range)
    except ValueError:
        return None
    if min_col is None or min_row is None or max_col is None or max_row is None:
        return None
    data_start = min_row + table.header_row_count
    data_end = max_row - table.totals_row_count
    if data_start > data_end:
        return None
    return min_col, data_start, max_col, data_end


def _profile_covers_table_extension(
    profile: DeliverableProfile | None,
    table: TableDescriptor,
    *,
    first_new_row: int,
    last_new_row: int,
    min_col: int,
    max_col: int,
) -> bool:
    sheet_profile = profile.sheet_profile(table.sheet) if profile is not None else None
    if sheet_profile is None:
        return False
    for cell_range in sheet_profile.refresh_ranges:
        try:
            left, top, right, bottom = range_boundaries(cell_range)
        except ValueError:
            continue
        if left is None or top is None or right is None or bottom is None:
            continue
        if (
            left <= min_col
            and right >= max_col
            and top <= first_new_row
            and bottom >= last_new_row
        ):
            return True
    return False


def _table_period_extension_is_proved(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    base: TableDescriptor,
    curr: TableDescriptor,
) -> bool:
    base_bounds = _table_data_rows(base)
    curr_bounds = _table_data_rows(curr)
    if base_bounds is None or curr_bounds is None:
        return False
    base_min_col, base_start, base_max_col, base_end = base_bounds
    curr_min_col, curr_start, _curr_max_col, curr_end = curr_bounds
    if (
        base_min_col != curr_min_col
        or base_start != curr_start
        or curr_end <= base_end
    ):
        return False
    try:
        base_sheet = baseline.sheet(base.sheet)
        curr_sheet = current.sheet(curr.sheet)
    except KeyError:
        return False
    for column in range(base_min_col, base_max_col + 1):
        base_cells = [
            base_sheet.cells.get((row, column))
            for row in range(base_start, base_end + 1)
        ]
        populated_base = [
            cell
            for cell in base_cells
            if cell is not None and cell.value is not None
        ]
        if not populated_base or any(cell.has_formula for cell in populated_base):
            continue
        base_periods = [parse_period(cell.value) for cell in populated_base]
        if any(period is None for period in base_periods):
            continue
        typed_base_periods = [period for period in base_periods if period is not None]
        kinds = {period.kind for period in typed_base_periods}
        if len(kinds) != 1:
            continue
        edge = max(typed_base_periods, key=lambda period: period.sort_key)
        new_cells = [
            curr_sheet.cells.get((row, column))
            for row in range(base_end + 1, curr_end + 1)
        ]
        if any(cell is None or cell.has_formula for cell in new_cells):
            continue
        new_periods = [parse_period(cell.value) for cell in new_cells if cell is not None]
        if new_periods and all(
            period is not None
            and period.kind == edge.kind
            and is_period_after(period, edge)
            for period in new_periods
        ):
            return True
    return False


def _table_extension_reason(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    base: TableDescriptor,
    curr: TableDescriptor,
    profile: DeliverableProfile | None,
) -> FindingExpectedReason | None:
    if (
        base.columns != curr.columns
        or base.sheet != curr.sheet
        or not _is_row_only_extension(base.cell_range, curr.cell_range)
        or not is_pure_range_extension(
            f"{base.sheet}!{base.cell_range}", f"{curr.sheet}!{curr.cell_range}"
        )
    ):
        return None
    base_bounds = _table_data_rows(base)
    curr_bounds = _table_data_rows(curr)
    if base_bounds is None or curr_bounds is None or curr_bounds[3] <= base_bounds[3]:
        return None
    if _profile_covers_table_extension(
        profile,
        curr,
        first_new_row=base_bounds[3] + 1,
        last_new_row=curr_bounds[3],
        min_col=curr_bounds[0],
        max_col=curr_bounds[2],
    ):
        return FindingExpectedReason.PROFILE_REFRESH
    if _table_period_extension_is_proved(baseline, current, base, curr):
        return FindingExpectedReason.CADENCE_EXTENSION
    return None


def _table_findings(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    profile: DeliverableProfile | None = None,
) -> list[Finding]:
    if not baseline.tables_available or not current.tables_available:
        return []
    pairs, removed, added = _table_pairs(baseline.tables, current.tables)
    findings = [
        Finding(
            artifact="excel",
            finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
            subtype=FindingSubtype.OBJECT_REMOVED,
            event_key=_table_event_key(table),
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
            subtype=FindingSubtype.OBJECT_ADDED,
            event_key=_table_event_key(table),
            sheet=table.sheet,
            element=table.name,
            current_value=table.cell_range,
            message=f"Excel table {table.name!r} added",
        )
        for table in added
    )
    for base, curr in pairs:
        event_key = _table_event_key(curr)
        if (base.name, base.display_name) != (curr.name, curr.display_name):
            findings.append(
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
                    subtype=FindingSubtype.OBJECT_RENAMED,
                    event_key=event_key,
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
            columns_changed = base.columns != curr.columns
            expected_reason = _table_extension_reason(
                baseline,
                current,
                base,
                curr,
                profile,
            )
            if expected_reason is FindingExpectedReason.CADENCE_EXTENSION:
                wording = "extended with proved later-period rows"
            elif expected_reason is FindingExpectedReason.PROFILE_REFRESH:
                wording = "extended inside a profile refresh range"
            elif columns_changed:
                wording = "changed alongside column additions or removals"
            elif _is_row_only_extension(base.cell_range, curr.cell_range):
                wording = "extended by rows without proved cadence evidence"
            else:
                wording = "changed"
            findings.append(
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
                    expected_reason=expected_reason,
                    subtype=FindingSubtype.OBJECT_TARGET_CHANGED,
                    event_key=event_key,
                    sheet=curr.sheet,
                    element=curr.name,
                    baseline_value=base_ref,
                    current_value=curr_ref,
                    message=f"Excel table {curr.name!r} range " + wording,
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
                    subtype=FindingSubtype.OBJECT_COLUMNS_CHANGED,
                    event_key=event_key,
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
                    subtype=FindingSubtype.OBJECT_SETTINGS_CHANGED,
                    event_key=event_key,
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
                    expected_reason=(
                        FindingExpectedReason.CADENCE_EXTENSION if expected else None
                    ),
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
    findings.extend(_table_findings(baseline, current, profile))
    findings.extend(diff_interaction_rules(baseline, current))
    findings.extend(diff_charts(baseline, current, profile))
    findings.extend(_pivot_findings(baseline, current))
    findings.extend(_unpaired_region_findings(alignment))
    return findings
