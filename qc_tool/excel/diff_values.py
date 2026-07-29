"""Cell-level diff over aligned regions: values, formats, styles, axis events.

Consumes `WorkbookAlignment` correspondences; never re-aligns. Numeric
comparisons honor the profile tolerance. Value changes inside profile
``refresh_ranges`` (cycle-snapshot aggregates) are flagged as expected
growth rather than errors; ``ignore_ranges`` suppress cell findings
entirely.
"""

import logging
from collections.abc import Iterable

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

from qc_tool.availability import excel_blank_allowed
from qc_tool.config.profile import DeliverableProfile, NumericTolerance, SheetProfile
from qc_tool.excel.align import AxisAlignment, RegionAlignment, WorkbookAlignment
from qc_tool.excel.periods import is_period_after, parse_period
from qc_tool.findings import Finding, FindingClass
from qc_tool.io.model import (
    CellRecord,
    SheetSnapshot,
    WorkbookSnapshot,
    display_cell_value,
)
from qc_tool.progress import CancellationToken, check_cancelled

logger = logging.getLogger(__name__)


def _ref(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


class _RangeSet:
    def __init__(self, ranges: Iterable[str]) -> None:
        self._boxes = [range_boundaries(r) for r in ranges]

    def __contains__(self, cell: tuple[int, int]) -> bool:
        row, col = cell
        return any(
            (min_col or 1) <= col <= (max_col or col)
            and (min_row or 1) <= row <= (max_row or row)
            for min_col, min_row, max_col, max_row in self._boxes
        )


def _numbers_match(base: float, curr: float, tolerance: NumericTolerance) -> bool:
    delta = abs(curr - base)
    if delta == 0:
        return True
    within_abs = delta <= tolerance.absolute
    within_rel = base != 0 and delta / abs(base) <= tolerance.relative
    return within_abs or within_rel


def _values_differ(
    base: CellRecord | None, curr: CellRecord | None, tolerance: NumericTolerance
) -> bool:
    base_value = base.value if base else None
    curr_value = curr.value if curr else None
    if base_value is None and curr_value is None:
        return False
    if isinstance(base_value, int | float) and isinstance(curr_value, int | float):
        if isinstance(base_value, bool) or isinstance(curr_value, bool):
            return base_value != curr_value
        return not _numbers_match(float(base_value), float(curr_value), tolerance)
    return base_value != curr_value


def _display(value: object) -> str:
    return display_cell_value(value)


def _is_blank(cell: CellRecord | None) -> bool:
    return cell is None or cell.value is None or (
        isinstance(cell.value, str) and not cell.value.strip()
    )


def _period_advanced(base_value: object, curr_value: object) -> bool:
    """True when both values are period labels and the current one is later."""
    base_period = parse_period(base_value)
    curr_period = parse_period(curr_value)
    return (
        base_period is not None
        and curr_period is not None
        and is_period_after(curr_period, base_period)
    )


def _is_refresh_block(
    base_sheet: SheetSnapshot, curr_sheet: SheetSnapshot, region: RegionAlignment
) -> bool:
    """A block region containing an advancing period label is a cycle-snapshot
    block (e.g. a KPI panel with "Cycle: Jun-26"): its value changes refresh
    every cadence. Long/wide historical tables never get this treatment —
    key alignment already isolates their growth."""
    if region.current.orientation != "block":
        return False
    for (base_row, base_col), (curr_row, curr_col) in region.cell_pairs():
        base_cell = base_sheet.cells.get((base_row, base_col))
        curr_cell = curr_sheet.cells.get((curr_row, curr_col))
        if base_cell is None or curr_cell is None:
            continue
        if _period_advanced(base_cell.value, curr_cell.value):
            return True
    return False


def diff_region_values(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    region: RegionAlignment,
    tolerance: NumericTolerance,
    *,
    ignore: _RangeSet,
    refresh: _RangeSet,
    sheet_profile: SheetProfile | None,
) -> list[Finding]:
    findings: list[Finding] = []
    sheet_name = curr_sheet.name
    refresh_block = _is_refresh_block(base_sheet, curr_sheet, region)
    for (base_row, base_col), (curr_row, curr_col) in region.cell_pairs():
        base_cell = base_sheet.cells.get((base_row, base_col))
        curr_cell = curr_sheet.cells.get((curr_row, curr_col))
        if base_cell is None and curr_cell is None:
            continue
        if (curr_row, curr_col) in ignore:
            continue
        location = _ref(curr_row, curr_col)
        baseline_location = _ref(base_row, base_col)

        # Formula-bearing cells are the formula engine's responsibility;
        # value comparison here covers constants and cached results.
        base_is_formula = base_cell is not None and base_cell.formula is not None
        curr_is_formula = curr_cell is not None and curr_cell.formula is not None

        if (
            not base_is_formula
            and not curr_is_formula
            and _values_differ(base_cell, curr_cell, tolerance)
            and not (
                _is_blank(curr_cell)
                and excel_blank_allowed(
                    curr_sheet,
                    sheet_profile,
                    curr_row,
                    curr_col,
                )
            )
        ):
            base_value = base_cell.value if base_cell else None
            curr_value = curr_cell.value if curr_cell else None
            if _is_blank(curr_cell):
                expected = False
                reason = " (required value is blank)"
            elif _period_advanced(base_value, curr_value):
                expected = True
                reason = " (period label advanced with the new cycle)"
            elif refresh_block:
                expected = True
                reason = " (cycle-snapshot block refresh)"
            elif (curr_row, curr_col) in refresh:
                expected = True
                reason = " (profile refresh range)"
            else:
                expected = False
                reason = ""
            findings.append(
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.VALUE_CHANGED,
                    expected_growth=expected,
                    sheet=sheet_name,
                    location=location,
                    baseline_location=baseline_location,
                    baseline_value=_display(base_value),
                    current_value=_display(curr_value),
                    message=f"{sheet_name}!{location}: historical value changed{reason}",
                )
            )

        if base_cell is not None and curr_cell is not None:
            if (
                base_cell.number_format is not None
                and curr_cell.number_format is not None
                and base_cell.number_format != curr_cell.number_format
            ):
                findings.append(
                    Finding(
                        artifact="excel",
                        finding_class=FindingClass.NUMBER_FORMAT_CHANGED,
                        sheet=sheet_name,
                        location=location,
                        baseline_location=baseline_location,
                        baseline_value=base_cell.number_format,
                        current_value=curr_cell.number_format,
                        message=f"{sheet_name}!{location}: number format changed",
                    )
                )
            if (
                base_cell.style_key is not None
                and curr_cell.style_key is not None
                and base_cell.style_key != curr_cell.style_key
            ):
                findings.append(
                    Finding(
                        artifact="excel",
                        finding_class=FindingClass.STYLE_CHANGED,
                        sheet=sheet_name,
                        location=location,
                        baseline_location=baseline_location,
                        baseline_value=base_cell.style_key,
                        current_value=curr_cell.style_key,
                        message=f"{sheet_name}!{location}: cell style changed",
                    )
                )
    return findings


def _axis_findings(
    sheet_name: str, axis: AxisAlignment, *, is_rows: bool, region_id: str
) -> list[Finding]:
    findings: list[Finding] = []

    def span(index: int) -> str:
        return f"row {index}" if is_rows else f"column {get_column_letter(index)}"

    for index in axis.deleted:
        findings.append(
            Finding(
                artifact="excel",
                finding_class=(
                    FindingClass.ROW_DELETED if is_rows else FindingClass.COLUMN_DELETED
                ),
                sheet=sheet_name,
                baseline_location=span(index),
                message=f"{sheet_name} ({region_id}): historical {span(index)} deleted",
            )
        )
    for index in axis.inserted:
        findings.append(
            Finding(
                artifact="excel",
                finding_class=(
                    FindingClass.ROW_INSERTED if is_rows else FindingClass.COLUMN_INSERTED
                ),
                sheet=sheet_name,
                location=span(index),
                message=(
                    f"{sheet_name} ({region_id}): unexpected {span(index)} inserted "
                    "outside the cadence growth pattern"
                ),
            )
        )
    for index in axis.growth:
        findings.append(
            Finding(
                artifact="excel",
                finding_class=(
                    FindingClass.ROW_GROWTH if is_rows else FindingClass.COLUMN_GROWTH
                ),
                expected_growth=True,
                sheet=sheet_name,
                location=span(index),
                message=f"{sheet_name} ({region_id}): new-cycle {span(index)} appended",
            )
        )
    return findings


def diff_workbook_values(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    alignment: WorkbookAlignment,
    profile: DeliverableProfile | None = None,
    *,
    cancellation_token: CancellationToken | None = None,
) -> list[Finding]:
    tolerance = profile.tolerance if profile else NumericTolerance()
    findings: list[Finding] = []
    for sheet_name, regions in alignment.regions.items():
        check_cancelled(cancellation_token)
        base_sheet = baseline.sheet(sheet_name)
        curr_sheet = current.sheet(sheet_name)
        sheet_profile = profile.sheet_profile(sheet_name) if profile else None
        ignore = _RangeSet(sheet_profile.ignore_ranges if sheet_profile else [])
        refresh = _RangeSet(sheet_profile.refresh_ranges if sheet_profile else [])
        for region in regions:
            check_cancelled(cancellation_token)
            if region.low_confidence:
                findings.append(
                    Finding(
                        artifact="excel",
                        finding_class=FindingClass.ALIGNMENT_LOW_CONFIDENCE,
                        sheet=sheet_name,
                        location=region.current.cell_range,
                        baseline_location=region.baseline.cell_range,
                        message=(
                            f"{sheet_name} ({region.current.region_id}): key matching "
                            "fell below 50%; cell-level value, format, style, and "
                            "formula comparison was skipped for this region"
                        ),
                    )
                )
                continue
            findings.extend(
                diff_region_values(
                    base_sheet,
                    curr_sheet,
                    region,
                    tolerance,
                    ignore=ignore,
                    refresh=refresh,
                    sheet_profile=sheet_profile,
                )
            )
            region_id = region.current.region_id
            findings.extend(
                _axis_findings(sheet_name, region.rows, is_rows=True, region_id=region_id)
            )
            findings.extend(
                _axis_findings(sheet_name, region.columns, is_rows=False, region_id=region_id)
            )
    return findings
