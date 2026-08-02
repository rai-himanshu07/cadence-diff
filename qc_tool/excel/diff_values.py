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
from qc_tool.config.profile import (
    AcceptanceBand,
    DeliverableProfile,
    NumericTolerance,
    RestatementWindows,
    SheetProfile,
)
from qc_tool.excel.align import AxisAlignment, RegionAlignment, WorkbookAlignment
from qc_tool.excel.materiality import (
    classify_numeric_pair,
    is_anomalous_magnitude,
    numeric_evidence_tags,
    temporal_contexts,
)
from qc_tool.excel.periods import is_period_after, parse_period
from qc_tool.excel.regions import TableRegion, period_positions
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingExpectedReason,
    FindingSubtype,
    FindingTemporalContext,
)
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


class _AcceptanceBands:
    """Declared per-range tolerance bands plus the run-level analyst band."""

    def __init__(
        self,
        bands: Iterable[AcceptanceBand],
        run_acceptance: NumericTolerance | None = None,
    ) -> None:
        self._bands = [(_RangeSet([band.cell_range]), band) for band in bands]
        self._run_acceptance = run_acceptance

    @staticmethod
    def _within(absolute: float, relative: float, base: float, curr: float) -> bool:
        delta = abs(curr - base)
        within_abs = absolute > 0 and delta <= absolute
        within_rel = relative > 0 and base != 0 and delta / abs(base) <= relative
        return within_abs or within_rel

    def accepts(self, cell: tuple[int, int], base: object, curr: object) -> bool:
        if not self._bands and self._run_acceptance is None:
            return False
        if isinstance(base, bool) or isinstance(curr, bool):
            return False
        if not isinstance(base, int | float) or not isinstance(curr, int | float):
            return False
        base_float, curr_float = float(base), float(curr)
        if self._run_acceptance is not None and self._within(
            self._run_acceptance.absolute,
            self._run_acceptance.relative,
            base_float,
            curr_float,
        ):
            return True
        for range_set, band in self._bands:
            if cell not in range_set:
                continue
            if self._within(band.absolute, band.relative, base_float, curr_float):
                return True
        return False


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


def _both_numeric(base_value: object, curr_value: object) -> bool:
    return (
        isinstance(base_value, int | float)
        and isinstance(curr_value, int | float)
        and not isinstance(base_value, bool)
        and not isinstance(curr_value, bool)
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
    windows: RestatementWindows | None = None,
    run_acceptance: NumericTolerance | None = None,
) -> list[Finding]:
    findings: list[Finding] = []
    sheet_name = curr_sheet.name
    refresh_block = _is_refresh_block(base_sheet, curr_sheet, region)
    recency_windows = windows if windows is not None else RestatementWindows()
    acceptance = _AcceptanceBands(
        sheet_profile.acceptance_bands if sheet_profile is not None else [],
        run_acceptance,
    )
    temporal_by_position: dict[int, FindingTemporalContext] = {}
    period_axis = region.current.period_axis
    if period_axis in ("rows", "columns"):
        is_rows_axis = period_axis == "rows"
        populated = _data_positions(
            curr_sheet, region.current, is_rows=is_rows_axis
        )
        positions = {
            position: period
            for position, period in period_positions(
                curr_sheet, region.current
            ).items()
            if position in populated
        }
        temporal_by_position = temporal_contexts(positions, recency_windows)
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
            base_blank = _is_blank(base_cell)
            curr_blank = _is_blank(curr_cell)
            if curr_blank and not base_blank:
                subtype = FindingSubtype.VALUE_CLEARED_POPULATION
                wording = "historical value cleared"
            elif base_blank and not curr_blank:
                subtype = FindingSubtype.VALUE_ADDED_POPULATION
                wording = "value added to a previously blank cell"
            else:
                subtype = FindingSubtype.VALUE_REPLACEMENT
                wording = "historical value changed"
            materiality = None
            temporal_context = None
            evidence_tags = set()
            if subtype is FindingSubtype.VALUE_REPLACEMENT:
                display_format = (
                    curr_cell.number_format if curr_cell else None
                ) or (base_cell.number_format if base_cell else None)
                position = curr_row if period_axis == "rows" else curr_col
                temporal_context = temporal_by_position.get(position)
                materiality = classify_numeric_pair(
                    base_value,
                    curr_value,
                    display_format,
                    within_acceptance=acceptance.accepts(
                        (curr_row, curr_col), base_value, curr_value
                    ),
                )
                evidence_tags = numeric_evidence_tags(
                    base_value,
                    curr_value,
                    display_format,
                )
            if curr_blank:
                expected_reason = None
                reason = " (required value is blank)"
            elif _period_advanced(base_value, curr_value):
                expected_reason = FindingExpectedReason.PERIOD_PROGRESSION
                reason = " (period label advanced with the new cycle)"
            elif (curr_row, curr_col) in refresh:
                expected_reason = FindingExpectedReason.PROFILE_REFRESH
                reason = " (profile refresh range)"
            elif refresh_block and _both_numeric(base_value, curr_value):
                expected_reason = None
                if is_anomalous_magnitude(base_value, curr_value):
                    reason = " (anomalous change in cycle-snapshot block)"
                else:
                    temporal_context = FindingTemporalContext.CURRENT_PERIOD
                    reason = " (implicit cycle-snapshot block refresh)"
            else:
                expected_reason = None
                reason = ""
            findings.append(
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.VALUE_CHANGED,
                    expected_reason=expected_reason,
                    subtype=subtype,
                    materiality=materiality,
                    temporal_context=temporal_context,
                    evidence_tags=evidence_tags,
                    sheet=sheet_name,
                    location=location,
                    baseline_location=baseline_location,
                    baseline_value=_display(base_value),
                    current_value=_display(curr_value),
                    message=f"{sheet_name}!{location}: {wording}{reason}",
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


_AXIS_WORDING = {
    FindingSubtype.AXIS_ROLLING_TURNOVER: "oldest {span} left the rolling window",
    FindingSubtype.AXIS_PHYSICAL_DELETION: "historical {span} deleted",
    FindingSubtype.AXIS_PHYSICAL_INSERTION: (
        "unexpected {span} inserted outside the cadence growth pattern"
    ),
    FindingSubtype.AXIS_EXTENT_GROWTH: "new-cycle {span} appended",
}


def _axis_events(axis: AxisAlignment) -> tuple[
    dict[int, FindingSubtype],
    dict[int, FindingSubtype],
    FindingSubtype,
    list[int],
]:
    """Per-index deleted/inserted events, the shared growth event, and the
    in-place key changes.

    A position that is both deleted and inserted changed its key in place; it
    is one key-change event, never a physical row or column edit pair.
    """
    deleted = set(axis.deleted)
    inserted = set(axis.inserted)
    reused = deleted & inserted
    sliding = bool(axis.growth) and not inserted
    deleted_events = {
        index: (
            FindingSubtype.AXIS_ROLLING_TURNOVER
            if sliding
            else FindingSubtype.AXIS_PHYSICAL_DELETION
        )
        for index in axis.deleted
        if index not in reused
    }
    inserted_events = {
        index: FindingSubtype.AXIS_PHYSICAL_INSERTION
        for index in axis.inserted
        if index not in reused
    }
    growth_event = (
        FindingSubtype.AXIS_ROLLING_TURNOVER
        if sliding and deleted
        else FindingSubtype.AXIS_EXTENT_GROWTH
    )
    return deleted_events, inserted_events, growth_event, sorted(reused)


def _key_cell(
    sheet: SheetSnapshot, region_table: TableRegion, index: int, *, is_rows: bool
) -> CellRecord | None:
    if is_rows:
        key_col = region_table.key_col
        return sheet.cells.get((index, key_col)) if key_col is not None else None
    header_row = region_table.header_row
    return sheet.cells.get((header_row, index)) if header_row is not None else None


def _data_positions(
    sheet: SheetSnapshot, region_table: TableRegion, *, is_rows: bool
) -> set[int]:
    """Axis positions holding at least one populated non-label cell.

    Trackers often pre-fill their period calendar years ahead; the
    restatement window must trail the last period WITH data, not the last
    printed label, so empty future positions never define the edge.
    """
    positions: set[int] = set()
    for (row, col), cell in sheet.cells.items():
        if cell.value is None:
            continue
        if not (
            region_table.min_row <= row <= region_table.max_row
            and region_table.min_col <= col <= region_table.max_col
        ):
            continue
        if is_rows:
            if col == region_table.key_col:
                continue
            positions.add(row)
        else:
            if row == region_table.header_row:
                continue
            positions.add(col)
    return positions


def _axis_findings(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    region: RegionAlignment,
    axis: AxisAlignment,
    *,
    is_rows: bool,
    region_id: str,
) -> list[Finding]:
    findings: list[Finding] = []
    sheet_name = curr_sheet.name
    if not (axis.deleted or axis.inserted or axis.growth):
        return findings
    deleted_events, inserted_events, growth_event, key_changes = _axis_events(axis)
    event_key = f"excel:{region_id}:{'rows' if is_rows else 'columns'}"

    def span(index: int) -> str:
        return f"row {index}" if is_rows else f"column {get_column_letter(index)}"

    def message(subtype: FindingSubtype, index: int) -> str:
        return (
            f"{sheet_name} ({region_id}): "
            + _AXIS_WORDING[subtype].format(span=span(index))
        )

    for index in key_changes:
        base_key = _key_cell(base_sheet, region.baseline, index, is_rows=is_rows)
        curr_key = _key_cell(curr_sheet, region.current, index, is_rows=is_rows)
        derived = (
            base_key is not None
            and curr_key is not None
            and base_key.has_formula
            and curr_key.has_formula
        )
        advanced = (
            base_key is not None
            and curr_key is not None
            and _period_advanced(base_key.value, curr_key.value)
        )
        subtype = (
            FindingSubtype.AXIS_KEY_DERIVED_LABEL
            if derived
            else FindingSubtype.AXIS_KEY_REPLACEMENT
        )
        base_label = display_cell_value(base_key.value) if base_key else ""
        curr_label = display_cell_value(curr_key.value) if curr_key else ""
        note = ""
        if advanced:
            note = " (tracking period advanced with the new cycle)"
        elif derived and curr_key is not None and curr_key.formula:
            note = (
                " [label is formula-derived ("
                + (curr_key.formula[:80])
                + "); the upstream driver changed, not this position]"
            )
        findings.append(
            Finding(
                artifact="excel",
                finding_class=(
                    FindingClass.ROW_KEY_CHANGED
                    if is_rows
                    else FindingClass.COLUMN_KEY_CHANGED
                ),
                expected_reason=(
                    FindingExpectedReason.PERIOD_PROGRESSION if advanced else None
                ),
                subtype=subtype,
                event_key=event_key,
                sheet=sheet_name,
                location=span(index),
                baseline_location=span(index),
                baseline_value=base_label,
                current_value=curr_label,
                message=(
                    f"{sheet_name} ({region_id}): historical key at {span(index)} "
                    f"replaced in place: {base_label!r} -> {curr_label!r}" + note
                ),
            )
        )
    for index in axis.deleted:
        if index not in deleted_events:
            continue
        subtype = deleted_events[index]
        findings.append(
            Finding(
                artifact="excel",
                finding_class=(
                    FindingClass.ROW_DELETED if is_rows else FindingClass.COLUMN_DELETED
                ),
                subtype=subtype,
                event_key=event_key,
                sheet=sheet_name,
                baseline_location=span(index),
                message=message(subtype, index),
            )
        )
    for index in axis.inserted:
        if index not in inserted_events:
            continue
        subtype = inserted_events[index]
        findings.append(
            Finding(
                artifact="excel",
                finding_class=(
                    FindingClass.ROW_INSERTED if is_rows else FindingClass.COLUMN_INSERTED
                ),
                subtype=subtype,
                event_key=event_key,
                sheet=sheet_name,
                location=span(index),
                message=message(subtype, index),
            )
        )
    for index in axis.growth:
        findings.append(
            Finding(
                artifact="excel",
                finding_class=(
                    FindingClass.ROW_GROWTH if is_rows else FindingClass.COLUMN_GROWTH
                ),
                expected_reason=(
                    FindingExpectedReason.ROLLING_WINDOW
                    if growth_event is FindingSubtype.AXIS_ROLLING_TURNOVER
                    else FindingExpectedReason.CADENCE_EXTENSION
                ),
                subtype=growth_event,
                event_key=event_key,
                sheet=sheet_name,
                location=span(index),
                message=(
                    f"{sheet_name} ({region_id}): new-cycle {span(index)} "
                    + (
                        "entered the rolling window"
                        if growth_event is FindingSubtype.AXIS_ROLLING_TURNOVER
                        else "appended"
                    )
                ),
            )
        )
    return findings


def diff_workbook_values(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    alignment: WorkbookAlignment,
    profile: DeliverableProfile | None = None,
    *,
    run_acceptance: NumericTolerance | None = None,
    cancellation_token: CancellationToken | None = None,
) -> list[Finding]:
    tolerance = profile.tolerance if profile else NumericTolerance()
    windows = profile.restatement_windows if profile else RestatementWindows()
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
                    windows=windows,
                    run_acceptance=run_acceptance,
                )
            )
            region_id = region.current.region_id
            findings.extend(
                _axis_findings(
                    base_sheet,
                    curr_sheet,
                    region,
                    region.rows,
                    is_rows=True,
                    region_id=region_id,
                )
            )
            findings.extend(
                _axis_findings(
                    base_sheet,
                    curr_sheet,
                    region,
                    region.columns,
                    is_rows=False,
                    region_id=region_id,
                )
            )
    return findings
