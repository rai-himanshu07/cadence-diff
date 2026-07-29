"""Profile-driven blank availability boundaries shared by Excel and PowerPoint."""

from __future__ import annotations

from openpyxl.utils.cell import range_boundaries

from qc_tool.config.profile import DeliverableProfile, PptProfile, SheetProfile
from qc_tool.coverage import CoverageItem, CoverageState
from qc_tool.excel.periods import is_period_after, parse_period
from qc_tool.io.model import SheetSnapshot, WorkbookSnapshot
from qc_tool.ppt.model import DeckSnapshot


def _bounds(cell_range: str) -> tuple[int, int, int, int] | None:
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    except ValueError:
        return None
    if min_col is None or min_row is None or max_col is None or max_row is None:
        return None
    return min_col, min_row, max_col, max_row


def cell_in_ranges(row: int, column: int, ranges: list[str]) -> bool:
    """Whether a cell belongs to any configured A1 range."""
    for cell_range in ranges:
        bounds = _bounds(cell_range)
        if bounds is None:
            continue
        min_col, min_row, max_col, max_row = bounds
        if min_row <= row <= max_row and min_col <= column <= max_col:
            return True
    return False


def _excel_period_value(
    sheet: SheetSnapshot,
    target_range: str,
    period_range: str,
    row: int,
    column: int,
) -> object | None:
    target = _bounds(target_range)
    periods = _bounds(period_range)
    if target is None or periods is None:
        return None
    target_min_col, target_min_row, target_max_col, target_max_row = target
    if not (
        target_min_row <= row <= target_max_row
        and target_min_col <= column <= target_max_col
    ):
        return None
    period_min_col, period_min_row, period_max_col, period_max_row = periods
    period_width = period_max_col - period_min_col + 1
    period_height = period_max_row - period_min_row + 1
    target_width = target_max_col - target_min_col + 1
    target_height = target_max_row - target_min_row + 1
    if period_width == period_height == 1:
        period_row, period_column = period_min_row, period_min_col
    elif period_height == 1 and period_width == target_width:
        period_row = period_min_row
        period_column = period_min_col + (column - target_min_col)
    elif period_width == 1 and period_height == target_height:
        period_row = period_min_row + (row - target_min_row)
        period_column = period_min_col
    else:
        return None
    cell = sheet.cells.get((period_row, period_column))
    return None if cell is None else cell.value


def excel_blank_allowed(
    sheet: SheetSnapshot,
    profile: SheetProfile | None,
    row: int,
    column: int,
) -> bool:
    """Whether every applicable resolved Excel rule permits this blank."""
    if profile is None:
        return False
    decisions: list[bool] = []
    for rule in profile.availability_rules:
        period_value = _excel_period_value(
            sheet,
            rule.cell_range,
            rule.period_range,
            row,
            column,
        )
        if period_value is None:
            continue
        period = parse_period(period_value)
        required = parse_period(rule.required_through)
        if period is None or required is None or period.kind != required.kind:
            decisions.append(False)
            continue
        decisions.append(
            bool(rule.allow_blank_after and is_period_after(period, required))
        )
    return bool(decisions) and all(decisions)


def _matches(value: str, configured: str) -> bool:
    return not configured or value.casefold() == configured.casefold()


def ppt_blank_allowed(
    profile: PptProfile,
    *,
    slide: str,
    scope: str,
    element: str,
    series: str,
    period_label: str,
) -> bool:
    """Whether every matching PowerPoint rule permits this future blank."""
    period = parse_period(period_label)
    if period is None:
        return False
    decisions: list[bool] = []
    for rule in profile.availability_rules:
        if (
            rule.scope != scope
            or rule.slide.casefold() != slide.casefold()
            or not _matches(element, rule.element)
            or not _matches(series, rule.series or "")
        ):
            continue
        required = parse_period(rule.required_through)
        if required is None or required.kind != period.kind:
            decisions.append(False)
            continue
        decisions.append(
            bool(rule.allow_blank_after and is_period_after(period, required))
        )
    return bool(decisions) and all(decisions)


def availability_coverage(
    *,
    artifact: str,
    rule_count: int,
    issues: list[str] | None = None,
) -> CoverageItem:
    """Disclose configured boundaries or the strict no-inference default."""
    issues = issues or []
    return CoverageItem(
        check_id=f"{artifact}-availability",
        label="Availability boundaries",
        artifact=artifact,
        state=CoverageState.DEGRADED if issues else CoverageState.CHECKED,
        detail=(
            (
                f"{rule_count} availability rules configured; "
                f"{len(issues)} unresolved: {'; '.join(issues)}"
            )
            if issues
            else f"{rule_count} explicit availability rules applied"
            if rule_count
            else "No availability rules configured; strict blank requirements applied"
        ),
    )


def excel_availability_issues(
    workbook: WorkbookSnapshot,
    profile: DeliverableProfile,
) -> list[str]:
    """Runtime resolution failures for Excel availability rules."""
    issues: list[str] = []
    for sheet_name, sheet_profile in profile.excel.sheets.items():
        if sheet_name in profile.excel.ignore_sheets or sheet_profile.ignore:
            continue
        if not sheet_profile.availability_rules:
            continue
        try:
            sheet = workbook.sheet(sheet_name)
        except KeyError:
            issues.append(f"sheet {sheet_name!r} not found")
            continue
        for index, rule in enumerate(sheet_profile.availability_rules):
            label = rule.name or str(index)
            target = _bounds(rule.cell_range)
            periods = _bounds(rule.period_range)
            required = parse_period(rule.required_through)
            if target is None or periods is None or required is None:
                issues.append(f"{sheet_name}/{label} has invalid ranges or period")
                continue
            target_min_col, target_min_row, target_max_col, target_max_row = target
            period_min_col, period_min_row, period_max_col, period_max_row = periods
            target_width = target_max_col - target_min_col + 1
            target_height = target_max_row - target_min_row + 1
            period_width = period_max_col - period_min_col + 1
            period_height = period_max_row - period_min_row + 1
            aligned = (
                period_width == period_height == 1
                or (period_height == 1 and period_width == target_width)
                or (period_width == 1 and period_height == target_height)
            )
            if not aligned:
                issues.append(f"{sheet_name}/{label} period and target ranges misalign")
                continue
            observed = []
            for row in range(period_min_row, period_max_row + 1):
                for column in range(period_min_col, period_max_col + 1):
                    cell = sheet.cells.get((row, column))
                    if cell is not None and cell.value not in {None, ""}:
                        observed.append(parse_period(cell.value))
            if not observed:
                issues.append(f"{sheet_name}/{label} resolves no period labels")
            elif any(
                period is None or period.kind != required.kind
                for period in observed
            ):
                issues.append(
                    f"{sheet_name}/{label} period labels do not match "
                    f"{required.kind!r} required_through"
                )
    return sorted(set(issues))


def ppt_availability_issues(
    deck: DeckSnapshot,
    profile: PptProfile,
) -> list[str]:
    """Runtime resolution failures for PowerPoint availability rules."""
    issues: list[str] = []
    slides = {slide.display_name.casefold(): slide for slide in deck.slides}
    for index, rule in enumerate(profile.availability_rules):
        label = rule.name or str(index)
        if parse_period(rule.required_through) is None:
            issues.append(f"{label} has invalid required_through")
            continue
        slide = slides.get(rule.slide.casefold())
        if slide is None:
            issues.append(f"{label} slide {rule.slide!r} not found")
            continue
        if rule.scope == "table":
            candidates = [
                (
                    table.name or f"table[{table.source_index}]",
                    {
                        row[0].casefold()
                        for row in table.rows[1:]
                        if row and row[0]
                    },
                )
                for table in slide.tables
            ]
        else:
            candidates = [
                (
                    chart.title or chart.name or f"chart[{chart.source_index}]",
                    {
                        series.name.casefold()
                        for series in chart.all_series
                        if series.name
                    },
                )
                for chart in slide.charts
            ]
        matching = [
            (element, names)
            for element, names in candidates
            if not rule.element or element.casefold() == rule.element.casefold()
        ]
        if rule.element and not matching:
            issues.append(f"{label} element {rule.element!r} not found")
        elif rule.series and not any(
            rule.series.casefold() in names for _, names in matching
        ):
            issues.append(f"{label} series/row {rule.series!r} not found")
    return sorted(set(issues))
