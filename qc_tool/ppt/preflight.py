"""Intrinsic QC for one current PowerPoint deck."""

import re
from collections import Counter
from dataclasses import dataclass, field

from qc_tool.availability import (
    availability_coverage,
    ppt_availability_issues,
    ppt_blank_allowed,
)
from qc_tool.config.profile import PptProfile
from qc_tool.coverage import CoverageItem, CoverageState
from qc_tool.excel.periods import Period, parse_period
from qc_tool.findings import Finding, FindingClass
from qc_tool.ppt.extract import DeckSnapshot

_PERIOD_PATTERNS = (
    re.compile(
        r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[A-Za-z]*"
        r"[-_ ]?'?\d{2,4}\b",
        re.IGNORECASE,
    ),
    re.compile(r"\b\d{4}-\d{2}(?:-\d{2})?\b"),
    re.compile(r"\b(?:CW|WK|W)[-_ ]?\d{1,2}(?:[-_ ]?'?\d{2,4})?\b", re.IGNORECASE),
    re.compile(r"\bQ[1-4](?:[-_ ]?(?:FY)?[-_ ]?'?\d{2,4})?\b", re.IGNORECASE),
)
_CURRENT_PERIOD_CONTEXT = re.compile(
    r"\b(?:as of|reporting period|cycle|month ending|week ending|quarter ending)\b",
    re.IGNORECASE,
)


@dataclass(slots=True)
class PptPreflightResult:
    findings: list[Finding] = field(default_factory=list)
    coverage: list[CoverageItem] = field(default_factory=list)


def _periods_in_text(text: str) -> list[tuple[str, Period]]:
    periods: list[tuple[str, Period]] = []
    for pattern in _PERIOD_PATTERNS:
        for match in pattern.finditer(text):
            parsed = parse_period(match.group(0))
            if parsed is not None:
                periods.append((match.group(0), parsed))
    return periods


def contextual_periods(deck: DeckSnapshot) -> dict[tuple[str, tuple[int, int, int]], set[str]]:
    """Reporting-period labels used in explicit cycle/as-of text."""
    periods: dict[tuple[str, tuple[int, int, int]], set[str]] = {}
    for slide in deck.slides:
        for text in [slide.title or "", *slide.texts]:
            if not _CURRENT_PERIOD_CONTEXT.search(text):
                continue
            for label, period in _periods_in_text(text):
                periods.setdefault((period.kind, period.sort_key), set()).add(
                    f"{slide.display_name}: {label}"
                )
    return periods


def preflight_deck(deck: DeckSnapshot, profile: PptProfile) -> PptPreflightResult:
    result = PptPreflightResult()

    inventory_start = len(result.findings)
    titles = [slide.title.strip() for slide in deck.slides if slide.title and slide.title.strip()]
    title_counts = Counter(title.casefold() for title in titles)
    for title, count in title_counts.items():
        if count <= 1:
            continue
        result.findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_DUPLICATE_TITLE,
                slide=title,
                current_value=str(count),
                message=f"slide title {title!r} appears {count} times",
            )
        )
    available_titles = {title.casefold() for title in titles}
    for required in profile.required_slides:
        if required.casefold() in available_titles:
            continue
        result.findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_REQUIRED_SLIDE_MISSING,
                slide=required,
                message=f"required slide {required!r} is missing",
            )
        )
    token_patterns = {
        token: re.compile(rf"(?<!\w){re.escape(token)}(?!\w)", re.IGNORECASE)
        for token in profile.draft_tokens
        if token
    }
    for slide in deck.slides:
        if slide.shape_count == 0:
            result.findings.append(
                Finding(
                    artifact="ppt",
                    finding_class=FindingClass.PPT_EMPTY_SLIDE,
                    slide=slide.display_name,
                    location=f"slide {slide.index + 1}",
                    message=f"slide {slide.index + 1} is empty",
                )
            )
        content = [slide.title or "", *slide.texts]
        content.extend(cell for table in slide.tables for row in table.rows for cell in row)
        for text in content:
            for token, pattern in token_patterns.items():
                if not pattern.search(text):
                    continue
                result.findings.append(
                    Finding(
                        artifact="ppt",
                        finding_class=FindingClass.PPT_DRAFT_TOKEN,
                        slide=slide.display_name,
                        location=f"slide {slide.index + 1}",
                        current_value=text,
                        element=token,
                        message=f"{slide.display_name}: unfinished token {token!r} remains",
                    )
                )
    result.coverage.append(
        CoverageItem(
            check_id="ppt-slide-inventory",
            label="Slide inventory, required slides, and draft tokens",
            artifact="ppt",
            state=CoverageState.CHECKED,
            findings=len(result.findings) - inventory_start,
        )
    )

    content_start = len(result.findings)
    for slide in deck.slides:
        for table_index, table in enumerate(slide.tables, start=1):
            if len(table.rows) < 2:
                continue
            width = max((len(row) for row in table.rows), default=0)
            headers = table.rows[0]
            table_element = table.name or f"table[{table.source_index}]"
            for row_index, row in enumerate(table.rows[1:], start=2):
                for col_index in range(1, width):
                    value = row[col_index].strip() if col_index < len(row) else ""
                    row_has_data = any(cell.strip() for cell in row)
                    if value or not row_has_data:
                        continue
                    row_label = row[0] if row else ""
                    period_label = (
                        headers[col_index] if col_index < len(headers) else ""
                    )
                    if ppt_blank_allowed(
                        profile,
                        slide=slide.display_name,
                        scope="table",
                        element=table_element,
                        series=row_label,
                        period_label=period_label,
                    ):
                        continue
                    result.findings.append(
                        Finding(
                            artifact="ppt",
                            finding_class=FindingClass.PPT_TABLE_BLANK,
                            slide=slide.display_name,
                            location=f"slide {slide.index + 1}",
                            element=f"table {table_index}, row {row_index}, column {col_index + 1}",
                            message=(
                                f"{slide.display_name}: blank cell inside a populated "
                                f"table row ({row_index}, {col_index + 1})"
                            ),
                        )
                    )
        for chart_index, chart in enumerate(slide.charts, start=1):
            complete_series = chart.all_series
            if complete_series:
                series_items = [
                    (series.name, series.categories, series.values)
                    for series in complete_series
                ]
            else:
                series_items = [
                    (series_name, chart.categories, values)
                    for series_name, values in chart.series
                ]
            for series_name, categories, values in series_items:
                chart_element = chart.title or chart.name or f"chart[{chart.source_index}]"
                series_element = series_name or f"series {chart_index}"
                missing_positions = [
                    index
                    for index, _category in enumerate(categories)
                    if index >= len(values) or values[index] is None
                ]
                unallowed_positions = [
                    index
                    for index in missing_positions
                    if not ppt_blank_allowed(
                        profile,
                        slide=slide.display_name,
                        scope="chart",
                        element=chart_element,
                        series=series_element,
                        period_label=categories[index],
                    )
                ]
                if len(values) != len(categories) and (
                    len(values) > len(categories) or unallowed_positions
                ):
                    result.findings.append(
                        Finding(
                            artifact="ppt",
                            finding_class=FindingClass.PPT_CHART_LENGTH_MISMATCH,
                            slide=slide.display_name,
                            location=f"slide {slide.index + 1}",
                            element=series_element,
                            baseline_value=str(len(categories)),
                            current_value=str(len(values)),
                            message=(
                                f"{slide.display_name}: chart series has "
                                f"{len(values)} values for {len(categories)} categories"
                            ),
                        )
                    )
                for index in unallowed_positions:
                    result.findings.append(
                        Finding(
                            artifact="ppt",
                            finding_class=FindingClass.PPT_CHART_VALUE_MISSING,
                            slide=slide.display_name,
                            location=f"slide {slide.index + 1}",
                            element=categories[index],
                            message=(
                                f"{slide.display_name}: {series_element!r} has "
                                f"no value for {categories[index]!r}"
                            ),
                        )
                    )
    result.coverage.append(
        CoverageItem(
            check_id="ppt-tables-charts",
            label="Table completeness and chart dimensions",
            artifact="ppt",
            state=(
                CoverageState.CHECKED
                if deck.charts_available
                else CoverageState.DEGRADED
            ),
            findings=len(result.findings) - content_start,
            detail="" if deck.charts_available else deck.chart_detail,
        )
    )
    result.coverage.append(
        availability_coverage(
            artifact="ppt",
            rule_count=len(profile.availability_rules),
            issues=ppt_availability_issues(deck, profile),
        )
    )

    period_start = len(result.findings)
    deck_periods = contextual_periods(deck)
    if len(deck_periods) > 1:
        labels = sorted(item for sources in deck_periods.values() for item in sources)
        result.findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_PERIOD_INCONSISTENT,
                element="reporting period",
                current_value="; ".join(labels),
                message="deck contains conflicting contextual reporting periods",
            )
        )
    result.coverage.append(
        CoverageItem(
            check_id="ppt-period-consistency",
            label="Contextual reporting-period consistency",
            artifact="ppt",
            state=CoverageState.CHECKED,
            findings=len(result.findings) - period_start,
        )
    )
    result.coverage.append(
        CoverageItem(
            check_id="ppt-notes",
            label="Speaker notes extraction",
            artifact="ppt",
            state=(
                CoverageState.CHECKED
                if deck.notes_available
                else CoverageState.DEGRADED
            ),
            detail=deck.notes_detail,
        )
    )
    result.coverage.append(
        CoverageItem(
            check_id="ppt-media-visual",
            label="Media integrity and rendered visual layout",
            artifact="ppt",
            state=CoverageState.UNAVAILABLE,
            detail="Requires a trusted rendering comparison",
        )
    )
    return result
