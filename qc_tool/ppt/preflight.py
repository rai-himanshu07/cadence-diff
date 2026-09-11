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
from qc_tool.findings import Finding, FindingClass
from qc_tool.ppt.claim_periods import CURRENT_PERIOD_CONTEXT, periods_in_text
from qc_tool.ppt.extract import DeckSnapshot
from qc_tool.ppt.repetition import check_repeated_claims


@dataclass(slots=True)
class PptPreflightResult:
    findings: list[Finding] = field(default_factory=list)
    coverage: list[CoverageItem] = field(default_factory=list)


def media_structural_coverage(
    *decks: DeckSnapshot,
    findings: int = 0,
    ambiguous_shapes: int = 0,
) -> CoverageItem:
    if not decks:
        return CoverageItem(
            check_id="ppt-media-structural",
            label="Embedded media byte structure",
            artifact="ppt",
            state=CoverageState.NOT_INCLUDED,
            detail="PowerPoint deck not supplied",
        )
    available = all(deck.media_available for deck in decks)
    complete = available and ambiguous_shapes == 0
    details = [
        deck.media_detail for deck in decks if not deck.media_available
    ]
    if ambiguous_shapes:
        details.append(
            f"{ambiguous_shapes} media shape(s) had ambiguous name/geometry "
            "keys and were not compared"
        )
    return CoverageItem(
        check_id="ppt-media-structural",
        label="Embedded media byte structure",
        artifact="ppt",
        state=CoverageState.CHECKED if complete else CoverageState.DEGRADED,
        findings=findings,
        detail=(
            "Embedded media bytes hashed without decoding"
            if complete
            else "; ".join(details)
        ),
    )


def contextual_periods(deck: DeckSnapshot) -> dict[tuple[str, tuple[int, int, int]], set[str]]:
    """Reporting-period labels used in explicit cycle/as-of text."""
    periods: dict[tuple[str, tuple[int, int, int]], set[str]] = {}
    for slide in deck.slides:
        for text in [slide.title or "", *slide.texts]:
            if not CURRENT_PERIOD_CONTEXT.search(text):
                continue
            for label, period in periods_in_text(text):
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
                    slide_index=slide.index + 1,
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
                        slide_index=slide.index + 1,
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

    repetition = check_repeated_claims(deck)
    result.findings.extend(repetition.findings)
    result.coverage.append(repetition.coverage)

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
                            slide_index=slide.index + 1,
                            location=f"slide {slide.index + 1}",
                            focus_shape_id=table.shape_id or None,
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
                            slide_index=slide.index + 1,
                            location=f"slide {slide.index + 1}",
                            focus_shape_id=chart.shape_id or None,
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
                            slide_index=slide.index + 1,
                            location=f"slide {slide.index + 1}",
                            focus_shape_id=chart.shape_id or None,
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
        media_structural_coverage(deck)
    )
    result.coverage.append(
        CoverageItem(
            check_id="ppt-media-visual",
            label="Rendered media and visual layout",
            artifact="ppt",
            state=CoverageState.UNAVAILABLE,
            detail=(
                "Embedded bytes are checked structurally; pixels, OCR text, and "
                "rendered layout are not inspected"
            ),
        )
    )
    return result
