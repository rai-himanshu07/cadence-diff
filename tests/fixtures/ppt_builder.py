"""Builds the baseline/current PowerPoint fixture pair with seeded defects.

Deck design:

- "Executive Summary" — bullets carrying the cross-check figures.
- "Revenue by Region" — table (regions x months); values stay consistent
  with the *baseline* workbook so the only baseline-vs-current deck diffs
  are the seeded P04 edit and the expected new-month column.
- "Revenue Trend" — full-history native line chart (P05 seeds one edited
  historical point; the appended Jun-26 point is expected growth).
- "Weekly Ops" — rolling-window native bar chart (window 4, shifts one
  week per cycle; entirely expected — a negative control for the
  rolling-window rule).
- "Notes & Definitions" — identical content, but *reordered* in current
  (fuzzy matching must still pair it).
- "Deep Dive Archive" — baseline only. "New Initiatives" — current only.
"""

import datetime as dt
from collections.abc import Iterable
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.presentation import Presentation as PresentationType
from pptx.slide import Slide
from pptx.util import Inches, Pt

from tests.fixtures import domain
from tests.fixtures.manifest_schema import (
    Artifact,
    CrosscheckEntry,
    DefectClass,
    ExpectedChange,
    SeededDefect,
)

FIXED_DOC_TIME = dt.datetime(2026, 7, 1, tzinfo=dt.UTC)

TITLE_ONLY_LAYOUT = 5


def _add_slide(prs: PresentationType, title: str) -> Slide:
    slide = prs.slides.add_slide(prs.slide_layouts[TITLE_ONLY_LAYOUT])
    title_shape = slide.shapes.title
    if title_shape is None:  # pragma: no cover - layout 5 always has a title
        raise RuntimeError("layout has no title placeholder")
    title_shape.text = title
    return slide


def _add_bullets(slide: Slide, lines: Iterable[str]) -> None:
    box = slide.shapes.add_textbox(Inches(0.6), Inches(1.6), Inches(8.8), Inches(4.5))
    frame = box.text_frame
    frame.word_wrap = True
    for index, line in enumerate(lines):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = line
        paragraph.font.size = Pt(18)


def build_grouped_text_deck(
    path: Path,
    *,
    grouped_lines: tuple[str, ...] = (),
    nested_lines: tuple[str, ...] = (),
    top_before: tuple[str, ...] = (),
    top_after: tuple[str, ...] = (),
    title: str = "Grouped KPIs",
) -> Path:
    """Write a synthetic nested-group deck for extraction contracts."""
    presentation = Presentation()
    slide = _add_slide(presentation, title)

    def add_line(shapes, text: str, offset: int) -> None:
        box = shapes.add_textbox(
            Inches(0.6 + offset * 0.1),
            Inches(1.4 + offset * 0.45),
            Inches(6.0),
            Inches(0.4),
        )
        box.text = text

    offset = 0
    for text in top_before:
        add_line(slide.shapes, text, offset)
        offset += 1
    group = slide.shapes.add_group_shape()
    for text in grouped_lines:
        add_line(group.shapes, text, offset)
        offset += 1
    if nested_lines:
        nested = group.shapes.add_group_shape()
        for text in nested_lines:
            add_line(nested.shapes, text, offset)
            offset += 1
    for text in top_after:
        add_line(slide.shapes, text, offset)
        offset += 1
    presentation.core_properties.created = FIXED_DOC_TIME
    presentation.core_properties.modified = FIXED_DOC_TIME
    presentation.save(str(path))
    return path


def _trend_values(*, current: bool) -> list[float]:
    months = domain.CURRENT_MONTHS if current else domain.BASELINE_MONTHS
    values = [domain.month_total_revenue(m, current=False) for m in range(months)]
    if current:
        values[2] += domain.P05_DELTA  # P05: historical chart point edited
    return values


def _exec_summary_lines(*, current: bool) -> list[str]:
    revenue = domain.total_revenue(current=current)
    margin = domain.margin_ratio(current=current)
    if current:
        margin += domain.X03_MARGIN_OFFSET  # X03: figure drifts from workbook
    outlook = "volatile" if current else "stable"  # P03 when current
    return [
        f"Total revenue {domain.fmt_millions(revenue)}",
        f"Margin {domain.fmt_pct(margin)}",
        f"Outlook remains {outlook} across all regions.",
    ]


def _slide_exec_summary(prs: PresentationType, *, current: bool) -> None:
    slide = _add_slide(prs, "Executive Summary")
    _add_bullets(slide, _exec_summary_lines(current=current))


def _slide_region_table(prs: PresentationType, *, current: bool) -> None:
    slide = _add_slide(prs, "Revenue by Region")
    months = domain.CURRENT_MONTHS if current else domain.BASELINE_MONTHS
    n_regions = len(domain.REGION_LABELS)
    shape = slide.shapes.add_table(
        n_regions + 1, months + 1, Inches(0.5), Inches(1.6), Inches(9.0), Inches(3.2)
    )
    table = shape.table
    table.cell(0, 0).text = "Region"
    for month in range(months):
        table.cell(0, month + 1).text = domain.MONTH_LABELS[month]
    for region_idx, region in enumerate(domain.REGION_LABELS):
        table.cell(region_idx + 1, 0).text = region
        for month in range(months):
            value = domain.monthly_revenue(month, region_idx)
            if current and (month, region_idx) == (3, 0):
                value += domain.P04_DELTA  # P04: historical table cell edited
            table.cell(region_idx + 1, month + 1).text = f"{value:,.0f}"


def _slide_trend_chart(prs: PresentationType, *, current: bool) -> None:
    slide = _add_slide(prs, "Revenue Trend")
    months = domain.CURRENT_MONTHS if current else domain.BASELINE_MONTHS
    chart_data = ChartData()
    chart_data.categories = domain.MONTH_LABELS[:months]
    chart_data.add_series("Revenue", _trend_values(current=current))
    slide.shapes.add_chart(
        XL_CHART_TYPE.LINE, Inches(0.5), Inches(1.6), Inches(9.0), Inches(4.5), chart_data
    )


def _slide_weekly_chart(prs: PresentationType, *, current: bool) -> None:
    slide = _add_slide(prs, "Weekly Ops")
    total_weeks = domain.CURRENT_WEEKS if current else domain.BASELINE_WEEKS
    window = domain.WEEK_LABELS[total_weeks - domain.ROLLING_WINDOW : total_weeks]
    chart_data = ChartData()
    chart_data.categories = window
    chart_data.add_series(
        "Weekly revenue", [domain.weekly_revenue(int(w[1:]) - 1) for w in window]
    )
    slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.5),
        Inches(1.6),
        Inches(9.0),
        Inches(4.5),
        chart_data,
    )


def _slide_notes(prs: PresentationType) -> None:
    slide = _add_slide(prs, "Notes & Definitions")
    _add_bullets(
        slide,
        [
            "Revenue is gross of returns.",
            "Cost includes allocated overhead.",
            "Margin = Revenue - Cost.",
        ],
    )


def _slide_deep_dive(prs: PresentationType) -> None:
    slide = _add_slide(prs, "Deep Dive Archive")
    _add_bullets(slide, ["Historical deep-dive material (retired)."])


def _slide_new_initiatives(prs: PresentationType) -> None:
    slide = _add_slide(prs, "New Initiatives")
    _add_bullets(slide, ["Pilot program launching next cycle."])


def _build_deck(*, current: bool) -> PresentationType:
    prs = Presentation()
    if current:
        _slide_exec_summary(prs, current=True)
        _slide_notes(prs)  # reordered vs baseline
        _slide_region_table(prs, current=True)
        _slide_new_initiatives(prs)  # P01
        _slide_trend_chart(prs, current=True)
        _slide_weekly_chart(prs, current=True)
    else:
        _slide_exec_summary(prs, current=False)
        _slide_region_table(prs, current=False)
        _slide_trend_chart(prs, current=False)
        _slide_weekly_chart(prs, current=False)
        _slide_notes(prs)
        _slide_deep_dive(prs)  # P02 (removed in current)
    prs.core_properties.created = FIXED_DOC_TIME
    prs.core_properties.modified = FIXED_DOC_TIME
    return prs


def build_decks(
    dest: Path,
) -> tuple[list[SeededDefect], list[ExpectedChange], list[CrosscheckEntry]]:
    """Write baseline.pptx and current.pptx; return seeded ground truth."""
    _build_deck(current=False).save(str(dest / "baseline.pptx"))
    _build_deck(current=True).save(str(dest / "current.pptx"))

    defects = [
        SeededDefect(
            defect_id="P01",
            artifact=Artifact.PPT,
            classes=[DefectClass.SLIDE_ADDED],
            slide_title="New Initiatives",
            note="current-only slide",
        ),
        SeededDefect(
            defect_id="P02",
            artifact=Artifact.PPT,
            classes=[DefectClass.SLIDE_REMOVED],
            slide_title="Deep Dive Archive",
            note="baseline-only slide",
        ),
        SeededDefect(
            defect_id="P03",
            artifact=Artifact.PPT,
            classes=[DefectClass.SLIDE_TEXT_CHANGED],
            slide_title="Executive Summary",
            baseline="Outlook remains stable across all regions.",
            current="Outlook remains volatile across all regions.",
            note="bullet text changed",
        ),
        SeededDefect(
            defect_id="P04",
            artifact=Artifact.PPT,
            classes=[DefectClass.TABLE_VALUE_CHANGED],
            slide_title="Revenue by Region",
            element="North / Apr-26",
            baseline=f"{domain.monthly_revenue(3, 0):,.0f}",
            current=f"{domain.monthly_revenue(3, 0) + domain.P04_DELTA:,.0f}",
            note="historical table cell edited",
        ),
        SeededDefect(
            defect_id="P05",
            artifact=Artifact.PPT,
            classes=[DefectClass.CHART_VALUE_CHANGED],
            slide_title="Revenue Trend",
            element="Mar-26",
            baseline=str(domain.month_total_revenue(2, current=False)),
            current=str(domain.month_total_revenue(2, current=False) + domain.P05_DELTA),
            note="historical chart point edited",
        ),
    ]

    expected = [
        ExpectedChange(
            change_id="PX01",
            artifact=Artifact.PPT,
            kind="slide_reordered",
            slide_title="Notes & Definitions",
            detail="moved from position 5 to position 2",
        ),
        ExpectedChange(
            change_id="PX02",
            artifact=Artifact.PPT,
            kind="chart_rolling_shift",
            slide_title="Weekly Ops",
            detail="window W17-W20 -> W18-W21 (size 4, shift 1)",
        ),
        ExpectedChange(
            change_id="PX03",
            artifact=Artifact.PPT,
            kind="chart_new_period",
            slide_title="Revenue Trend",
            detail="Jun-26 point appended",
        ),
        ExpectedChange(
            change_id="PX04",
            artifact=Artifact.PPT,
            kind="table_new_period",
            slide_title="Revenue by Region",
            detail="Jun-26 column appended",
        ),
    ]

    crosscheck = [
        CrosscheckEntry(
            slide_title="Executive Summary",
            figure_label="Total revenue",
            figure_text=domain.fmt_millions(domain.total_revenue(current=True)),
            source_cell="Dashboard!B2",
            matches=True,
        ),
        CrosscheckEntry(
            slide_title="Executive Summary",
            figure_label="Margin",
            figure_text=domain.fmt_pct(
                domain.margin_ratio(current=True) + domain.X03_MARGIN_OFFSET
            ),
            source_cell="Dashboard!B4",
            matches=False,
            defect_id="X03",
        ),
        CrosscheckEntry(
            slide_title="Revenue by Region",
            figure_label="North / Jan-26",
            figure_text=f"{domain.monthly_revenue(0, 0):,.0f}",
            source_cell="Long_Monthly!C2",
            matches=True,
        ),
    ]

    defects.append(
        SeededDefect(
            defect_id="X03",
            artifact=Artifact.CROSSCHECK,
            classes=[DefectClass.CROSSCHECK_MISMATCH],
            slide_title="Executive Summary",
            element="Margin",
            baseline=domain.fmt_pct(domain.margin_ratio(current=True)),
            current=domain.fmt_pct(domain.margin_ratio(current=True) + domain.X03_MARGIN_OFFSET),
            note="deck margin drifts from Dashboard!B4",
        )
    )
    return defects, expected, crosscheck
