"""Repeated PowerPoint claim identity and disagreement contracts."""

from __future__ import annotations

from qc_tool.coverage import CoverageState
from qc_tool.findings import FindingClass
from qc_tool.ppt.claim_periods import periods_in_text
from qc_tool.ppt.model import (
    ChartContent,
    DeckSnapshot,
    PptChartLabel,
    PptChartPlot,
    PptChartSeries,
    ShapeContent,
    SlideContent,
    TableContent,
)
from qc_tool.ppt.repetition import check_repeated_claims


def _text_slide(index: int, title: str, *lines: str) -> SlideContent:
    return SlideContent(
        index=index,
        title=title,
        texts=list(lines),
        shape_count=max(1, len(lines)),
    )


def _deck(*slides: SlideContent, charts_available: bool = True) -> DeckSnapshot:
    return DeckSnapshot(
        source_name="claims.pptx",
        slides=list(slides),
        charts_available=charts_available,
    )


def _mismatches(deck: DeckSnapshot):
    return tuple(
        finding
        for finding in check_repeated_claims(deck).findings
        if finding.finding_class is FindingClass.PPT_REPEATED_CLAIM_MISMATCH
    )


def _chart_slide(index: int, title: str, label_text: str) -> SlideContent:
    series = PptChartSeries(
        index=0,
        order=0,
        name="Actual",
        categories=["Jan-26"],
        values=[100.0],
        plot_index=0,
        source_index=0,
        source_id=f"series-{index}",
    )
    plot = PptChartPlot(
        index=0,
        chart_type="bar",
        series=[series],
        visible_labels=[
            PptChartLabel(
                text=label_text,
                series_source_index=0,
                category_index=0,
                category="Jan-26",
            )
        ],
    )
    chart = ChartContent(
        chart_type="bar",
        categories=["Jan-26"],
        series=[("Actual", [100.0])],
        source_index=0,
        shape_id=index + 10,
        title="Revenue",
        plots=[plot],
    )
    return SlideContent(
        index=index,
        title=title,
        texts=[],
        charts=[chart],
        shape_count=1,
    )


def test_period_labels_are_returned_in_source_order_across_grammars() -> None:
    labels = [
        label
        for label, _period in periods_in_text(
            "Q1 2023 ended before Jan 2024 and 2024-03-31; W14-2024 followed"
        )
    ]

    assert labels == ["Q1 2023", "Jan 2024", "2024-03-31", "W14-2024"]


def test_repeated_text_claims_agree_without_a_finding() -> None:
    result = check_repeated_claims(
        _deck(
            _text_slide(0, "Executive", "Revenue $100M for Jan-26"),
            _text_slide(1, "Appendix", "Revenue $100M for Jan-26"),
        )
    )

    assert result.findings == ()
    assert result.coverage.state is CoverageState.CHECKED
    assert "eligible_groups=1" in result.coverage.detail
    assert "consistent_groups=1" in result.coverage.detail


def test_repeated_text_claim_disagreement_emits_one_deterministic_finding() -> None:
    deck = _deck(
        _text_slide(0, "Executive", "Revenue $100M for Jan-26"),
        _text_slide(1, "Appendix", "Revenue $200M for Jan-26"),
    )

    first = _mismatches(deck)
    second = _mismatches(deck)

    assert len(first) == 1
    assert first[0].event_key == second[0].event_key
    assert first[0].event_key.startswith("ppt-repetition:")
    assert first[0].event_key == first[0].root_cause_key
    assert first[0].location == "slide 1; slide 2"
    assert first[0].baseline_value == "slide 1: $100M"
    assert first[0].current_value == "slide 2: $200M"


def test_three_slide_disagreement_fans_out_to_one_decision() -> None:
    mismatches = _mismatches(
        _deck(
            _text_slide(0, "Executive", "Revenue $100M for Jan-26"),
            _text_slide(1, "Operations", "Revenue $110M for Jan-26"),
            _text_slide(2, "Appendix", "Revenue $120M for Jan-26"),
        )
    )

    assert len(mismatches) == 1
    assert mismatches[0].location == "slide 1; slide 2; slide 3"
    assert "slide 2: $110M" in (mismatches[0].current_value or "")
    assert "slide 3: $120M" in (mismatches[0].current_value or "")


def test_period_unit_and_precision_axes_prevent_false_comparisons() -> None:
    deck = _deck(
        _text_slide(0, "January", "Revenue $100M for Jan-26"),
        _text_slide(1, "February", "Revenue $200M for Feb-26"),
        _text_slide(2, "Units", "Revenue 300% for Jan-26"),
        _text_slide(3, "Precision", "Revenue $100.0M for Jan-26"),
    )

    result = check_repeated_claims(deck)

    assert result.findings == ()
    assert "mismatch_groups=0" in result.coverage.detail


def test_ambiguous_period_is_unavailable_not_guessed() -> None:
    result = check_repeated_claims(
        _deck(
            _text_slide(0, "A", "Revenue $100M for Jan-26 and Feb-26"),
            _text_slide(1, "B", "Revenue $200M for Jan-26 and Feb-26"),
        )
    )

    assert result.findings == ()
    assert result.coverage.state is CoverageState.DEGRADED
    assert "identity_unavailable=2" in result.coverage.detail
    assert "ambiguous_period=2" in result.coverage.detail


def test_repeated_table_claims_are_compared_by_row_and_header_anchor() -> None:
    first = _text_slide(0, "Executive")
    first.tables = [
        TableContent(rows=[["Metric", "Jan-26"], ["Revenue", "$100M"]])
    ]
    second = _text_slide(1, "Appendix")
    second.tables = [
        TableContent(rows=[["Metric", "Jan-26"], ["Revenue", "$200M"]])
    ]

    mismatches = _mismatches(_deck(first, second))

    assert len(mismatches) == 1
    assert mismatches[0].element == "table:revenue/"


def test_repeated_native_chart_labels_are_compared() -> None:
    mismatches = _mismatches(
        _deck(
            _chart_slide(0, "Executive", "$100M"),
            _chart_slide(1, "Appendix", "$200M"),
        )
    )

    assert len(mismatches) == 1
    assert mismatches[0].element == "chart:revenue/actual/"


def test_opaque_shapes_and_unavailable_chart_labels_degrade_coverage() -> None:
    opaque = _text_slide(0, "Executive", "Revenue $100M for Jan-26")
    opaque.shapes = [
        ShapeContent(
            source_id="picture-1",
            source_index=0,
            shape_id=1,
            shape_type="PICTURE",
            name="chart.png",
            left=0,
            top=0,
            width=100,
            height=100,
            z_order=0,
        )
    ]
    result = check_repeated_claims(
        _deck(opaque, charts_available=False)
    )

    assert result.coverage.state is CoverageState.DEGRADED
    assert "opaque_surfaces=1" in result.coverage.detail
    assert "visible native chart labels unavailable" in result.coverage.detail


def test_repetition_coverage_population_fields_always_reconcile() -> None:
    result = check_repeated_claims(
        _deck(
            _text_slide(0, "Executive", "Revenue $100M for Jan-26"),
            _text_slide(1, "Appendix", "Revenue $200M for Jan-26"),
            _text_slide(2, "Narrative", "No figures here"),
        )
    )

    for field in (
        "readable=2",
        "identity_available=2",
        "identity_unavailable=0",
        "eligible_groups=1",
        "consistent_groups=0",
        "mismatch_groups=1",
        "opaque_surfaces=0",
    ):
        assert field in result.coverage.detail
