"""Within-slide PowerPoint element matching and complete chart/table diff."""

from qc_tool.config.profile import PptProfile
from qc_tool.coverage import CoverageState
from qc_tool.crosscheck.trace import extract_deck_figures
from qc_tool.findings import FindingClass
from qc_tool.io.model import ChartAxis, ChartDataLabels, ChartLegend
from qc_tool.ppt.diff import diff_decks
from qc_tool.ppt.element_match import match_slide_elements
from qc_tool.ppt.match import SlideMatching
from qc_tool.ppt.model import (
    ChartContent,
    DeckSnapshot,
    PptChartLabel,
    PptChartPlot,
    PptChartSeries,
    SlideContent,
    TableContent,
)


def _table(
    metric: str,
    *,
    source_index: int,
    left: int,
    z_order: int,
) -> TableContent:
    return TableContent(
        rows=[["Metric", "Jan"], [metric, "10"]],
        source_id=f"table-{source_index}",
        source_index=source_index,
        left=left,
        top=100,
        width=200,
        height=100,
        z_order=z_order,
    )


def _series(
    name: str,
    values: list[float | None],
    categories: list[str] | None = None,
    *,
    index: int = 0,
    order: int | None = None,
    plot_index: int = 0,
) -> PptChartSeries:
    return PptChartSeries(
        index=index,
        order=index if order is None else order,
        name=name,
        categories=categories or ["Jan-26", "Feb-26", "Mar-26"],
        values=values,
        plot_index=plot_index,
        source_index=index,
        source_id=f"plot-{plot_index}/series-{index}",
    )


def _plot(
    chart_type: str,
    series: list[PptChartSeries],
    *,
    index: int,
    axis_group: str = "primary",
    labels: ChartDataLabels | None = None,
) -> PptChartPlot:
    for series_item in series:
        series_item.plot_index = index
    return PptChartPlot(
        index=index,
        chart_type=chart_type,
        series=series,
        axis_ids=(str(index * 2 + 10), str(index * 2 + 11)),
        axis_group=axis_group,
        data_labels=labels,
        source_id=f"plot-{index}",
    )


def _chart(
    plots: list[PptChartPlot],
    *,
    title: str | None = "Panel",
    source_index: int = 0,
    left: int = 0,
    z_order: int = 1,
    axes: list[ChartAxis] | None = None,
    legend: ChartLegend | None = None,
) -> ChartContent:
    first = plots[0]
    return ChartContent(
        chart_type=first.chart_type,
        categories=first.series[0].categories if first.series else [],
        series=[(series.name, series.values) for series in first.series],
        source_id=f"chart-{source_index}",
        source_index=source_index,
        title=title,
        plots=plots,
        axes=axes or [],
        legend=legend,
        left=left,
        top=100,
        width=400,
        height=300,
        z_order=z_order,
    )


def _slide(
    *,
    tables: list[TableContent] | None = None,
    charts: list[ChartContent] | None = None,
) -> SlideContent:
    return SlideContent(
        index=0,
        title="Dashboard",
        texts=[],
        tables=tables or [],
        charts=charts or [],
    )


def _findings(
    baseline: SlideContent,
    current: SlideContent,
    profile: PptProfile | None = None,
):
    return diff_decks(SlideMatching(pairs=[(baseline, current)]), profile)


def test_element_matcher_survives_insertions_reorders_and_duplicate_panels() -> None:
    revenue_table = _table("Revenue", source_index=0, left=0, z_order=1)
    margin_table = _table("Margin", source_index=1, left=300, z_order=2)
    revenue_chart = _chart(
        [_plot("barChart", [_series("Revenue", [10, 20, 30])], index=0)],
        source_index=0,
        left=0,
        z_order=3,
    )
    margin_chart = _chart(
        [_plot("lineChart", [_series("Margin", [1, 2, 3])], index=0)],
        source_index=1,
        left=500,
        z_order=4,
    )
    baseline = _slide(
        tables=[revenue_table, margin_table],
        charts=[revenue_chart, margin_chart],
    )
    current = _slide(
        tables=[
            _table("Cost", source_index=0, left=600, z_order=1),
            _table("Margin", source_index=1, left=300, z_order=2),
            _table("Revenue", source_index=2, left=0, z_order=3),
        ],
        charts=[
            _chart(
                [_plot("lineChart", [_series("Margin", [1, 2, 3])], index=0)],
                source_index=0,
                left=500,
                z_order=4,
            ),
            _chart(
                [_plot("barChart", [_series("Revenue", [10, 20, 30])], index=0)],
                source_index=1,
                left=0,
                z_order=5,
            ),
            _chart(
                [_plot("lineChart", [_series("Cost", [7, 8, 9])], index=0)],
                source_index=2,
                left=900,
                z_order=6,
            ),
        ],
    )

    matching = match_slide_elements(baseline, current)

    assert {(base.rows[1][0], curr.rows[1][0]) for base, curr in matching.table_pairs} == {
        ("Revenue", "Revenue"),
        ("Margin", "Margin"),
    }
    assert [table.rows[1][0] for table in matching.added_tables] == ["Cost"]
    assert len(matching.reordered_tables) == 1
    assert {
        (base.all_series[0].name, curr.all_series[0].name)
        for base, curr in matching.chart_pairs
    } == {("Revenue", "Revenue"), ("Margin", "Margin")}
    assert [chart.all_series[0].name for chart in matching.added_charts] == ["Cost"]
    assert len(matching.reordered_charts) == 1


def test_diff_emits_explicit_table_chart_plot_and_series_inventory_changes() -> None:
    baseline = _slide(
        tables=[_table("Revenue", source_index=0, left=0, z_order=1)],
        charts=[
            _chart(
                [
                    _plot("barChart", [_series("Revenue", [10, 20, 30])], index=0),
                    _plot("lineChart", [_series("Margin", [1, 2, 3])], index=1),
                ]
            )
        ],
    )
    current = _slide(
        tables=[_table("Cost", source_index=0, left=600, z_order=1)],
        charts=[
            _chart(
                [
                    _plot(
                        "barChart",
                        [
                            _series("Revenue", [10, 20, 30]),
                            _series("Cost", [7, 8, 9], index=1),
                        ],
                        index=0,
                    )
                ]
            ),
            _chart(
                [_plot("lineChart", [_series("Headcount", [5, 6, 7])], index=0)],
                title="Second panel",
                source_index=1,
                left=500,
            ),
        ],
    )

    findings = _findings(baseline, current)
    classes = {finding.finding_class for finding in findings}

    assert FindingClass.PPT_TABLE_STRUCTURE_CHANGED in classes
    assert FindingClass.PPT_CHART_STRUCTURE_CHANGED in classes
    assert FindingClass.PPT_CHART_PLOT_CHANGED in classes
    assert FindingClass.PPT_CHART_SERIES_CHANGED in classes


def test_combo_plot_type_and_series_rename_are_detected_without_mispairing() -> None:
    baseline = _slide(
        charts=[
            _chart(
                [
                    _plot("barChart", [_series("Revenue", [10, 20, 30])], index=0),
                    _plot("lineChart", [_series("Margin", [1, 2, 3])], index=1),
                ]
            )
        ]
    )
    current = _slide(
        charts=[
            _chart(
                [
                    _plot("barChart", [_series("Revenue", [10, 20, 30])], index=0),
                    _plot("areaChart", [_series("Gross margin", [1, 2, 3])], index=1),
                ]
            )
        ]
    )

    findings = _findings(baseline, current)

    plot_findings = [
        finding
        for finding in findings
        if finding.finding_class is FindingClass.PPT_CHART_PLOT_CHANGED
    ]
    series_findings = [
        finding
        for finding in findings
        if finding.finding_class is FindingClass.PPT_CHART_SERIES_CHANGED
    ]
    assert len(plot_findings) == 1
    assert "lineChart" in plot_findings[0].message
    assert "areaChart" in plot_findings[0].message
    assert len(series_findings) == 1
    assert "Margin" in series_findings[0].message
    assert "Gross margin" in series_findings[0].message
    assert not any("Revenue" in finding.message for finding in findings)


def test_series_windows_are_classified_independently_and_values_still_checked() -> None:
    baseline = _slide(
        charts=[
            _chart(
                [
                    _plot(
                        "lineChart",
                        [
                            _series(
                                "Rolling",
                                [10, 20, 30],
                                ["Jan-26", "Feb-26", "Mar-26"],
                            )
                        ],
                        index=0,
                    ),
                    _plot(
                        "lineChart",
                        [
                            _series(
                                "Full",
                                [100, 200, 300],
                                ["Jan-26", "Feb-26", "Mar-26"],
                            )
                        ],
                        index=1,
                    ),
                ]
            )
        ]
    )
    current = _slide(
        charts=[
            _chart(
                [
                    _plot(
                        "lineChart",
                        [
                            _series(
                                "Rolling",
                                [20, 30, 40],
                                ["Feb-26", "Mar-26", "Apr-26"],
                            )
                        ],
                        index=0,
                    ),
                    _plot(
                        "lineChart",
                        [
                            _series(
                                "Full",
                                [100, 999, 300, 400],
                                ["Jan-26", "Feb-26", "Mar-26", "Apr-26"],
                            )
                        ],
                        index=1,
                    ),
                ]
            )
        ]
    )

    findings = _findings(baseline, current)

    rolling = [finding for finding in findings if "Rolling" in finding.message]
    full = [finding for finding in findings if "Full" in finding.message]
    assert rolling and all(finding.expected_growth for finding in rolling)
    assert any(finding.expected_growth for finding in full)
    assert any(not finding.expected_growth and "Feb-26" in finding.message for finding in full)


def test_chart_axis_legend_labels_geometry_and_order_changes_are_detected() -> None:
    labels_off = ChartDataLabels(show_value=False)
    labels_on = ChartDataLabels(show_value=True)
    baseline = _slide(
        charts=[
            _chart(
                [_plot("barChart", [_series("Revenue", [10, 20, 30])], index=0, labels=labels_off)],
                axes=[ChartAxis("10", "valAx", 0, maximum=40, major_unit=10)],
                legend=ChartLegend(position="r"),
                left=0,
                z_order=1,
            )
        ]
    )
    current = _slide(
        charts=[
            _chart(
                [_plot("barChart", [_series("Revenue", [10, 20, 30])], index=0, labels=labels_on)],
                axes=[ChartAxis("10", "valAx", 0, maximum=50, major_unit=5)],
                legend=ChartLegend(position="b"),
                left=100_000,
                z_order=2,
            )
        ]
    )

    classes = {finding.finding_class for finding in _findings(baseline, current)}

    assert FindingClass.PPT_CHART_AXIS_CHANGED in classes
    assert FindingClass.PPT_CHART_LEGEND_CHANGED in classes
    assert FindingClass.PPT_CHART_LABELS_CHANGED in classes
    assert FindingClass.PPT_SHAPE_GEOMETRY_CHANGED in classes


def test_visible_chart_labels_have_stable_anchors_and_dedupe_slide_text() -> None:
    series = _series("Revenue", [10, 20], ["Jan-26", "Feb-26"])
    plot = _plot("lineChart", [series], index=0, labels=ChartDataLabels(show_value=True))
    plot.visible_labels = [
        PptChartLabel("10", 0, 0, "Jan-26"),
        PptChartLabel("20", 0, 1, "Feb-26"),
    ]
    plot.visible_label_values = [label.text for label in plot.visible_labels]
    chart = _chart([plot], title="Trend")
    slide = _slide(charts=[chart])
    slide.texts = ["Revenue 10"]
    deck = DeckSnapshot("deck.pptx", slides=[slide])

    occurrences = extract_deck_figures(deck)

    assert sum(occurrence.figure.raw == "10" for occurrence in occurrences) == 1
    chart_occurrence = next(
        occurrence for occurrence in occurrences if occurrence.figure.raw == "20"
    )
    assert chart_occurrence.line_skeleton == "chart:Trend/Revenue/Feb-26"


def test_plot_and_series_reorders_are_explicit_findings() -> None:
    revenue = _series("Revenue", [10, 20, 30], index=0, order=0)
    cost = _series("Cost", [7, 8, 9], index=1, order=1)
    margin = _series("Margin", [1, 2, 3], index=0, order=0)
    baseline = _slide(
        charts=[
            _chart(
                [
                    _plot("barChart", [revenue, cost], index=0),
                    _plot("lineChart", [margin], index=1),
                ]
            )
        ]
    )
    current_cost = _series("Cost", [7, 8, 9], index=0, order=0)
    current_revenue = _series("Revenue", [10, 20, 30], index=1, order=1)
    current_margin = _series("Margin", [1, 2, 3], index=0, order=0)
    current = _slide(
        charts=[
            _chart(
                [
                    _plot("lineChart", [current_margin], index=0),
                    _plot("barChart", [current_cost, current_revenue], index=1),
                ]
            )
        ]
    )

    findings = _findings(baseline, current)

    assert any(
        finding.finding_class is FindingClass.PPT_CHART_PLOT_CHANGED
        and "reordered" in finding.message
        for finding in findings
    )
    assert any(
        finding.finding_class is FindingClass.PPT_CHART_SERIES_CHANGED
        and "reordered" in finding.message
        for finding in findings
    )


def test_full_pipeline_marks_semantic_ppt_comparison_checked(qc_result) -> None:
    coverage = next(
        item for item in qc_result.coverage if item.check_id == "ppt-comparison"
    )

    assert coverage.state is CoverageState.CHECKED
    assert coverage.detail == "Complete semantic slide-element comparison"


def test_unrelated_leftover_tables_remain_explicitly_unmatched() -> None:
    baseline = _slide(
        tables=[_table("Revenue", source_index=0, left=0, z_order=1)]
    )
    current = _slide(
        tables=[
            TableContent(
                rows=[["Other", "Feb"], ["Cost", "7"]],
                source_id="replacement",
                source_index=5,
                left=900,
                top=500,
                width=100,
                height=50,
                z_order=4,
            )
        ]
    )

    matching = match_slide_elements(baseline, current)

    assert matching.table_pairs == []
    assert [table.rows[1][0] for table in matching.removed_tables] == ["Revenue"]
    assert [table.rows[1][0] for table in matching.added_tables] == ["Cost"]


def test_untitled_combo_chart_matches_across_plot_reorder() -> None:
    baseline_chart = _chart(
        [
            _plot("barChart", [_series("Revenue", [10, 20, 30])], index=0),
            _plot("lineChart", [_series("Margin", [1, 2, 3])], index=1),
        ],
        title=None,
        source_index=0,
        left=0,
    )
    current_chart = _chart(
        [
            _plot("lineChart", [_series("Margin", [1, 2, 3])], index=0),
            _plot("barChart", [_series("Revenue", [10, 20, 30])], index=1),
        ],
        title=None,
        source_index=3,
        left=900,
    )

    matching = match_slide_elements(
        _slide(charts=[baseline_chart]),
        _slide(charts=[current_chart]),
    )

    assert matching.chart_pairs == [(baseline_chart, current_chart)]
    assert matching.removed_charts == []
    assert matching.added_charts == []


def test_repeated_chart_label_value_consumes_only_one_text_duplicate() -> None:
    series = _series("Revenue", [10, 10], ["Jan-26", "Feb-26"])
    plot = _plot("lineChart", [series], index=0)
    plot.visible_labels = [
        PptChartLabel("10", 0, 0, "Jan-26"),
        PptChartLabel("10", 0, 1, "Feb-26"),
    ]
    chart = _chart([plot], title="Trend")
    slide = _slide(charts=[chart])
    slide.texts = ["Revenue 10"]

    occurrences = extract_deck_figures(DeckSnapshot("deck.pptx", slides=[slide]))

    tens = [occurrence for occurrence in occurrences if occurrence.figure.raw == "10"]
    assert len(tens) == 2
    assert sum(occurrence.line_skeleton.startswith("chart:") for occurrence in tens) == 1


def test_chart_window_override_lookup_is_case_insensitive() -> None:
    baseline = _slide(
        charts=[
            _chart(
                [
                    _plot(
                        "lineChart",
                        [
                            _series(
                                "Rolling",
                                [10, 20, 30],
                                ["Jan-26", "Feb-26", "Mar-26"],
                            )
                        ],
                        index=0,
                    )
                ],
                title="Panel",
            )
        ]
    )
    current = _slide(
        charts=[
            _chart(
                [
                    _plot(
                        "lineChart",
                        [
                            _series(
                                "Rolling",
                                [20, 30, 40],
                                ["Feb-26", "Mar-26", "Apr-26"],
                            )
                        ],
                        index=0,
                    )
                ],
                title="Panel",
            )
        ]
    )
    profile = PptProfile(
        chart_windows={"dashboard/panel/rolling": "full"}
    )

    findings = _findings(baseline, current, profile)

    removed = [
        finding
        for finding in findings
        if finding.element == "Jan-26" and "removed" in finding.message
    ]
    assert len(removed) == 1
    assert not removed[0].expected_growth
