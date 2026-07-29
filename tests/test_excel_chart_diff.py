"""Semantic Excel chart, plot, series, axis, and geometry QC."""

from qc_tool.config.profile import DeliverableProfile
from qc_tool.coverage import CoverageState
from qc_tool.excel.charts import (
    annotate_chart_impacts,
    chart_reference_coverage,
    diff_charts,
)
from qc_tool.findings import Finding, FindingClass
from qc_tool.io.model import (
    ChartAnchor,
    ChartAxis,
    ChartDataLabels,
    ChartDescriptor,
    ChartLegend,
    ChartPlot,
    ChartSeries,
    SheetSnapshot,
    WorkbookSnapshot,
)


def _series(
    index: int,
    name: str,
    values: str,
    categories: str = "Data!$A$2:$A$4",
) -> ChartSeries:
    return ChartSeries(
        index=index,
        order=index,
        name_text=name,
        values_ref=values,
        categories_ref=categories,
        source_index=index,
    )


def _chart(
    plots: list[ChartPlot],
    *,
    axes: list[ChartAxis] | None = None,
    legend: ChartLegend | None = None,
    anchor: ChartAnchor | None = None,
) -> ChartDescriptor:
    chart = ChartDescriptor(
        sheet="Data",
        title="Combo",
        chart_type=plots[0].chart_type,
        source_index=0,
        source_id="Data!chart[0]",
        plots=plots,
        axes=axes or [],
        legend=legend,
        anchor=anchor,
    )
    chart.series = [series for plot in plots for series in plot.series]
    return chart


def _workbook(chart: ChartDescriptor) -> WorkbookSnapshot:
    return WorkbookSnapshot(
        "chart.xlsx",
        "xlsx",
        True,
        True,
        sheets=[SheetSnapshot("Data", "visible", 20, 20)],
        charts=[chart],
    )


def test_plot_type_change_inside_combo_is_detected_without_mispairing() -> None:
    baseline = _chart(
        [
            ChartPlot(0, "barChart", [_series(0, "Revenue", "Data!$B$2:$B$4")]),
            ChartPlot(1, "lineChart", [_series(0, "Margin", "Data!$C$2:$C$4")]),
        ]
    )
    current = _chart(
        [
            ChartPlot(0, "barChart", [_series(0, "Revenue", "Data!$B$2:$B$4")]),
            ChartPlot(1, "areaChart", [_series(0, "Margin", "Data!$C$2:$C$4")]),
        ]
    )

    findings = diff_charts(_workbook(baseline), _workbook(current))

    assert [finding.finding_class for finding in findings] == [
        FindingClass.CHART_PLOT_CHANGED
    ]
    assert "lineChart" in findings[0].message
    assert "areaChart" in findings[0].message


def test_added_primary_series_does_not_change_secondary_series() -> None:
    baseline = _chart(
        [
            ChartPlot(0, "barChart", [_series(0, "Revenue", "Data!$B$2:$B$4")]),
            ChartPlot(1, "lineChart", [_series(0, "Margin", "Data!$C$2:$C$4")]),
        ]
    )
    current = _chart(
        [
            ChartPlot(
                0,
                "barChart",
                [
                    _series(0, "Revenue", "Data!$B$2:$B$4"),
                    _series(1, "Cost", "Data!$D$2:$D$4"),
                ],
            ),
            ChartPlot(1, "lineChart", [_series(0, "Margin", "Data!$C$2:$C$4")]),
        ]
    )

    findings = diff_charts(_workbook(baseline), _workbook(current))

    assert len(findings) == 1
    assert findings[0].finding_class is FindingClass.CHART_SERIES_CHANGED
    assert "Cost" in findings[0].message
    assert "added" in findings[0].message


def test_series_name_and_source_changes_are_detected() -> None:
    baseline = _chart(
        [ChartPlot(0, "barChart", [_series(0, "Revenue", "Data!$B$2:$B$4")])]
    )
    current = _chart(
        [ChartPlot(0, "barChart", [_series(0, "Sales", "Data!$C$2:$C$4")])]
    )

    findings = diff_charts(_workbook(baseline), _workbook(current))

    assert len(findings) == 2
    assert all(
        finding.finding_class is FindingClass.CHART_SERIES_CHANGED
        for finding in findings
    )
    assert any("name" in finding.message for finding in findings)
    assert any("values" in finding.message for finding in findings)


def test_axis_group_scale_legend_labels_and_geometry_changes_are_detected() -> None:
    base_plot = ChartPlot(
        0,
        "barChart",
        [_series(0, "Revenue", "Data!$B$2:$B$4")],
        axis_ids=("10", "100"),
        axis_group="primary",
        data_labels=ChartDataLabels(show_value=False),
    )
    curr_plot = ChartPlot(
        0,
        "barChart",
        [_series(0, "Revenue", "Data!$B$2:$B$4")],
        axis_ids=("20", "200"),
        axis_group="secondary-1",
        data_labels=ChartDataLabels(show_value=True),
    )
    baseline = _chart(
        [base_plot],
        axes=[ChartAxis("100", "valAx", 0, maximum=40, major_unit=10)],
        legend=ChartLegend(position="r"),
        anchor=ChartAnchor("oneCellAnchor", from_col=4, from_row=1, width=5_000_000),
    )
    current = _chart(
        [curr_plot],
        axes=[ChartAxis("200", "valAx", 0, maximum=50, major_unit=5)],
        legend=ChartLegend(position="b"),
        anchor=ChartAnchor("oneCellAnchor", from_col=5, from_row=1, width=6_000_000),
    )

    findings = diff_charts(_workbook(baseline), _workbook(current))
    classes = {finding.finding_class for finding in findings}

    assert FindingClass.CHART_AXIS_CHANGED in classes
    assert FindingClass.CHART_LABELS_CHANGED in classes
    assert FindingClass.CHART_LEGEND_CHANGED in classes
    assert FindingClass.CHART_GEOMETRY_CHANGED in classes


def test_range_extension_and_rolling_shift_remain_expected() -> None:
    baseline = _chart(
        [
            ChartPlot(
                0,
                "lineChart",
                [
                    _series(
                        0,
                        "Revenue",
                        "Data!$B$2:$B$4",
                        "Data!$A$2:$A$4",
                    )
                ],
            )
        ]
    )
    extended = _chart(
        [
            ChartPlot(
                0,
                "lineChart",
                [
                    _series(
                        0,
                        "Revenue",
                        "Data!$B$2:$B$5",
                        "Data!$A$2:$A$5",
                    )
                ],
            )
        ]
    )
    rolling = _chart(
        [
            ChartPlot(
                0,
                "lineChart",
                [
                    _series(
                        0,
                        "Revenue",
                        "Data!$B$3:$B$5",
                        "Data!$A$3:$A$5",
                    )
                ],
            )
        ]
    )

    extension_findings = diff_charts(_workbook(baseline), _workbook(extended))
    rolling_findings = diff_charts(_workbook(baseline), _workbook(rolling))

    assert extension_findings and all(
        finding.expected_growth for finding in extension_findings
    )
    assert rolling_findings and all(finding.expected_growth for finding in rolling_findings)

    full_profile = DeliverableProfile.model_validate(
        {
            "name": "full-history",
            "excel": {
                "sheets": {"Data": {"chart_windows": {"Combo": "full"}}}
            },
        }
    )
    full_findings = diff_charts(
        _workbook(baseline),
        _workbook(rolling),
        full_profile,
    )
    assert full_findings and not any(
        finding.expected_growth for finding in full_findings
    )


def test_source_cell_findings_include_chart_series_impact() -> None:
    chart = _chart(
        [ChartPlot(0, "barChart", [_series(0, "Revenue", "Data!$B$2:$B$4")])]
    )
    workbook = _workbook(chart)
    finding = Finding(
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        sheet="Data",
        location="B3",
        message="value changed",
    )

    annotate_chart_impacts([finding], workbook)

    assert finding.impacts == ["Excel chart 'Combo' series 'Revenue'"]


def test_unsupported_chart_sources_degrade_and_invalid_sources_are_findings() -> None:
    unsupported = _chart(
        [
            ChartPlot(
                0,
                "barChart",
                [_series(0, "Revenue", "[Book.xlsx]Data!$B$2:$B$4")],
            )
        ]
    )
    state, detail = chart_reference_coverage(_workbook(unsupported))

    assert state is CoverageState.DEGRADED
    assert "unsupported" in detail.lower()

    baseline = _chart(
        [ChartPlot(0, "barChart", [_series(0, "Revenue", "Data!$B$2:$B$4")])]
    )
    invalid = _chart(
        [ChartPlot(0, "barChart", [_series(0, "Revenue", "Data!$Z$2:$Z$4")])]
    )
    findings = diff_charts(_workbook(baseline), _workbook(invalid))

    assert any(
        finding.finding_class is FindingClass.CHART_REFERENCE_INVALID
        for finding in findings
    )


def test_series_addition_resequences_without_false_reorder() -> None:
    baseline = _chart(
        [
            ChartPlot(
                0,
                "barChart",
                [
                    _series(0, "Revenue", "Data!$B$2:$B$4"),
                    _series(1, "Margin", "Data!$C$2:$C$4"),
                ],
            )
        ]
    )
    new_series = _series(0, "Cost", "Data!$D$2:$D$4")
    revenue = _series(1, "Revenue", "Data!$B$2:$B$4")
    margin = _series(2, "Margin", "Data!$C$2:$C$4")
    current = _chart(
        [ChartPlot(0, "barChart", [new_series, revenue, margin])]
    )

    findings = diff_charts(_workbook(baseline), _workbook(current))

    assert len(findings) == 1
    assert "Cost" in findings[0].message and "added" in findings[0].message


def test_true_series_reorder_is_detected_once() -> None:
    revenue = _series(0, "Revenue", "Data!$B$2:$B$4")
    margin = _series(1, "Margin", "Data!$C$2:$C$4")
    baseline = _chart([ChartPlot(0, "barChart", [revenue, margin])])
    current_revenue = _series(1, "Revenue", "Data!$B$2:$B$4")
    current_margin = _series(0, "Margin", "Data!$C$2:$C$4")
    current = _chart(
        [ChartPlot(0, "barChart", [current_margin, current_revenue])]
    )

    findings = diff_charts(_workbook(baseline), _workbook(current))

    assert len(findings) == 1
    assert "series reordered" in findings[0].message


def test_ooxml_default_false_axis_flags_compare_equal() -> None:
    plot = ChartPlot(
        0,
        "barChart",
        [_series(0, "Revenue", "Data!$B$2:$B$4")],
        axis_ids=("10",),
    )
    baseline = _chart([plot], axes=[ChartAxis("10", "valAx", 0, deleted=None)])
    current = _chart([plot], axes=[ChartAxis("10", "valAx", 0, deleted=False)])

    assert diff_charts(_workbook(baseline), _workbook(current)) == []


def test_plot_grouping_message_names_the_changed_setting() -> None:
    baseline = _chart(
        [
            ChartPlot(
                0,
                "barChart",
                [_series(0, "Revenue", "Data!$B$2:$B$4")],
                grouping="clustered",
            )
        ]
    )
    current = _chart(
        [
            ChartPlot(
                0,
                "barChart",
                [_series(0, "Revenue", "Data!$B$2:$B$4")],
                grouping="stacked",
            )
        ]
    )

    findings = diff_charts(_workbook(baseline), _workbook(current))

    assert len(findings) == 1
    assert "grouping" in findings[0].message
    assert "clustered" in findings[0].message
    assert "stacked" in findings[0].message


def test_series_without_sources_degrades_chart_reference_coverage() -> None:
    chart = _chart([ChartPlot(0, "barChart", [ChartSeries(0, None, None)])])

    state, detail = chart_reference_coverage(_workbook(chart))

    assert state is CoverageState.DEGRADED
    assert "no resolvable source" in detail.lower()


def test_duplicate_titled_and_untitled_charts_match_by_semantics() -> None:
    revenue = _chart(
        [ChartPlot(0, "barChart", [_series(0, "Revenue", "Data!$B$2:$B$4")])]
    )
    margin = _chart(
        [ChartPlot(0, "lineChart", [_series(0, "Margin", "Data!$C$2:$C$4")])]
    )
    revenue.title = "Duplicate"
    margin.title = "Duplicate"
    revenue.source_index = 0
    margin.source_index = 1
    baseline = _workbook(revenue)
    baseline.charts = [revenue, margin]
    current = _workbook(margin)
    current.charts = [margin, revenue]

    assert diff_charts(baseline, current) == []

    revenue.title = None
    margin.title = None
    assert diff_charts(baseline, current) == []


def test_chart_length_mismatch_and_bubble_source_change_are_detected() -> None:
    mismatched = _chart(
        [
            ChartPlot(
                0,
                "barChart",
                [
                    _series(
                        0,
                        "Revenue",
                        "Data!$B$2:$B$4",
                        "Data!$A$2:$A$3",
                    )
                ],
            )
        ]
    )
    mismatch_findings = diff_charts(_workbook(mismatched), _workbook(mismatched))
    assert any(
        finding.finding_class is FindingClass.CHART_LENGTH_MISMATCH
        for finding in mismatch_findings
    )

    baseline_series = _series(0, "Bubbles", "Data!$B$2:$B$4")
    baseline_series.bubble_size_ref = "Data!$C$2:$C$4"
    current_series = _series(0, "Bubbles", "Data!$B$2:$B$4")
    current_series.bubble_size_ref = "Data!$D$2:$D$4"
    baseline = _chart([ChartPlot(0, "bubbleChart", [baseline_series])])
    current = _chart([ChartPlot(0, "bubbleChart", [current_series])])

    bubble_findings = diff_charts(_workbook(baseline), _workbook(current))
    assert len(bubble_findings) == 1
    assert "bubble size" in bubble_findings[0].message
