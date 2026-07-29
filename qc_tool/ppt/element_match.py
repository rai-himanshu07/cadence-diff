"""Semantic one-to-one matching inside an already matched slide pair."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from typing import TypeVar

from qc_tool.ppt.model import (
    ChartContent,
    PptChartPlot,
    PptChartSeries,
    SlideContent,
    TableContent,
)

_T = TypeVar("_T")


@dataclass(slots=True)
class NestedMatching:
    pairs: list[tuple[object, object]] = field(default_factory=list)
    removed: list[object] = field(default_factory=list)
    added: list[object] = field(default_factory=list)
    reordered: list[tuple[object, object]] = field(default_factory=list)


@dataclass(slots=True)
class PlotMatching:
    pairs: list[tuple[PptChartPlot, PptChartPlot]] = field(default_factory=list)
    removed: list[PptChartPlot] = field(default_factory=list)
    added: list[PptChartPlot] = field(default_factory=list)
    reordered: list[tuple[PptChartPlot, PptChartPlot]] = field(default_factory=list)


@dataclass(slots=True)
class SeriesMatching:
    pairs: list[tuple[PptChartSeries, PptChartSeries]] = field(default_factory=list)
    removed: list[PptChartSeries] = field(default_factory=list)
    added: list[PptChartSeries] = field(default_factory=list)
    reordered: list[tuple[PptChartSeries, PptChartSeries]] = field(default_factory=list)


@dataclass(slots=True)
class SlideElementMatching:
    table_pairs: list[tuple[TableContent, TableContent]] = field(default_factory=list)
    removed_tables: list[TableContent] = field(default_factory=list)
    added_tables: list[TableContent] = field(default_factory=list)
    reordered_tables: list[tuple[TableContent, TableContent]] = field(default_factory=list)
    chart_pairs: list[tuple[ChartContent, ChartContent]] = field(default_factory=list)
    removed_charts: list[ChartContent] = field(default_factory=list)
    added_charts: list[ChartContent] = field(default_factory=list)
    reordered_charts: list[tuple[ChartContent, ChartContent]] = field(default_factory=list)


def _consume_unique_matches(
    baseline: list[_T],
    current: list[_T],
    key: Callable[[_T], object | None],
) -> list[tuple[_T, _T]]:
    baseline_keys: dict[object, list[_T]] = {}
    current_keys: dict[object, list[_T]] = {}
    for item in baseline:
        value = key(item)
        if value is not None:
            baseline_keys.setdefault(value, []).append(item)
    for item in current:
        value = key(item)
        if value is not None:
            current_keys.setdefault(value, []).append(item)
    pairs: list[tuple[_T, _T]] = []
    for value in sorted(baseline_keys.keys() & current_keys.keys(), key=repr):
        baseline_items = baseline_keys[value]
        current_items = current_keys[value]
        if len(baseline_items) != 1 or len(current_items) != 1:
            continue
        baseline_item = baseline_items[0]
        current_item = current_items[0]
        pairs.append((baseline_item, current_item))
        baseline.remove(baseline_item)
        current.remove(current_item)
    return pairs


def _consume_ordered_matches(
    baseline: list[_T],
    current: list[_T],
    predicate: Callable[[_T, _T], bool],
) -> list[tuple[_T, _T]]:
    pairs: list[tuple[_T, _T]] = []
    for baseline_item in list(baseline):
        current_item = next(
            (item for item in current if predicate(baseline_item, item)),
            None,
        )
        if current_item is None:
            continue
        pairs.append((baseline_item, current_item))
        baseline.remove(baseline_item)
        current.remove(current_item)
    return pairs


def _longest_increasing_subsequence(values: list[int]) -> set[int]:
    if not values:
        return set()
    lengths = [1] * len(values)
    parents = [-1] * len(values)
    for index in range(1, len(values)):
        for previous in range(index):
            if values[previous] < values[index] and lengths[previous] + 1 > lengths[index]:
                lengths[index] = lengths[previous] + 1
                parents[index] = previous
    best = max(range(len(values)), key=lambda index: lengths[index])
    keep: set[int] = set()
    while best != -1:
        keep.add(best)
        best = parents[best]
    return keep


def _reordered_pairs(
    pairs: list[tuple[_T, _T]],
    baseline_position: Callable[[_T], int],
    current_position: Callable[[_T], int],
) -> list[tuple[_T, _T]]:
    ordered = sorted(pairs, key=lambda pair: baseline_position(pair[0]))
    positions = [current_position(current) for _, current in ordered]
    stable = _longest_increasing_subsequence(positions)
    return [pair for index, pair in enumerate(ordered) if index not in stable]


def _table_header(table: TableContent) -> tuple[str, ...] | None:
    if not table.rows:
        return None
    return tuple(cell.strip().casefold() for cell in table.rows[0])


def _table_row_labels(table: TableContent) -> tuple[str, ...] | None:
    labels = tuple(
        row[0].strip().casefold()
        for row in table.rows[1:]
        if row and row[0].strip()
    )
    return labels or None


def _table_geometry(table: TableContent) -> tuple[int, int, int, int]:
    return table.left, table.top, table.width, table.height


def _chart_title(chart: ChartContent) -> str | None:
    return chart.title.casefold() if chart.title else None


def _series_name(series: PptChartSeries) -> str | None:
    return series.name.casefold() if series.name else None


def _category_signature(series: PptChartSeries) -> tuple[str, ...] | None:
    return tuple(series.categories) or None


def _plot_signature(plot: PptChartPlot) -> tuple[str, ...] | None:
    names = tuple(sorted(name for series in plot.series if (name := _series_name(series))))
    return names or None


def _chart_signature(chart: ChartContent) -> tuple[tuple[str, ...], ...] | None:
    signatures = tuple(sorted(
        signature for plot in chart.plots if (signature := _plot_signature(plot))
    ))
    return signatures or None


def _chart_geometry(chart: ChartContent) -> tuple[int, int, int, int]:
    return chart.left, chart.top, chart.width, chart.height


def match_slide_elements(
    baseline: SlideContent,
    current: SlideContent,
) -> SlideElementMatching:
    matching = SlideElementMatching()
    baseline_tables = list(baseline.tables)
    current_tables = list(current.tables)
    matching.table_pairs.extend(
        _consume_unique_matches(baseline_tables, current_tables, _table_header)
    )
    matching.table_pairs.extend(
        _consume_unique_matches(baseline_tables, current_tables, _table_row_labels)
    )
    matching.table_pairs.extend(
        _consume_unique_matches(baseline_tables, current_tables, _table_geometry)
    )
    matching.table_pairs.extend(
        _consume_unique_matches(
            baseline_tables,
            current_tables,
            lambda table: table.source_index,
        )
    )
    matching.removed_tables = baseline_tables
    matching.added_tables = current_tables
    matching.reordered_tables = _reordered_pairs(
        matching.table_pairs,
        lambda table: table.z_order,
        lambda table: table.z_order,
    )

    baseline_charts = list(baseline.charts)
    current_charts = list(current.charts)
    matching.chart_pairs.extend(
        _consume_unique_matches(baseline_charts, current_charts, _chart_title)
    )
    matching.chart_pairs.extend(
        _consume_unique_matches(baseline_charts, current_charts, _chart_signature)
    )
    matching.chart_pairs.extend(
        _consume_unique_matches(baseline_charts, current_charts, _chart_geometry)
    )
    matching.chart_pairs.extend(
        _consume_ordered_matches(
            baseline_charts,
            current_charts,
            lambda baseline_chart, current_chart: (
                _chart_title(baseline_chart) is not None
                and _chart_title(baseline_chart) == _chart_title(current_chart)
            ),
        )
    )
    matching.chart_pairs.extend(
        _consume_ordered_matches(
            baseline_charts,
            current_charts,
            lambda baseline_chart, current_chart: (
                _chart_signature(baseline_chart) is not None
                and _chart_signature(baseline_chart) == _chart_signature(current_chart)
            ),
        )
    )
    matching.chart_pairs.extend(
        _consume_unique_matches(
            baseline_charts,
            current_charts,
            lambda chart: chart.source_index,
        )
    )
    matching.removed_charts = baseline_charts
    matching.added_charts = current_charts
    matching.reordered_charts = _reordered_pairs(
        matching.chart_pairs,
        lambda chart: chart.z_order,
        lambda chart: chart.z_order,
    )
    return matching


def match_plots(
    baseline: ChartContent,
    current: ChartContent,
) -> PlotMatching:
    baseline_plots = list(baseline.plots)
    current_plots = list(current.plots)
    matching = PlotMatching()
    matching.pairs.extend(
        _consume_unique_matches(baseline_plots, current_plots, _plot_signature)
    )
    matching.pairs.extend(
        _consume_unique_matches(
            baseline_plots,
            current_plots,
            lambda plot: plot.index,
        )
    )
    matching.removed = baseline_plots
    matching.added = current_plots
    matching.reordered = _reordered_pairs(
        matching.pairs,
        lambda plot: plot.index,
        lambda plot: plot.index,
    )
    return matching


def match_series(
    baseline: PptChartPlot,
    current: PptChartPlot,
) -> SeriesMatching:
    baseline_series = list(baseline.series)
    current_series = list(current.series)
    matching = SeriesMatching()
    matching.pairs.extend(
        _consume_unique_matches(baseline_series, current_series, _series_name)
    )
    matching.pairs.extend(
        _consume_unique_matches(
            baseline_series,
            current_series,
            _category_signature,
        )
    )
    matching.pairs.extend(
        _consume_unique_matches(
            baseline_series,
            current_series,
            lambda series: series.source_index,
        )
    )
    matching.removed = baseline_series
    matching.added = current_series
    matching.reordered = _reordered_pairs(
        matching.pairs,
        lambda series: series.order if series.order is not None else series.source_index,
        lambda series: series.order if series.order is not None else series.source_index,
    )
    return matching
