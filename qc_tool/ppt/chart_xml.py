"""Complete PowerPoint native-chart extraction from raw chart XML parts."""

from __future__ import annotations

from xml.etree import ElementTree

from qc_tool.io.model import ChartAxis, ChartDataLabels, ChartLegend
from qc_tool.ppt.model import (
    ChartContent,
    PptChartLabel,
    PptChartPlot,
    PptChartSeries,
)

_PLOT_TYPES = frozenset(
    {
        "areaChart",
        "area3DChart",
        "barChart",
        "bar3DChart",
        "bubbleChart",
        "doughnutChart",
        "lineChart",
        "line3DChart",
        "ofPieChart",
        "pieChart",
        "pie3DChart",
        "radarChart",
        "scatterChart",
        "stockChart",
        "surfaceChart",
        "surface3DChart",
    }
)
_AXIS_TYPES = frozenset({"catAx", "dateAx", "serAx", "valAx"})


class PptChartParseError(Exception):
    """A native PowerPoint chart part could not be represented completely."""


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _direct_child(
    parent: ElementTree.Element | None, name: str
) -> ElementTree.Element | None:
    if parent is None:
        return None
    return next((child for child in parent if _local_name(child.tag) == name), None)


def _descendant(
    parent: ElementTree.Element | None, name: str
) -> ElementTree.Element | None:
    if parent is None:
        return None
    return next((item for item in parent.iter() if _local_name(item.tag) == name), None)


def _value(parent: ElementTree.Element | None, name: str) -> str | None:
    element = _direct_child(parent, name)
    return element.get("val") if element is not None else None


def _float_value(parent: ElementTree.Element | None, name: str) -> float | None:
    value = _value(parent, name)
    if value is None:
        return None
    try:
        return float(value)
    except ValueError:
        return None


def _bool_value(parent: ElementTree.Element | None, name: str) -> bool | None:
    element = _direct_child(parent, name)
    if element is None:
        return None
    value = element.get("val")
    return True if value is None else value not in {"0", "false", "off"}


def _formula(parent: ElementTree.Element | None) -> str | None:
    element = _descendant(parent, "f")
    return element.text if element is not None and element.text else None


def _text(parent: ElementTree.Element | None) -> str | None:
    if parent is None:
        return None
    values = [
        item.text
        for item in parent.iter()
        if _local_name(item.tag) in {"t", "v"} and item.text
    ]
    return "".join(values) or None


def _cache(parent: ElementTree.Element | None) -> ElementTree.Element | None:
    if parent is None:
        return None
    return next(
        (
            item
            for item in parent.iter()
            if _local_name(item.tag) in {"strCache", "numCache", "strLit", "numLit"}
        ),
        None,
    )


def _cache_size(cache: ElementTree.Element) -> tuple[int, list[ElementTree.Element]]:
    count_text = _value(cache, "ptCount")
    points = [item for item in cache if _local_name(item.tag) == "pt"]
    indexes = [
        int(item.get("idx", "0"))
        for item in points
        if item.get("idx", "0").isdigit()
    ]
    size = int(count_text) if count_text and count_text.isdigit() else 0
    if indexes:
        size = max(size, max(indexes) + 1)
    return size, points


def _text_cache(parent: ElementTree.Element | None) -> list[str]:
    cache = _cache(parent)
    if cache is None:
        return []
    size, points = _cache_size(cache)
    output = [""] * size
    for point in points:
        index_text = point.get("idx", "0")
        value = _direct_child(point, "v")
        if index_text.isdigit() and value is not None and value.text is not None:
            output[int(index_text)] = value.text
    return output


def _numeric_cache(parent: ElementTree.Element | None) -> list[float | None]:
    cache = _cache(parent)
    if cache is None:
        return []
    size, points = _cache_size(cache)
    output: list[float | None] = [None] * size
    for point in points:
        index_text = point.get("idx", "0")
        value = _direct_child(point, "v")
        if not index_text.isdigit() or value is None or value.text is None:
            continue
        try:
            output[int(index_text)] = float(value.text)
        except ValueError:
            output[int(index_text)] = None
    return output


def _parse_data_labels(
    element: ElementTree.Element | None,
) -> ChartDataLabels | None:
    if element is None:
        return None
    number_format = _direct_child(element, "numFmt")
    separator = _direct_child(element, "separator")
    return ChartDataLabels(
        position=_value(element, "dLblPos"),
        show_value=_bool_value(element, "showVal"),
        show_category_name=_bool_value(element, "showCatName"),
        show_series_name=_bool_value(element, "showSerName"),
        show_percent=_bool_value(element, "showPercent"),
        show_legend_key=_bool_value(element, "showLegendKey"),
        show_bubble_size=_bool_value(element, "showBubbleSize"),
        show_leader_lines=_bool_value(element, "showLeaderLines"),
        number_format=(
            number_format.get("formatCode") if number_format is not None else None
        ),
        separator=separator.text if separator is not None else None,
    )


def _series_source(
    element: ElementTree.Element,
    *names: str,
) -> ElementTree.Element | None:
    for name in names:
        source = _direct_child(element, name)
        if source is not None:
            return source
    return None


def _parse_series(
    element: ElementTree.Element,
    *,
    plot_index: int,
    source_index: int,
    chart_source_id: str,
) -> PptChartSeries:
    index_text = _value(element, "idx")
    order_text = _value(element, "order")
    tx = _direct_child(element, "tx")
    categories_source = _series_source(element, "cat", "xVal")
    values_source = _series_source(element, "val", "yVal")
    return PptChartSeries(
        index=int(index_text) if index_text and index_text.isdigit() else source_index,
        order=int(order_text) if order_text and order_text.isdigit() else None,
        name=_text(tx),
        categories=_text_cache(categories_source),
        values=_numeric_cache(values_source),
        plot_index=plot_index,
        source_index=source_index,
        source_id=f"{chart_source_id}/plot[{plot_index}]/series[{source_index}]",
        name_ref=_formula(tx),
        categories_ref=_formula(categories_source),
        values_ref=_formula(values_source),
        bubble_size_ref=_formula(_direct_child(element, "bubbleSize")),
        data_labels=_parse_data_labels(_direct_child(element, "dLbls")),
    )


def _assign_axis_groups(plots: list[PptChartPlot]) -> None:
    groups: list[tuple[str, ...]] = []
    for plot in plots:
        if not plot.axis_ids:
            plot.axis_group = "none"
            continue
        group_key = tuple(sorted(plot.axis_ids))
        if group_key not in groups:
            groups.append(group_key)
        index = groups.index(group_key)
        plot.axis_group = "primary" if index == 0 else f"secondary-{index}"


def _visible_labels(plot: PptChartPlot) -> list[PptChartLabel]:
    output: list[PptChartLabel] = []
    for series in plot.series:
        labels = series.data_labels or plot.data_labels
        if labels is None:
            continue
        count = max(len(series.categories), len(series.values))
        for index in range(count):
            category = (
                series.categories[index] if index < len(series.categories) else ""
            )
            value = series.values[index] if index < len(series.values) else None
            if not category and value is None:
                continue
            parts: list[str] = []
            if labels.show_series_name and series.name:
                parts.append(series.name)
            if labels.show_category_name and category:
                parts.append(category)
            if labels.show_value and value is not None:
                parts.append(str(value))
            if parts:
                separator = labels.separator if labels.separator is not None else "\n"
                output.append(
                    PptChartLabel(
                        text=separator.join(parts),
                        series_source_index=series.source_index,
                        category_index=index,
                        category=category,
                    )
                )
    return output


def _parse_plots(
    plot_area: ElementTree.Element,
    chart_source_id: str,
) -> list[PptChartPlot]:
    plots: list[PptChartPlot] = []
    for child in plot_area:
        chart_type = _local_name(child.tag)
        if chart_type not in _PLOT_TYPES:
            continue
        plot_index = len(plots)
        series_elements = [item for item in child if _local_name(item.tag) == "ser"]
        plot = PptChartPlot(
            index=plot_index,
            chart_type=chart_type,
            series=[
                _parse_series(
                    item,
                    plot_index=plot_index,
                    source_index=series_index,
                    chart_source_id=chart_source_id,
                )
                for series_index, item in enumerate(series_elements)
            ],
            axis_ids=tuple(
                item.get("val", "")
                for item in child
                if _local_name(item.tag) == "axId" and item.get("val")
            ),
            grouping=_value(child, "grouping"),
            direction=_value(child, "barDir"),
            style=_value(child, "scatterStyle") or _value(child, "radarStyle"),
            data_labels=_parse_data_labels(_direct_child(child, "dLbls")),
            source_id=f"{chart_source_id}/plot[{plot_index}]",
        )
        plot.visible_labels = _visible_labels(plot)
        plot.visible_label_values = [label.text for label in plot.visible_labels]
        plots.append(plot)
    _assign_axis_groups(plots)
    return plots


def _parse_axis(element: ElementTree.Element, source_index: int) -> ChartAxis:
    scaling = _direct_child(element, "scaling")
    number_format = _direct_child(element, "numFmt")
    display_units = _direct_child(element, "dispUnits")
    display_unit = None
    if display_units is not None:
        display_unit = _value(display_units, "builtInUnit")
        if display_unit is None:
            custom = _float_value(display_units, "custUnit")
            display_unit = str(custom) if custom is not None else None
    return ChartAxis(
        axis_id=_value(element, "axId") or f"unknown-{source_index}",
        axis_type=_local_name(element.tag),
        source_index=source_index,
        position=_value(element, "axPos"),
        cross_axis_id=_value(element, "crossAx"),
        minimum=_float_value(scaling, "min"),
        maximum=_float_value(scaling, "max"),
        major_unit=_float_value(element, "majorUnit"),
        minor_unit=_float_value(element, "minorUnit"),
        log_base=_float_value(scaling, "logBase"),
        orientation=_value(scaling, "orientation"),
        crosses=_value(element, "crosses"),
        crosses_at=_float_value(element, "crossesAt"),
        number_format=(
            number_format.get("formatCode") if number_format is not None else None
        ),
        display_unit=display_unit,
        major_tick_mark=_value(element, "majorTickMark"),
        minor_tick_mark=_value(element, "minorTickMark"),
        title=_text(_direct_child(element, "title")),
        deleted=_bool_value(element, "delete"),
    )


def _parse_legend(chart: ElementTree.Element) -> ChartLegend | None:
    legend = _direct_child(chart, "legend")
    if legend is None:
        return None
    return ChartLegend(
        position=_value(legend, "legendPos"),
        overlay=_bool_value(legend, "overlay"),
        deleted=_bool_value(legend, "delete"),
    )


def parse_ppt_chart(
    data: bytes,
    *,
    source_id: str,
    source_index: int,
    shape_id: int,
    name: str,
    left: int,
    top: int,
    width: int,
    height: int,
    z_order: int,
    source_part: str | None,
) -> ChartContent:
    """Parse one raw native-chart part into every plot and series."""
    try:
        root = ElementTree.fromstring(data)
    except ElementTree.ParseError as exc:
        raise PptChartParseError(f"malformed chart XML: {exc}") from exc
    chart = _descendant(root, "chart")
    plot_area = _direct_child(chart, "plotArea")
    if chart is None or plot_area is None:
        raise PptChartParseError("chart part has no chart/plot area")
    plots = _parse_plots(plot_area, source_id)
    if not plots:
        raise PptChartParseError("chart part has no supported plot")
    axes = [
        _parse_axis(item, axis_index)
        for axis_index, item in enumerate(plot_area)
        if _local_name(item.tag) in _AXIS_TYPES
    ]
    captured_axis_ids = {axis.axis_id for axis in axes}
    missing_axis_ids = sorted(
        {
            axis_id
            for plot in plots
            for axis_id in plot.axis_ids
            if axis_id not in captured_axis_ids
        }
    )
    if missing_axis_ids:
        raise PptChartParseError(
            f"chart references uncaptured axis IDs {missing_axis_ids}"
        )
    first_plot = plots[0]
    categories = first_plot.series[0].categories if first_plot.series else []
    return ChartContent(
        chart_type=first_plot.chart_type,
        categories=categories,
        series=[(series.name, series.values) for series in first_plot.series],
        source_id=source_id,
        source_index=source_index,
        shape_id=shape_id,
        name=name,
        title=_text(_direct_child(chart, "title")),
        plots=plots,
        axes=axes,
        legend=_parse_legend(chart),
        display_blanks_as=_value(chart, "dispBlanksAs"),
        left=left,
        top=top,
        width=width,
        height=height,
        z_order=z_order,
        source_part=source_part,
    )
