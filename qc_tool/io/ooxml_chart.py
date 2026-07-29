"""Complete Excel chart extraction from OOXML package parts."""

from __future__ import annotations

import io
import posixpath
import zipfile
from xml.etree import ElementTree

from qc_tool.io.model import (
    ChartAnchor,
    ChartAxis,
    ChartDataLabels,
    ChartDescriptor,
    ChartLegend,
    ChartPlot,
    ChartSeries,
)
from qc_tool.progress import CancellationToken, check_cancelled

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


class ChartParseError(Exception):
    """A linked chart package part could not be parsed completely."""


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _relationship_id(element: ElementTree.Element) -> str | None:
    return next(
        (
            value
            for attribute, value in element.attrib.items()
            if _local_name(attribute) == "id"
        ),
        None,
    )


def _relationship_targets(
    archive: zipfile.ZipFile,
    relationship_part: str,
    source_part: str,
) -> dict[str, str]:
    if relationship_part not in archive.namelist():
        return {}
    root = ElementTree.fromstring(archive.read(relationship_part))
    targets: dict[str, str] = {}
    for relationship in root:
        if _local_name(relationship.tag) != "Relationship":
            continue
        relationship_id = relationship.get("Id")
        target = relationship.get("Target")
        if (
            not relationship_id
            or not target
            or relationship.get("TargetMode") == "External"
        ):
            continue
        targets[relationship_id] = (
            target.lstrip("/")
            if target.startswith("/")
            else posixpath.normpath(
                posixpath.join(posixpath.dirname(source_part), target)
            )
        )
    return targets


def _rels_part(source_part: str) -> str:
    return (
        f"{posixpath.dirname(source_part)}/_rels/"
        f"{posixpath.basename(source_part)}.rels"
    )


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


def _int_text(parent: ElementTree.Element | None, name: str) -> int | None:
    element = _direct_child(parent, name)
    if element is None or element.text is None:
        return None
    try:
        return int(element.text)
    except ValueError:
        return None


def _text(parent: ElementTree.Element | None) -> str | None:
    if parent is None:
        return None
    runs = [
        item.text
        for item in parent.iter()
        if _local_name(item.tag) == "t" and item.text
    ]
    if not runs:
        runs = [
            item.text
            for item in parent.iter()
            if _local_name(item.tag) == "v" and item.text
        ]
    return "".join(runs) or None


def _formula(parent: ElementTree.Element | None) -> str | None:
    element = _descendant(parent, "f")
    return element.text if element is not None and element.text else None


def _parse_data_labels(
    element: ElementTree.Element | None,
) -> ChartDataLabels | None:
    if element is None:
        return None
    number_format = _value(element, "numFmt")
    if number_format is None:
        number_format_element = _direct_child(element, "numFmt")
        if number_format_element is not None:
            number_format = number_format_element.get("formatCode")
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
        number_format=number_format,
        separator=separator.text if separator is not None else None,
    )


def _parse_series(
    element: ElementTree.Element,
    *,
    plot_index: int,
    source_index: int,
    chart_source_id: str,
) -> ChartSeries:
    index_text = _value(element, "idx")
    order_text = _value(element, "order")
    tx = _direct_child(element, "tx")
    series_source_id = f"{chart_source_id}/plot[{plot_index}]/series[{source_index}]"
    return ChartSeries(
        index=int(index_text) if index_text and index_text.isdigit() else source_index,
        order=int(order_text) if order_text and order_text.isdigit() else None,
        values_ref=_formula(
            _direct_child(element, "val") or _direct_child(element, "yVal")
        ),
        categories_ref=_formula(
            _direct_child(element, "cat") or _direct_child(element, "xVal")
        ),
        name_ref=_formula(tx),
        name_text=_text(tx),
        plot_index=plot_index,
        source_index=source_index,
        source_id=series_source_id,
        bubble_size_ref=_formula(_direct_child(element, "bubbleSize")),
        data_labels=_parse_data_labels(_direct_child(element, "dLbls")),
    )


def _assign_axis_groups(plots: list[ChartPlot]) -> None:
    known: list[tuple[str, ...]] = []
    for plot in plots:
        if not plot.axis_ids:
            plot.axis_group = "none"
            continue
        if plot.axis_ids not in known:
            known.append(plot.axis_ids)
        index = known.index(plot.axis_ids)
        plot.axis_group = "primary" if index == 0 else f"secondary-{index}"


def _parse_plots(
    plot_area: ElementTree.Element,
    chart_source_id: str,
) -> list[ChartPlot]:
    plots: list[ChartPlot] = []
    for child in plot_area:
        chart_type = _local_name(child.tag)
        if chart_type not in _PLOT_TYPES:
            continue
        plot_index = len(plots)
        source_id = f"{chart_source_id}/plot[{plot_index}]"
        series_elements = [
            item for item in child if _local_name(item.tag) == "ser"
        ]
        series = [
            _parse_series(
                item,
                plot_index=plot_index,
                source_index=series_index,
                chart_source_id=chart_source_id,
            )
            for series_index, item in enumerate(series_elements)
        ]
        plots.append(
            ChartPlot(
                index=plot_index,
                chart_type=chart_type,
                series=series,
                axis_ids=tuple(
                    item.get("val", "")
                    for item in child
                    if _local_name(item.tag) == "axId" and item.get("val")
                ),
                grouping=_value(child, "grouping"),
                direction=_value(child, "barDir"),
                style=(
                    _value(child, "scatterStyle") or _value(child, "radarStyle")
                ),
                data_labels=_parse_data_labels(_direct_child(child, "dLbls")),
                source_id=source_id,
            )
        )
    _assign_axis_groups(plots)
    return plots


def _parse_axis(element: ElementTree.Element, source_index: int) -> ChartAxis:
    scaling = _direct_child(element, "scaling")
    num_fmt = _direct_child(element, "numFmt")
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
        number_format=num_fmt.get("formatCode") if num_fmt is not None else None,
        display_unit=display_unit,
        major_tick_mark=_value(element, "majorTickMark"),
        minor_tick_mark=_value(element, "minorTickMark"),
        title=_text(_direct_child(element, "title")),
        deleted=_bool_value(element, "delete"),
    )


def _parse_legend(chart: ElementTree.Element) -> ChartLegend | None:
    element = _direct_child(chart, "legend")
    if element is None:
        return None
    return ChartLegend(
        position=_value(element, "legendPos"),
        overlay=_bool_value(element, "overlay"),
        deleted=_bool_value(element, "delete"),
    )


def _parse_marker(
    anchor: ElementTree.Element,
    marker_name: str,
) -> tuple[int | None, int | None, int | None, int | None]:
    marker = _direct_child(anchor, marker_name)
    return (
        _int_text(marker, "col"),
        _int_text(marker, "row"),
        _int_text(marker, "colOff"),
        _int_text(marker, "rowOff"),
    )


def _parse_anchor(element: ElementTree.Element) -> ChartAnchor:
    anchor_type = _local_name(element.tag)
    from_col, from_row, from_col_offset, from_row_offset = _parse_marker(
        element, "from"
    )
    to_col, to_row, to_col_offset, to_row_offset = _parse_marker(element, "to")
    position = _direct_child(element, "pos")
    extent = _direct_child(element, "ext")

    def integer_attribute(
        item: ElementTree.Element | None, attribute: str
    ) -> int | None:
        value = item.get(attribute) if item is not None else None
        try:
            return int(value) if value is not None else None
        except ValueError:
            return None

    return ChartAnchor(
        anchor_type=anchor_type,
        from_col=from_col,
        from_row=from_row,
        from_col_offset=from_col_offset,
        from_row_offset=from_row_offset,
        to_col=to_col,
        to_row=to_row,
        to_col_offset=to_col_offset,
        to_row_offset=to_row_offset,
        x=integer_attribute(position, "x"),
        y=integer_attribute(position, "y"),
        width=integer_attribute(extent, "cx"),
        height=integer_attribute(extent, "cy"),
    )


def _parse_chart(
    archive: zipfile.ZipFile,
    chart_part: str,
    *,
    sheet: str,
    source_index: int,
    source_name: str | None,
    anchor: ChartAnchor,
) -> ChartDescriptor:
    if chart_part not in archive.namelist():
        raise ChartParseError(f"linked chart part {chart_part!r} is missing")
    root = ElementTree.fromstring(archive.read(chart_part))
    chart = _descendant(root, "chart")
    plot_area = _direct_child(chart, "plotArea")
    if chart is None or plot_area is None:
        raise ChartParseError(f"chart part {chart_part!r} has no chart/plot area")
    source_id = f"{sheet}!chart[{source_index}]@{anchor.signature}"
    plots = _parse_plots(plot_area, source_id)
    if not plots:
        raise ChartParseError(f"chart part {chart_part!r} has no supported plot")
    axes = [
        _parse_axis(item, axis_index)
        for axis_index, item in enumerate(plot_area)
        if _local_name(item.tag) in _AXIS_TYPES
    ]
    return ChartDescriptor(
        sheet=sheet,
        title=_text(_direct_child(chart, "title")),
        chart_type=plots[0].chart_type,
        series=[series for plot in plots for series in plot.series],
        source_index=source_index,
        source_id=source_id,
        source_name=source_name,
        source_part=chart_part,
        plots=plots,
        axes=axes,
        legend=_parse_legend(chart),
        anchor=anchor,
        display_blanks_as=_value(chart, "dispBlanksAs"),
    )


def _drawing_charts(
    archive: zipfile.ZipFile,
    drawing_part: str,
    *,
    sheet: str,
    start_index: int,
    cancellation_token: CancellationToken | None = None,
) -> list[ChartDescriptor]:
    if drawing_part not in archive.namelist():
        raise ChartParseError(f"linked drawing part {drawing_part!r} is missing")
    relationships = _relationship_targets(
        archive,
        _rels_part(drawing_part),
        drawing_part,
    )
    root = ElementTree.fromstring(archive.read(drawing_part))
    charts: list[ChartDescriptor] = []
    for anchor_element in root:
        chart_element = _descendant(anchor_element, "chart")
        if chart_element is None:
            continue
        check_cancelled(cancellation_token)
        relationship_id = _relationship_id(chart_element)
        chart_part = relationships.get(relationship_id or "")
        if chart_part is None:
            raise ChartParseError(
                f"drawing {drawing_part!r} has an unresolved chart relationship"
            )
        source_name_element = _descendant(anchor_element, "cNvPr")
        charts.append(
            _parse_chart(
                archive,
                chart_part,
                sheet=sheet,
                source_index=start_index + len(charts),
                source_name=(
                    source_name_element.get("name")
                    if source_name_element is not None
                    else None
                ),
                anchor=_parse_anchor(anchor_element),
            )
        )
    return charts


def _worksheet_drawing_ids(
    archive: zipfile.ZipFile,
    sheet_part: str,
) -> list[str]:
    relationship_ids: list[str] = []
    with archive.open(sheet_part) as stream:
        for _event, element in ElementTree.iterparse(stream, events=("end",)):
            if _local_name(element.tag) == "drawing":
                relationship_id = _relationship_id(element)
                if relationship_id:
                    relationship_ids.append(relationship_id)
            element.clear()
    return relationship_ids


def parse_ooxml_charts(
    data: bytes,
    *,
    cancellation_token: CancellationToken | None = None,
) -> list[ChartDescriptor]:
    """Return every worksheet-linked chart using raw package relationships."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            if "xl/workbook.xml" not in archive.namelist():
                raise ChartParseError("OOXML package has no workbook part")
            workbook_relationships = _relationship_targets(
                archive,
                "xl/_rels/workbook.xml.rels",
                "xl/workbook.xml",
            )
            workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
            charts: list[ChartDescriptor] = []
            sheet_counts: dict[str, int] = {}
            for sheet_element in workbook.iter():
                if _local_name(sheet_element.tag) != "sheet":
                    continue
                check_cancelled(cancellation_token)
                sheet_name = sheet_element.get("name")
                relationship_id = _relationship_id(sheet_element)
                sheet_part = workbook_relationships.get(relationship_id or "")
                if not sheet_name or sheet_part not in archive.namelist():
                    continue
                sheet_relationships = _relationship_targets(
                    archive,
                    _rels_part(sheet_part),
                    sheet_part,
                )
                for drawing_id in _worksheet_drawing_ids(archive, sheet_part):
                    drawing_part = sheet_relationships.get(drawing_id)
                    if drawing_part is None:
                        raise ChartParseError(
                            f"sheet {sheet_name!r} has an unresolved drawing relationship"
                        )
                    parsed = _drawing_charts(
                        archive,
                        drawing_part,
                        sheet=sheet_name,
                        start_index=sheet_counts.get(sheet_name, 0),
                        cancellation_token=cancellation_token,
                    )
                    charts.extend(parsed)
                    sheet_counts[sheet_name] = sheet_counts.get(sheet_name, 0) + len(
                        parsed
                    )
            return charts
    except (ElementTree.ParseError, zipfile.BadZipFile) as exc:
        raise ChartParseError(f"malformed OOXML chart package: {exc}") from exc
