"""Semantic diff for tables and complete native charts within matched slides."""

from __future__ import annotations

from collections import Counter

from qc_tool.availability import ppt_blank_allowed
from qc_tool.config.profile import PptProfile
from qc_tool.excel.periods import Period, is_period_after, parse_period
from qc_tool.excel.references import is_pure_range_extension
from qc_tool.findings import Finding, FindingClass, FindingExpectedReason
from qc_tool.io.model import ChartAxis, ChartDataLabels, ChartLegend
from qc_tool.ppt.element_match import match_plots, match_series, match_slide_elements
from qc_tool.ppt.model import (
    ChartContent,
    PptChartPlot,
    PptChartSeries,
    SlideContent,
    TableContent,
)

_GEOMETRY_OFFSET_TOLERANCE = 50_000
_GEOMETRY_SIZE_RATIO = 0.02


def _latest_periods(labels: list[str]) -> dict[str, Period]:
    latest: dict[str, Period] = {}
    for period in map(parse_period, labels):
        if period is None:
            continue
        previous = latest.get(period.kind)
        if previous is None or is_period_after(period, previous):
            latest[period.kind] = period
    return latest


def _table_label(table: TableContent) -> str:
    if len(table.rows) > 1 and table.rows[1]:
        return table.rows[1][0] or f"table {table.source_index + 1}"
    return table.name or f"table {table.source_index + 1}"


def _table_schema(table: TableContent) -> tuple[tuple[str, ...], tuple[str, ...]]:
    headers = tuple(cell.strip().casefold() for cell in table.rows[0]) if table.rows else ()
    row_labels = tuple(
        row[0].strip().casefold() for row in table.rows[1:] if row
    )
    return headers, row_labels


def _expected_table_header_growth(
    baseline: TableContent,
    current: TableContent,
) -> bool:
    baseline_headers, baseline_rows = _table_schema(baseline)
    current_headers, current_rows = _table_schema(current)
    if baseline_rows != current_rows or len(current_headers) <= len(baseline_headers):
        return False
    if current_headers[: len(baseline_headers)] != baseline_headers:
        return False
    latest = _latest_periods(list(baseline_headers))
    for label in current_headers[len(baseline_headers) :]:
        period = parse_period(label)
        if period is None:
            return False
        previous = latest.get(period.kind)
        if previous is not None and not is_period_after(period, previous):
            return False
        latest[period.kind] = period
    return True


def _table_findings(
    slide: SlideContent,
    baseline: TableContent,
    current: TableContent,
    profile: PptProfile,
) -> list[Finding]:
    findings: list[Finding] = []
    display = slide.display_name
    baseline_rows = baseline.rows
    current_rows = current.rows
    baseline_headers = baseline_rows[0] if baseline_rows else []
    current_headers = current_rows[0] if current_rows else []
    baseline_schema = _table_schema(baseline)
    current_schema = _table_schema(current)
    shared_header_count = min(len(baseline_schema[0]), len(current_schema[0]))
    shared_row_count = min(len(baseline_schema[1]), len(current_schema[1]))
    if (
        baseline_schema[0][:shared_header_count]
        != current_schema[0][:shared_header_count]
        or baseline_schema[1][:shared_row_count]
        != current_schema[1][:shared_row_count]
    ):
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_TABLE_STRUCTURE_CHANGED,
                slide=display,
                element=_table_label(current),
                message=f"{display}: table schema changed",
            )
        )
    baseline_latest = _latest_periods(baseline_headers)
    shared_rows = min(len(baseline_rows), len(current_rows))
    shared_columns = min(len(baseline_headers), len(current_headers))
    for row_index in range(shared_rows):
        for column_index in range(shared_columns):
            baseline_value = baseline_rows[row_index][column_index]
            current_value = current_rows[row_index][column_index]
            if baseline_value == current_value:
                continue
            if row_index == 0 or column_index == 0:
                continue
            row_label = (
                current_rows[row_index][0]
                if current_rows[row_index]
                else str(row_index)
            )
            column_label = (
                current_headers[column_index]
                if column_index < len(current_headers)
                else str(column_index)
            )
            if (
                not current_value.strip()
                and ppt_blank_allowed(
                    profile,
                    slide=display,
                    scope="table",
                    element=current.name or f"table[{current.source_index}]",
                    series=row_label,
                    period_label=column_label,
                )
            ):
                continue
            findings.append(
                Finding(
                    artifact="ppt",
                    finding_class=FindingClass.TABLE_VALUE_CHANGED,
                    slide=display,
                    element=f"{row_label} / {column_label}",
                    baseline_value=baseline_value,
                    current_value=current_value,
                    message=(
                        f"{display}: table cell {row_label} / {column_label} changed"
                    ),
                )
            )
    for column_index in range(shared_columns, len(current_headers)):
        header = current_headers[column_index]
        period = parse_period(header)
        previous = baseline_latest.get(period.kind) if period is not None else None
        expected = bool(
            period is not None
            and (previous is None or is_period_after(period, previous))
        )
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.TABLE_VALUE_CHANGED,
                expected_reason=(
                    FindingExpectedReason.CADENCE_EXTENSION if expected else None
                ),
                slide=display,
                element=header,
                message=(
                    f"{display}: table column {header!r} "
                    + ("appended (new cycle)" if expected else "added unexpectedly")
                ),
            )
        )
        for row_index, row in enumerate(current_rows[1:], start=2):
            value = row[column_index].strip() if column_index < len(row) else ""
            if value:
                continue
            row_label = row[0] if row else str(row_index)
            if ppt_blank_allowed(
                profile,
                slide=display,
                scope="table",
                element=current.name or f"table[{current.source_index}]",
                series=row_label,
                period_label=header,
            ):
                continue
            findings.append(
                Finding(
                    artifact="ppt",
                    finding_class=FindingClass.PPT_TABLE_BLANK,
                    slide=display,
                    element=f"{row_label} / {header}",
                    message=(
                        f"{display}: blank value in added table column "
                        f"{header!r} for {row_label!r}"
                    ),
                )
            )
    for column_index in range(shared_columns, len(baseline_headers)):
        header = baseline_headers[column_index]
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.TABLE_VALUE_CHANGED,
                slide=display,
                element=header,
                message=f"{display}: table column {header!r} removed",
            )
        )
    for row_index in range(shared_rows, len(current_rows)):
        label = current_rows[row_index][0] if current_rows[row_index] else str(row_index)
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.TABLE_VALUE_CHANGED,
                slide=display,
                element=label,
                message=f"{display}: table row {label!r} added",
            )
        )
    for row_index in range(shared_rows, len(baseline_rows)):
        label = baseline_rows[row_index][0] if baseline_rows[row_index] else str(row_index)
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.TABLE_VALUE_CHANGED,
                slide=display,
                element=label,
                message=f"{display}: table row {label!r} removed",
            )
        )
    return findings


def _chart_label(chart: ChartContent) -> str:
    return chart.title or chart.name or f"chart {chart.source_index + 1}"


def _series_label(series: PptChartSeries) -> str:
    return series.name or f"series {series.source_index + 1}"


def _labels_signature(labels: ChartDataLabels | None) -> tuple[object, ...] | None:
    if labels is None:
        return None
    return (
        labels.position,
        bool(labels.show_value),
        bool(labels.show_category_name),
        bool(labels.show_series_name),
        bool(labels.show_percent),
        bool(labels.show_legend_key),
        bool(labels.show_bubble_size),
        bool(labels.show_leader_lines),
        labels.number_format,
        labels.separator,
    )


def _legend_signature(legend: ChartLegend | None) -> tuple[object, ...] | None:
    if legend is None:
        return None
    return legend.position, bool(legend.overlay), bool(legend.deleted)


def _axis_signature(axis: ChartAxis) -> tuple[object, ...]:
    return (
        axis.axis_type,
        axis.position,
        axis.minimum,
        axis.maximum,
        axis.major_unit,
        axis.minor_unit,
        axis.log_base,
        axis.orientation,
        axis.crosses,
        axis.crosses_at,
        axis.number_format,
        axis.display_unit,
        axis.major_tick_mark,
        axis.minor_tick_mark,
        axis.title,
        bool(axis.deleted),
    )


def _plot_axes(
    chart: ChartContent,
    plot: PptChartPlot,
) -> tuple[list[ChartAxis], tuple[str, ...]]:
    by_id = {axis.axis_id: axis for axis in chart.axes}
    axes = [by_id[axis_id] for axis_id in plot.axis_ids if axis_id in by_id]
    missing = tuple(axis_id for axis_id in plot.axis_ids if axis_id not in by_id)
    return axes, missing


def _material_number_change(
    baseline: int,
    current: int,
    *,
    relative: bool = False,
) -> bool:
    tolerance = _GEOMETRY_OFFSET_TOLERANCE
    if relative:
        tolerance = max(tolerance, int(abs(baseline) * _GEOMETRY_SIZE_RATIO))
    return abs(current - baseline) > tolerance


def _material_geometry_change(
    baseline: ChartContent,
    current: ChartContent,
) -> bool:
    return any(
        (
            _material_number_change(
                getattr(baseline, attribute),
                getattr(current, attribute),
                relative=attribute in {"width", "height"},
            )
        )
        for attribute in ("left", "top", "width", "height")
    )


def _window_override(
    profile: PptProfile,
    slide: SlideContent,
    chart: ChartContent,
    series: PptChartSeries,
) -> str | None:
    overrides = {key.casefold(): value for key, value in profile.chart_windows.items()}
    chart_key = (chart.title or f"chart[{chart.source_index}]").casefold()
    series_key = (series.name or f"series[{series.source_index}]").casefold()
    for key in (
        f"{slide.display_name}/{chart_key}/{series_key}",
        f"{slide.display_name}/{chart_key}",
        slide.display_name,
    ):
        normalized = key.casefold()
        if normalized in overrides:
            return overrides[normalized]
    return None


def _classify_window(
    baseline: list[str],
    current: list[str],
    override: str | None,
) -> str:
    if override in {"rolling", "full"}:
        return override
    if baseline and len(baseline) == len(current) and baseline != current:
        for shift in range(1, len(baseline)):
            if current[:-shift] == baseline[shift:]:
                return "rolling"
    return "full"


def _category_keys(categories: list[str]) -> list[tuple[str, int]]:
    counts: Counter[str] = Counter()
    keys: list[tuple[str, int]] = []
    for category in categories:
        keys.append((category, counts[category]))
        counts[category] += 1
    return keys


def _series_data_findings(
    slide: SlideContent,
    chart: ChartContent,
    baseline: PptChartSeries,
    current: PptChartSeries,
    profile: PptProfile,
) -> list[Finding]:
    findings: list[Finding] = []
    display = slide.display_name
    series_name = _series_label(current)
    override = _window_override(profile, slide, chart, current)
    window = _classify_window(baseline.categories, current.categories, override)
    baseline_latest = _latest_periods(baseline.categories)
    baseline_keys = _category_keys(baseline.categories)
    current_keys = _category_keys(current.categories)
    baseline_index = {key: index for index, key in enumerate(baseline_keys)}
    current_index = {key: index for index, key in enumerate(current_keys)}
    if window == "rolling" and baseline.categories != current.categories:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.CHART_VALUE_CHANGED,
                expected_reason=FindingExpectedReason.ROLLING_WINDOW,
                slide=display,
                element="window",
                baseline_value=(
                    f"{baseline.categories[0]}..{baseline.categories[-1]}"
                    if baseline.categories
                    else ""
                ),
                current_value=(
                    f"{current.categories[0]}..{current.categories[-1]}"
                    if current.categories
                    else ""
                ),
                message=f"{display}: {series_name!r} rolling chart window advanced",
            )
        )
    for key in baseline_keys:
        if key in current_index or window == "rolling":
            continue
        category = key[0]
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.CHART_VALUE_CHANGED,
                slide=display,
                element=category,
                baseline_value=category,
                message=f"{display}: {series_name!r} chart category {category!r} removed",
            )
        )
    for key in current_keys:
        if key in baseline_index:
            continue
        category = key[0]
        period = parse_period(category)
        previous = baseline_latest.get(period.kind) if period is not None else None
        expected = bool(
            period is not None
            and (previous is None or is_period_after(period, previous))
        )
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.CHART_VALUE_CHANGED,
                expected_reason=(
                    FindingExpectedReason.ROLLING_WINDOW
                    if window == "rolling"
                    else (
                        FindingExpectedReason.CADENCE_EXTENSION
                        if expected
                        else None
                    )
                ),
                slide=display,
                element=category,
                current_value=category,
                message=(
                    f"{display}: {series_name!r} chart category {category!r} "
                    + ("appended (new cycle)" if expected else "added")
                ),
            )
        )
        current_position = current_index[key]
        current_value = (
            current.values[current_position]
            if current_position < len(current.values)
            else None
        )
        if current_value is None and not ppt_blank_allowed(
            profile,
            slide=display,
            scope="chart",
            element=_chart_label(chart),
            series=series_name,
            period_label=category,
        ):
            findings.append(
                Finding(
                    artifact="ppt",
                    finding_class=FindingClass.PPT_CHART_VALUE_MISSING,
                    slide=display,
                    element=category,
                    message=(
                        f"{display}: {series_name!r} has no value for "
                        f"{category!r}"
                    ),
                )
            )
    for key in baseline_keys:
        if key not in current_index:
            continue
        baseline_position = baseline_index[key]
        current_position = current_index[key]
        baseline_value = (
            baseline.values[baseline_position]
            if baseline_position < len(baseline.values)
            else None
        )
        current_value = (
            current.values[current_position]
            if current_position < len(current.values)
            else None
        )
        if baseline_value == current_value:
            continue
        if current_value is None and ppt_blank_allowed(
            profile,
            slide=display,
            scope="chart",
            element=_chart_label(chart),
            series=series_name,
            period_label=key[0],
        ):
            continue
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.CHART_VALUE_CHANGED,
                slide=display,
                element=key[0],
                baseline_value=str(baseline_value),
                current_value=str(current_value),
                message=(
                    f"{display}: {series_name!r} chart value for {key[0]!r} changed"
                ),
            )
        )
    return findings


def _series_findings(
    slide: SlideContent,
    chart: ChartContent,
    baseline: PptChartSeries,
    current: PptChartSeries,
    profile: PptProfile,
) -> list[Finding]:
    findings: list[Finding] = []
    display = slide.display_name
    if baseline.name != current.name:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_SERIES_CHANGED,
                slide=display,
                element=_chart_label(chart),
                baseline_value=baseline.name,
                current_value=current.name,
                message=(
                    f"{display}: chart series {baseline.name!r} renamed to "
                    f"{current.name!r}"
                ),
            )
        )
    for attribute, label in (
        ("name_ref", "name source"),
        ("categories_ref", "category source"),
        ("values_ref", "value source"),
        ("bubble_size_ref", "bubble-size source"),
    ):
        baseline_ref = getattr(baseline, attribute)
        current_ref = getattr(current, attribute)
        if baseline_ref == current_ref:
            continue
        if (
            baseline_ref is not None
            and current_ref is not None
            and is_pure_range_extension(baseline_ref, current_ref)
        ):
            continue
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_SERIES_CHANGED,
                slide=display,
                element=_series_label(current),
                baseline_value=baseline_ref,
                current_value=current_ref,
                message=(
                    f"{display}: {_series_label(current)!r} {label} changed"
                ),
            )
        )
    if _labels_signature(baseline.data_labels) != _labels_signature(
        current.data_labels
    ):
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_LABELS_CHANGED,
                slide=display,
                element=_series_label(current),
                message=f"{display}: {_series_label(current)!r} data labels changed",
            )
        )
    findings.extend(_series_data_findings(slide, chart, baseline, current, profile))
    return findings


def _plot_findings(
    slide: SlideContent,
    baseline_chart: ChartContent,
    current_chart: ChartContent,
    baseline: PptChartPlot,
    current: PptChartPlot,
    profile: PptProfile,
) -> list[Finding]:
    findings: list[Finding] = []
    display = slide.display_name
    baseline_settings = (
        baseline.chart_type,
        baseline.grouping,
        baseline.direction,
        baseline.style,
    )
    current_settings = (
        current.chart_type,
        current.grouping,
        current.direction,
        current.style,
    )
    if baseline_settings != current_settings:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_PLOT_CHANGED,
                slide=display,
                element=_chart_label(current_chart),
                baseline_value=baseline.chart_type,
                current_value=current.chart_type,
                message=(
                    f"{display}: chart plot changed from {baseline.chart_type} "
                    f"to {current.chart_type}"
                ),
            )
        )
    baseline_axes, baseline_missing_axes = _plot_axes(baseline_chart, baseline)
    current_axes, current_missing_axes = _plot_axes(current_chart, current)
    if (
        baseline.axis_group != current.axis_group
        or baseline_missing_axes
        or current_missing_axes
        or [_axis_signature(axis) for axis in baseline_axes]
        != [_axis_signature(axis) for axis in current_axes]
    ):
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_AXIS_CHANGED,
                slide=display,
                element=_chart_label(current_chart),
                message=(
                    f"{display}: chart plot axis group/settings changed"
                    + (
                        f"; uncaptured axis IDs: "
                        f"{sorted(set(baseline_missing_axes + current_missing_axes))}"
                        if baseline_missing_axes or current_missing_axes
                        else ""
                    )
                ),
            )
        )
    if _labels_signature(baseline.data_labels) != _labels_signature(
        current.data_labels
    ):
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_LABELS_CHANGED,
                slide=display,
                element=_chart_label(current_chart),
                message=f"{display}: chart plot data labels changed",
            )
        )
    matching = match_series(baseline, current)
    for series in matching.removed:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_SERIES_CHANGED,
                slide=display,
                element=_series_label(series),
                message=f"{display}: chart series {_series_label(series)!r} removed",
            )
        )
    for series in matching.added:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_SERIES_CHANGED,
                slide=display,
                element=_series_label(series),
                message=f"{display}: chart series {_series_label(series)!r} added",
            )
        )
    for baseline_series, current_series in matching.pairs:
        findings.extend(
            _series_findings(
                slide,
                current_chart,
                baseline_series,
                current_series,
                profile,
            )
        )
    for _baseline_series, current_series in matching.reordered:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_SERIES_CHANGED,
                slide=display,
                element=_series_label(current_series),
                message=f"{display}: chart series {_series_label(current_series)!r} reordered",
            )
        )
    return findings


def _chart_findings(
    slide: SlideContent,
    baseline: ChartContent,
    current: ChartContent,
    profile: PptProfile,
) -> list[Finding]:
    findings: list[Finding] = []
    display = slide.display_name
    if baseline.title != current.title:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_STRUCTURE_CHANGED,
                slide=display,
                element=_chart_label(current),
                baseline_value=baseline.title,
                current_value=current.title,
                message=f"{display}: chart title changed",
            )
        )
    if baseline.display_blanks_as != current.display_blanks_as:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_STRUCTURE_CHANGED,
                slide=display,
                element=_chart_label(current),
                baseline_value=baseline.display_blanks_as,
                current_value=current.display_blanks_as,
                message=f"{display}: chart blank-display setting changed",
            )
        )
    if _legend_signature(baseline.legend) != _legend_signature(current.legend):
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_LEGEND_CHANGED,
                slide=display,
                element=_chart_label(current),
                message=f"{display}: chart legend settings changed",
            )
        )
    if _material_geometry_change(baseline, current):
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_SHAPE_GEOMETRY_CHANGED,
                slide=display,
                element=_chart_label(current),
                baseline_value=f"{baseline.left}:{baseline.top}:{baseline.width}:{baseline.height}",
                current_value=f"{current.left}:{current.top}:{current.width}:{current.height}",
                message=f"{display}: chart moved or resized materially",
            )
        )
    matching = match_plots(baseline, current)
    for plot in matching.removed:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_PLOT_CHANGED,
                slide=display,
                element=_chart_label(current),
                message=f"{display}: {plot.chart_type} plot removed",
            )
        )
    for plot in matching.added:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_PLOT_CHANGED,
                slide=display,
                element=_chart_label(current),
                message=f"{display}: {plot.chart_type} plot added",
            )
        )
    for baseline_plot, current_plot in matching.pairs:
        findings.extend(
            _plot_findings(
                slide,
                baseline,
                current,
                baseline_plot,
                current_plot,
                profile,
            )
        )
    for _baseline_plot, current_plot in matching.reordered:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_PLOT_CHANGED,
                slide=display,
                element=_chart_label(current),
                message=f"{display}: {current_plot.chart_type} plot reordered",
            )
        )
    return findings


def _shape_findings(
    findings: list[Finding],
    *,
    baseline_shape_id: int | None = None,
    current_shape_id: int | None = None,
) -> list[Finding]:
    baseline_shape_id = (
        baseline_shape_id
        if baseline_shape_id is not None and baseline_shape_id > 0
        else None
    )
    current_shape_id = (
        current_shape_id
        if current_shape_id is not None and current_shape_id > 0
        else None
    )
    for finding in findings:
        finding.baseline_focus_shape_id = baseline_shape_id
        finding.focus_shape_id = current_shape_id
    return findings


def diff_slide_elements(
    baseline: SlideContent,
    current: SlideContent,
    profile: PptProfile,
) -> list[Finding]:
    """Diff every table and chart in one matched slide pair."""
    findings: list[Finding] = []
    display = current.display_name
    matching = match_slide_elements(baseline, current)
    for table in matching.removed_tables:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_TABLE_STRUCTURE_CHANGED,
                slide=display,
                element=_table_label(table),
                baseline_focus_shape_id=table.shape_id or None,
                message=f"{display}: table {_table_label(table)!r} removed",
            )
        )
    for table in matching.added_tables:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_TABLE_STRUCTURE_CHANGED,
                slide=display,
                element=_table_label(table),
                focus_shape_id=table.shape_id or None,
                message=f"{display}: table {_table_label(table)!r} added",
            )
        )
    for baseline_table, current_table in matching.table_pairs:
        findings.extend(
            _shape_findings(
                _table_findings(
                    current,
                    baseline_table,
                    current_table,
                    profile,
                ),
                baseline_shape_id=baseline_table.shape_id,
                current_shape_id=current_table.shape_id,
            )
        )
    for baseline_table, current_table in matching.reordered_tables:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_TABLE_STRUCTURE_CHANGED,
                slide=display,
                element=_table_label(current_table),
                focus_shape_id=current_table.shape_id or None,
                baseline_focus_shape_id=baseline_table.shape_id or None,
                message=f"{display}: table {_table_label(current_table)!r} reordered",
            )
        )

    for chart in matching.removed_charts:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_STRUCTURE_CHANGED,
                slide=display,
                element=_chart_label(chart),
                baseline_focus_shape_id=chart.shape_id or None,
                message=f"{display}: chart {_chart_label(chart)!r} removed",
            )
        )
    for chart in matching.added_charts:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_STRUCTURE_CHANGED,
                slide=display,
                element=_chart_label(chart),
                focus_shape_id=chart.shape_id or None,
                message=f"{display}: chart {_chart_label(chart)!r} added",
            )
        )
    for baseline_chart, current_chart in matching.chart_pairs:
        findings.extend(
            _shape_findings(
                _chart_findings(current, baseline_chart, current_chart, profile),
                baseline_shape_id=baseline_chart.shape_id,
                current_shape_id=current_chart.shape_id,
            )
        )
    for baseline_chart, current_chart in matching.reordered_charts:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_CHART_STRUCTURE_CHANGED,
                slide=display,
                element=_chart_label(current_chart),
                focus_shape_id=current_chart.shape_id or None,
                baseline_focus_shape_id=baseline_chart.shape_id or None,
                message=f"{display}: chart {_chart_label(current_chart)!r} reordered",
            )
        )
    for finding in findings:
        finding.slide_index = current.index + 1
        finding.baseline_slide_index = baseline.index + 1
    return findings
