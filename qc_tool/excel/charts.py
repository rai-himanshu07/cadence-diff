"""Semantic Excel chart matching, QC, and source-cell impact annotation."""

from __future__ import annotations

from collections.abc import Callable
from typing import TypeVar

from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

from qc_tool.config.profile import DeliverableProfile
from qc_tool.coverage import CoverageState
from qc_tool.excel.dependency import DependencyGraph, Node, dependent_nodes_of
from qc_tool.excel.references import (
    ReferenceStatus,
    is_pure_range_extension,
    resolve_reference,
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
    WorkbookSnapshot,
)

_T = TypeVar("_T")
_OFFSET_TOLERANCE = 50_000
_SIZE_TOLERANCE_RATIO = 0.02


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


def _chart_title_key(chart: ChartDescriptor) -> tuple[str, str] | None:
    if not chart.title:
        return None
    return chart.sheet, chart.title.casefold()


def _series_name(series: ChartSeries) -> str | None:
    return series.name_text or series.name_ref


def _series_name_key(series: ChartSeries) -> str | None:
    name = _series_name(series)
    return name.casefold() if name else None


def _plot_signature(plot: ChartPlot) -> tuple[str, ...] | None:
    names = tuple(sorted(name for item in plot.series if (name := _series_name_key(item))))
    return names or None


def _chart_signature(chart: ChartDescriptor) -> tuple[str, tuple[tuple[str, ...], ...]] | None:
    signatures = tuple(
        signature for plot in chart.plots if (signature := _plot_signature(plot))
    )
    return (chart.sheet, signatures) if signatures else None


def _chart_geometry_key(chart: ChartDescriptor) -> tuple[object, ...] | None:
    anchor = chart.anchor
    if anchor is None:
        return None
    return (
        chart.sheet,
        anchor.anchor_type,
        anchor.from_col,
        anchor.from_row,
        anchor.to_col,
        anchor.to_row,
    )


def _match_charts(
    baseline: list[ChartDescriptor], current: list[ChartDescriptor]
) -> tuple[
    list[tuple[ChartDescriptor, ChartDescriptor]],
    list[ChartDescriptor],
    list[ChartDescriptor],
]:
    unmatched_baseline = list(baseline)
    unmatched_current = list(current)
    pairs = _consume_unique_matches(
        unmatched_baseline, unmatched_current, _chart_title_key
    )
    pairs.extend(
        _consume_unique_matches(
            unmatched_baseline,
            unmatched_current,
            _chart_signature,
        )
    )
    pairs.extend(
        _consume_unique_matches(
            unmatched_baseline,
            unmatched_current,
            _chart_geometry_key,
        )
    )
    pairs.extend(
        _consume_ordered_matches(
            unmatched_baseline,
            unmatched_current,
            lambda baseline_chart, current_chart: (
                baseline_chart.sheet == current_chart.sheet
                and baseline_chart.title is not None
                and current_chart.title is not None
                and baseline_chart.title.casefold() == current_chart.title.casefold()
            ),
        )
    )
    pairs.extend(
        _consume_ordered_matches(
            unmatched_baseline,
            unmatched_current,
            lambda baseline_chart, current_chart: (
                _chart_signature(baseline_chart) is not None
                and _chart_signature(baseline_chart) == _chart_signature(current_chart)
            ),
        )
    )
    pairs.extend(
        _consume_unique_matches(
            unmatched_baseline,
            unmatched_current,
            lambda chart: (chart.sheet, chart.source_index),
        )
    )
    return pairs, unmatched_baseline, unmatched_current


def _legacy_plots(chart: ChartDescriptor) -> list[ChartPlot]:
    return chart.plots or [
        ChartPlot(index=0, chart_type=chart.chart_type, series=chart.series)
    ]


def _match_plots(
    baseline: list[ChartPlot], current: list[ChartPlot]
) -> tuple[list[tuple[ChartPlot, ChartPlot]], list[ChartPlot], list[ChartPlot]]:
    unmatched_baseline = list(baseline)
    unmatched_current = list(current)
    pairs = _consume_unique_matches(
        unmatched_baseline, unmatched_current, _plot_signature
    )
    pairs.extend(
        _consume_unique_matches(
            unmatched_baseline,
            unmatched_current,
            lambda plot: plot.index,
        )
    )
    return pairs, unmatched_baseline, unmatched_current


def _series_source_key(series: ChartSeries) -> tuple[str | None, str | None] | None:
    if series.values_ref is None and series.categories_ref is None:
        return None
    return series.values_ref, series.categories_ref


def _series_reorder_identity(series: ChartSeries) -> str | None:
    name = _series_name_key(series)
    if name is not None:
        return f"name:{name}"
    source = _series_source_key(series)
    return f"source:{source!r}" if source is not None else None


def _series_position(series: ChartSeries) -> int:
    return series.order if series.order is not None else series.source_index


def _match_series(
    baseline: list[ChartSeries], current: list[ChartSeries]
) -> tuple[
    list[tuple[ChartSeries, ChartSeries]], list[ChartSeries], list[ChartSeries]
]:
    unmatched_baseline = list(baseline)
    unmatched_current = list(current)
    pairs = _consume_unique_matches(
        unmatched_baseline, unmatched_current, _series_name_key
    )
    pairs.extend(
        _consume_unique_matches(
            unmatched_baseline,
            unmatched_current,
            _series_source_key,
        )
    )
    pairs.extend(
        _consume_unique_matches(
            unmatched_baseline,
            unmatched_current,
            lambda series: series.source_index,
        )
    )
    return pairs, unmatched_baseline, unmatched_current


def _chart_label(chart: ChartDescriptor) -> str:
    return chart.title or f"{chart.chart_type} chart {chart.source_index + 1}"


def _series_label(series: ChartSeries) -> str:
    return _series_name(series) or f"series {series.source_index + 1}"


def _a1_bounds(reference: str) -> tuple[str, int, int, int, int] | None:
    sheet_part, separator, cell_range = reference.rpartition("!")
    if not separator or "[" in reference:
        return None
    try:
        min_col, min_row, max_col, max_row = range_boundaries(
            cell_range.replace("$", "")
        )
    except ValueError:
        return None
    if min_col is None or min_row is None or max_col is None or max_row is None:
        return None
    sheet = sheet_part.strip("'").replace("''", "'")
    return sheet, min_col, min_row, max_col, max_row


def _is_rolling_shift(baseline: str, current: str) -> bool:
    baseline_bounds = _a1_bounds(baseline)
    current_bounds = _a1_bounds(current)
    if baseline_bounds is None or current_bounds is None:
        return False
    base_sheet, base_min_col, base_min_row, base_max_col, base_max_row = baseline_bounds
    curr_sheet, curr_min_col, curr_min_row, curr_max_col, curr_max_row = current_bounds
    row_shift = curr_min_row - base_min_row
    column_shift = curr_min_col - base_min_col
    return (
        base_sheet == curr_sheet
        and base_max_row - base_min_row == curr_max_row - curr_min_row
        and base_max_col - base_min_col == curr_max_col - curr_min_col
        and curr_max_row - base_max_row == row_shift
        and curr_max_col - base_max_col == column_shift
        and (row_shift > 0 or column_shift > 0)
        and row_shift >= 0
        and column_shift >= 0
    )


def _window_override(
    profile: DeliverableProfile | None, chart: ChartDescriptor
) -> str | None:
    if profile is None:
        return None
    sheet_profile = profile.sheet_profile(chart.sheet)
    if sheet_profile is None:
        return None
    for key in (chart.title, f"chart[{chart.source_index}]"):
        if key and key in sheet_profile.chart_windows:
            return sheet_profile.chart_windows[key]
    return None


def _expected_source_change(
    baseline: str | None,
    current: str | None,
    *,
    window_override: str | None,
) -> bool:
    if baseline is None or current is None:
        return False
    if is_pure_range_extension(baseline, current):
        return True
    if window_override == "full":
        return False
    return _is_rolling_shift(baseline, current)


def _data_labels_signature(value: ChartDataLabels | None) -> tuple[object, ...] | None:
    if value is None:
        return None
    return (
        value.position,
        bool(value.show_value),
        bool(value.show_category_name),
        bool(value.show_series_name),
        bool(value.show_percent),
        bool(value.show_legend_key),
        bool(value.show_bubble_size),
        bool(value.show_leader_lines),
        value.number_format,
        value.separator,
    )


def _legend_signature(value: ChartLegend | None) -> tuple[object, ...] | None:
    if value is None:
        return None
    return value.position, bool(value.overlay), bool(value.deleted)


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


def _plot_axes(chart: ChartDescriptor, plot: ChartPlot) -> list[ChartAxis]:
    by_id = {axis.axis_id: axis for axis in chart.axes}
    return [by_id[axis_id] for axis_id in plot.axis_ids if axis_id in by_id]


def _material_numeric_change(
    baseline: int | None,
    current: int | None,
    *,
    relative: bool = False,
) -> bool:
    if baseline is None or current is None:
        return baseline != current
    tolerance = _OFFSET_TOLERANCE
    if relative:
        tolerance = max(tolerance, int(abs(baseline) * _SIZE_TOLERANCE_RATIO))
    return abs(current - baseline) > tolerance


def _material_geometry_change(
    baseline: ChartAnchor | None, current: ChartAnchor | None
) -> bool:
    if baseline is None or current is None:
        return baseline != current
    if baseline.anchor_type != current.anchor_type:
        return True
    if (
        baseline.from_col,
        baseline.from_row,
        baseline.to_col,
        baseline.to_row,
    ) != (
        current.from_col,
        current.from_row,
        current.to_col,
        current.to_row,
    ):
        return True
    return any(
        (
            _material_numeric_change(
                getattr(baseline, attribute),
                getattr(current, attribute),
                relative=attribute in {"width", "height"},
            )
        )
        for attribute in (
            "from_col_offset",
            "from_row_offset",
            "to_col_offset",
            "to_row_offset",
            "x",
            "y",
            "width",
            "height",
        )
    )


def _finding(
    finding_class: FindingClass,
    chart: ChartDescriptor,
    message: str,
    *,
    baseline_value: str | None = None,
    current_value: str | None = None,
    expected_growth: bool = False,
) -> Finding:
    return Finding(
        artifact="excel",
        finding_class=finding_class,
        expected_growth=expected_growth,
        sheet=chart.sheet,
        element=_chart_label(chart),
        baseline_value=baseline_value,
        current_value=current_value,
        message=message,
    )


def _series_findings(
    baseline: ChartSeries,
    current: ChartSeries,
    chart: ChartDescriptor,
    plot: ChartPlot,
    *,
    window_override: str | None,
) -> list[Finding]:
    findings: list[Finding] = []
    baseline_name = _series_name(baseline)
    current_name = _series_name(current)
    if baseline_name != current_name:
        findings.append(
            _finding(
                FindingClass.CHART_SERIES_CHANGED,
                chart,
                f"chart {_chart_label(chart)!r} plot {plot.index} series name changed",
                baseline_value=baseline_name,
                current_value=current_name,
            )
        )
    for attribute, label in (
        ("categories_ref", "categories"),
        ("values_ref", "values"),
        ("bubble_size_ref", "bubble size"),
    ):
        baseline_ref = getattr(baseline, attribute)
        current_ref = getattr(current, attribute)
        if baseline_ref == current_ref:
            continue
        expected = _expected_source_change(
            baseline_ref,
            current_ref,
            window_override=window_override,
        )
        findings.append(
            _finding(
                FindingClass.CHART_SERIES_CHANGED,
                chart,
                (
                    f"chart {_chart_label(chart)!r} plot {plot.index} "
                    f"series {_series_label(current)!r} {label} "
                    + ("advanced with expected cadence data" if expected else "changed")
                ),
                baseline_value=baseline_ref,
                current_value=current_ref,
                expected_growth=expected,
            )
        )
    if _data_labels_signature(baseline.data_labels) != _data_labels_signature(
        current.data_labels
    ):
        findings.append(
            _finding(
                FindingClass.CHART_LABELS_CHANGED,
                chart,
                (
                    f"chart {_chart_label(chart)!r} series "
                    f"{_series_label(current)!r} data labels changed"
                ),
            )
        )
    return findings


def _plot_findings(
    baseline: ChartPlot,
    current: ChartPlot,
    baseline_chart: ChartDescriptor,
    current_chart: ChartDescriptor,
    *,
    window_override: str | None,
) -> list[Finding]:
    findings: list[Finding] = []
    plot_settings = (
        baseline.chart_type,
        baseline.grouping,
        baseline.direction,
        baseline.style,
    )
    current_plot_settings = (
        current.chart_type,
        current.grouping,
        current.direction,
        current.style,
    )
    if plot_settings != current_plot_settings:
        setting_names = ("type", "grouping", "direction", "style")
        changes = ", ".join(
            f"{name}: {before!r} -> {after!r}"
            for name, before, after in zip(
                setting_names,
                plot_settings,
                current_plot_settings,
                strict=True,
            )
            if before != after
        )
        findings.append(
            _finding(
                FindingClass.CHART_PLOT_CHANGED,
                current_chart,
                (
                    f"chart {_chart_label(current_chart)!r} plot {current.index} "
                    f"settings changed ({changes})"
                ),
                baseline_value=baseline.chart_type,
                current_value=current.chart_type,
            )
        )
    if baseline.axis_group != current.axis_group:
        findings.append(
            _finding(
                FindingClass.CHART_AXIS_CHANGED,
                current_chart,
                f"chart {_chart_label(current_chart)!r} plot {current.index} axis group changed",
                baseline_value=baseline.axis_group,
                current_value=current.axis_group,
            )
        )
    baseline_axes = _plot_axes(baseline_chart, baseline)
    current_axes = _plot_axes(current_chart, current)
    if [_axis_signature(axis) for axis in baseline_axes] != [
        _axis_signature(axis) for axis in current_axes
    ]:
        findings.append(
            _finding(
                FindingClass.CHART_AXIS_CHANGED,
                current_chart,
                f"chart {_chart_label(current_chart)!r} plot {current.index} axis settings changed",
            )
        )
    if _data_labels_signature(baseline.data_labels) != _data_labels_signature(
        current.data_labels
    ):
        findings.append(
            _finding(
                FindingClass.CHART_LABELS_CHANGED,
                current_chart,
                f"chart {_chart_label(current_chart)!r} plot {current.index} data labels changed",
            )
        )

    series_pairs, removed, added = _match_series(baseline.series, current.series)
    for series in removed:
        findings.append(
            _finding(
                FindingClass.CHART_SERIES_CHANGED,
                current_chart,
                f"chart {_chart_label(current_chart)!r} series {_series_label(series)!r} removed",
            )
        )
    for series in added:
        findings.append(
            _finding(
                FindingClass.CHART_SERIES_CHANGED,
                current_chart,
                f"chart {_chart_label(current_chart)!r} series {_series_label(series)!r} added",
            )
        )
    for baseline_series, current_series in series_pairs:
        findings.extend(
            _series_findings(
                baseline_series,
                current_series,
                current_chart,
                current,
                window_override=window_override,
            )
        )
    if not removed and not added:
        baseline_pairs = [
            (identity, series)
            for series in baseline.series
            if (identity := _series_reorder_identity(series)) is not None
        ]
        current_pairs = [
            (identity, series)
            for series in current.series
            if (identity := _series_reorder_identity(series)) is not None
        ]
        baseline_by_identity = dict(baseline_pairs)
        current_by_identity = dict(current_pairs)
        if (
            len(baseline_pairs) == len(baseline.series)
            and len(current_pairs) == len(current.series)
            and baseline_by_identity.keys() == current_by_identity.keys()
            and len(baseline_by_identity) == len(baseline.series)
            and len(current_by_identity) == len(current.series)
        ):
            baseline_order = sorted(
                baseline_by_identity,
                key=lambda identity: _series_position(baseline_by_identity[identity]),
            )
            current_order = sorted(
                current_by_identity,
                key=lambda identity: _series_position(current_by_identity[identity]),
            )
            if baseline_order != current_order:
                findings.append(
                    _finding(
                        FindingClass.CHART_SERIES_CHANGED,
                        current_chart,
                        (
                            f"chart {_chart_label(current_chart)!r} plot "
                            f"{current.index} series reordered"
                        ),
                    )
                )
    return findings


def _chart_pair_findings(
    baseline: ChartDescriptor,
    current: ChartDescriptor,
    *,
    profile: DeliverableProfile | None,
) -> list[Finding]:
    findings: list[Finding] = []
    if baseline.title != current.title:
        findings.append(
            _finding(
                FindingClass.CHART_STRUCTURE_CHANGED,
                current,
                f"chart {_chart_label(current)!r} title changed",
                baseline_value=baseline.title,
                current_value=current.title,
            )
        )
    if baseline.display_blanks_as != current.display_blanks_as:
        findings.append(
            _finding(
                FindingClass.CHART_STRUCTURE_CHANGED,
                current,
                f"chart {_chart_label(current)!r} blank-display setting changed",
                baseline_value=baseline.display_blanks_as,
                current_value=current.display_blanks_as,
            )
        )
    if _legend_signature(baseline.legend) != _legend_signature(current.legend):
        findings.append(
            _finding(
                FindingClass.CHART_LEGEND_CHANGED,
                current,
                f"chart {_chart_label(current)!r} legend settings changed",
            )
        )
    if _material_geometry_change(baseline.anchor, current.anchor):
        findings.append(
            _finding(
                FindingClass.CHART_GEOMETRY_CHANGED,
                current,
                f"chart {_chart_label(current)!r} moved or resized materially",
                baseline_value=(
                    baseline.anchor.signature if baseline.anchor is not None else None
                ),
                current_value=(
                    current.anchor.signature if current.anchor is not None else None
                ),
            )
        )

    plot_pairs, removed, added = _match_plots(
        _legacy_plots(baseline), _legacy_plots(current)
    )
    for plot in removed:
        findings.append(
            _finding(
                FindingClass.CHART_PLOT_CHANGED,
                current,
                f"chart {_chart_label(current)!r} {plot.chart_type} plot removed",
            )
        )
    for plot in added:
        findings.append(
            _finding(
                FindingClass.CHART_PLOT_CHANGED,
                current,
                f"chart {_chart_label(current)!r} {plot.chart_type} plot added",
            )
        )
    override = _window_override(profile, current)
    for baseline_plot, current_plot in plot_pairs:
        findings.extend(
            _plot_findings(
                baseline_plot,
                current_plot,
                baseline,
                current,
                window_override=override,
            )
        )
    return findings


def _reference_findings(workbook: WorkbookSnapshot) -> list[Finding]:
    findings: list[Finding] = []
    for chart in workbook.charts:
        for series in chart.series:
            sizes: dict[str, int] = {}
            for attribute, label in (
                ("categories_ref", "categories"),
                ("values_ref", "values"),
                ("bubble_size_ref", "bubble size"),
            ):
                target = getattr(series, attribute)
                if target is None:
                    continue
                resolution = resolve_reference(
                    workbook,
                    target,
                    host_sheet=chart.sheet,
                )
                if resolution.status is ReferenceStatus.INVALID:
                    findings.append(
                        _finding(
                            FindingClass.CHART_REFERENCE_INVALID,
                            chart,
                            (
                                f"chart {_chart_label(chart)!r} series "
                                f"{_series_label(series)!r} has an invalid {label} reference"
                            ),
                            current_value=target,
                        )
                    )
                elif (
                    resolution.status is ReferenceStatus.RESOLVED
                    and attribute != "bubble_size_ref"
                ):
                    sizes[label] = resolution.size
            if len(sizes) == 2 and sizes["categories"] != sizes["values"]:
                findings.append(
                    _finding(
                        FindingClass.CHART_LENGTH_MISMATCH,
                        chart,
                        (
                            f"chart {_chart_label(chart)!r} series "
                            f"{_series_label(series)!r} has different category and value lengths"
                        ),
                        baseline_value=str(sizes["categories"]),
                        current_value=str(sizes["values"]),
                    )
                )
    return findings


def diff_charts(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    profile: DeliverableProfile | None = None,
) -> list[Finding]:
    """Compare complete Excel chart graphs with explicit unmatched elements."""
    if not baseline.charts_available or not current.charts_available:
        return []
    pairs, removed, added = _match_charts(baseline.charts, current.charts)
    findings = [
        _finding(
            FindingClass.CHART_STRUCTURE_CHANGED,
            chart,
            f"chart {_chart_label(chart)!r} removed",
        )
        for chart in removed
    ]
    findings.extend(
        _finding(
            FindingClass.CHART_STRUCTURE_CHANGED,
            chart,
            f"chart {_chart_label(chart)!r} added",
        )
        for chart in added
    )
    for baseline_chart, current_chart in pairs:
        findings.extend(
            _chart_pair_findings(
                baseline_chart,
                current_chart,
                profile=profile,
            )
        )
    findings.extend(_reference_findings(current))
    return findings


def chart_reference_coverage(workbook: WorkbookSnapshot) -> tuple[CoverageState, str]:
    """Coverage for chart source formulas, distinct from invalid-source findings."""
    if not workbook.charts_available:
        return CoverageState.DEGRADED, workbook.chart_detail
    unsupported = 0
    missing = 0
    checked = 0
    for chart in workbook.charts:
        for series in chart.series:
            series_targets = (
                series.categories_ref,
                series.values_ref,
                series.bubble_size_ref,
            )
            if all(target is None for target in series_targets):
                missing += 1
            for target in (series.categories_ref, series.values_ref, series.bubble_size_ref):
                if target is None:
                    continue
                resolution = resolve_reference(
                    workbook,
                    target,
                    host_sheet=chart.sheet,
                )
                unsupported += resolution.status is ReferenceStatus.UNSUPPORTED
                checked += resolution.status is not ReferenceStatus.UNSUPPORTED
    if unsupported or missing:
        details = []
        if unsupported:
            details.append(f"{unsupported} unsupported chart source references")
        if missing:
            details.append(f"{missing} chart series have no resolvable source reference")
        return (
            CoverageState.DEGRADED,
            "; ".join(details) + " and were not fully validated",
        )
    if workbook.charts and checked == 0:
        return CoverageState.DEGRADED, "No chart source references were available to validate"
    return CoverageState.CHECKED, "Complete chart sources were validated"


def annotate_chart_impacts(
    findings: list[Finding],
    workbook: WorkbookSnapshot,
    dependency_graph: DependencyGraph | None = None,
) -> None:
    """Attach Excel chart series affected directly or through formula dependents."""
    sources: list[tuple[str, int, int, int, int, str]] = []
    for chart in workbook.charts:
        for series in chart.series:
            impact = (
                f"Excel chart {_chart_label(chart)!r} series "
                f"{_series_label(series)!r}"
            )
            for target in (series.categories_ref, series.values_ref, series.bubble_size_ref):
                if target is None:
                    continue
                resolution = resolve_reference(
                    workbook,
                    target,
                    host_sheet=chart.sheet,
                    require_within_sheet=False,
                )
                if resolution.status is not ReferenceStatus.RESOLVED:
                    continue
                for item in resolution.ranges:
                    sources.append(
                        (
                            item.sheet,
                            item.min_row,
                            item.min_col,
                            item.max_row,
                            item.max_col,
                            impact,
                        )
                    )
    for finding in findings:
        if finding.sheet is None or finding.location is None:
            continue
        try:
            row, column = coordinate_to_tuple(finding.location)
        except ValueError:
            continue
        finding_node: Node = (finding.sheet, row, column)
        affected_nodes = {finding_node}
        if dependency_graph is not None:
            affected_nodes.update(
                dependent_nodes_of(dependency_graph, finding_node)
            )
        impacts = {
            impact
            for sheet, min_row, min_col, max_row, max_col, impact in sources
            if any(
                source_sheet == sheet
                and min_row <= source_row <= max_row
                and min_col <= source_column <= max_col
                for source_sheet, source_row, source_column in affected_nodes
            )
        }
        finding.impacts = sorted({*finding.impacts, *impacts})
