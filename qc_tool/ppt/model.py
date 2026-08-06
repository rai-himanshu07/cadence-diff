"""Typed PowerPoint slide, shape, table, and complete chart snapshots."""

from __future__ import annotations

from dataclasses import dataclass, field

from qc_tool.io.model import ChartAxis, ChartDataLabels, ChartLegend


@dataclass(slots=True)
class ShapeContent:
    source_id: str
    source_index: int
    shape_id: int
    shape_type: str
    name: str
    left: int
    top: int
    width: int
    height: int
    z_order: int
    texts: list[str] = field(default_factory=list)
    is_placeholder: bool = False
    placeholder_type: str | None = None
    media_kind: str | None = None
    media_digest: str | None = None


@dataclass(slots=True)
class TableContent:
    rows: list[list[str]]
    source_id: str = ""
    source_index: int = 0
    shape_id: int = 0
    name: str = ""
    left: int = 0
    top: int = 0
    width: int = 0
    height: int = 0
    z_order: int = 0


@dataclass(slots=True)
class PptChartSeries:
    index: int
    order: int | None
    name: str | None
    categories: list[str]
    values: list[float | None]
    plot_index: int
    source_index: int
    source_id: str
    name_ref: str | None = None
    categories_ref: str | None = None
    values_ref: str | None = None
    bubble_size_ref: str | None = None
    data_labels: ChartDataLabels | None = None


@dataclass(frozen=True, slots=True)
class PptChartLabel:
    text: str
    series_source_index: int
    category_index: int
    category: str


@dataclass(slots=True)
class PptChartPlot:
    index: int
    chart_type: str
    series: list[PptChartSeries] = field(default_factory=list)
    axis_ids: tuple[str, ...] = ()
    axis_group: str = "none"
    grouping: str | None = None
    direction: str | None = None
    style: str | None = None
    data_labels: ChartDataLabels | None = None
    visible_labels: list[PptChartLabel] = field(default_factory=list)
    visible_label_values: list[str] = field(default_factory=list)
    source_id: str = ""


@dataclass(slots=True)
class ChartContent:
    #: First-plot compatibility views. New comparison code must use ``plots``.
    chart_type: str
    categories: list[str]
    series: list[tuple[str | None, list[float | None]]]
    source_id: str = ""
    source_index: int = 0
    shape_id: int = 0
    name: str = ""
    title: str | None = None
    plots: list[PptChartPlot] = field(default_factory=list)
    axes: list[ChartAxis] = field(default_factory=list)
    legend: ChartLegend | None = None
    display_blanks_as: str | None = None
    left: int = 0
    top: int = 0
    width: int = 0
    height: int = 0
    z_order: int = 0
    source_part: str | None = None

    @property
    def all_series(self) -> list[PptChartSeries]:
        return [series for plot in self.plots for series in plot.series]


@dataclass(slots=True)
class SlideContent:
    index: int
    title: str | None
    texts: list[str]
    tables: list[TableContent] = field(default_factory=list)
    charts: list[ChartContent] = field(default_factory=list)
    shapes: list[ShapeContent] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    shape_count: int = 0

    @property
    def display_name(self) -> str:
        return self.title or f"slide {self.index + 1}"


@dataclass(slots=True)
class DeckSnapshot:
    source_name: str
    slides: list[SlideContent] = field(default_factory=list)
    charts_available: bool = True
    chart_detail: str = "Charts parsed from raw PowerPoint chart parts"
    notes_available: bool = True
    notes_detail: str = "Speaker notes read separately from visible slide text"
    media_available: bool = True
    media_detail: str = "Embedded media bytes hashed without decoding"
