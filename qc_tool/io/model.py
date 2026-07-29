"""Unified read-only snapshot model for workbooks across xlsx/xlsm/xlsb.

Snapshots are the single contract the diff engines consume; loaders for
every format populate the same structures. Fields a format cannot provide
stay ``None`` and workbook-level capability flags distinguish formula-record
presence from decoded formula text.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from typing import TypeGuard

CellValue = str | float | int | bool | dt.date | dt.datetime | None
SerializedCellValue = str | float | int | bool | None


def is_cell_value(value: object) -> TypeGuard[CellValue]:
    """Whether a parsed value belongs to the cross-format snapshot contract."""
    return value is None or isinstance(
        value, str | float | int | bool | dt.datetime | dt.date
    )


def display_cell_value(value: object) -> str:
    """Stable display text for findings, excerpts, and report boundaries."""
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        return value.isoformat()
    if isinstance(value, dt.date):
        return value.isoformat()
    return str(value)


def serialize_cell_value(value: CellValue) -> SerializedCellValue:
    """Return a JSON-safe scalar without losing temporal meaning."""
    if isinstance(value, dt.datetime | dt.date):
        return value.isoformat()
    return value

#: Excel error literals (values or embedded in formula text).
ERROR_LITERALS = frozenset(
    {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}
)


@dataclass(slots=True)
class CellRecord:
    """One populated cell. ``value`` is the constant or cached formula result."""

    row: int  # 1-based
    column: int  # 1-based
    value: CellValue
    formula: str | None = None
    is_formula: bool | None = None
    number_format: str | None = None
    style_key: str | None = None

    @property
    def has_formula(self) -> bool:
        """Whether the loader positively identified a formula at this cell."""
        return self.is_formula is True or self.formula is not None

    @property
    def is_error(self) -> bool:
        return isinstance(self.value, str) and self.value in ERROR_LITERALS


@dataclass(slots=True)
class NamedRange:
    name: str
    target: str


@dataclass(slots=True)
class ChartSeries:
    index: int
    values_ref: str | None
    categories_ref: str | None
    order: int | None = None
    name_ref: str | None = None
    name_text: str | None = None
    plot_index: int = 0
    source_index: int = 0
    source_id: str = ""
    bubble_size_ref: str | None = None
    data_labels: ChartDataLabels | None = None


@dataclass(slots=True)
class ChartDataLabels:
    position: str | None = None
    show_value: bool | None = None
    show_category_name: bool | None = None
    show_series_name: bool | None = None
    show_percent: bool | None = None
    show_legend_key: bool | None = None
    show_bubble_size: bool | None = None
    show_leader_lines: bool | None = None
    number_format: str | None = None
    separator: str | None = None


@dataclass(slots=True)
class ChartPlot:
    index: int
    chart_type: str
    series: list[ChartSeries] = field(default_factory=list)
    axis_ids: tuple[str, ...] = ()
    axis_group: str = "none"
    grouping: str | None = None
    direction: str | None = None
    style: str | None = None
    data_labels: ChartDataLabels | None = None
    source_id: str = ""


@dataclass(slots=True)
class ChartAxis:
    axis_id: str
    axis_type: str
    source_index: int
    position: str | None = None
    cross_axis_id: str | None = None
    minimum: float | None = None
    maximum: float | None = None
    major_unit: float | None = None
    minor_unit: float | None = None
    log_base: float | None = None
    orientation: str | None = None
    crosses: str | None = None
    crosses_at: float | None = None
    number_format: str | None = None
    display_unit: str | None = None
    major_tick_mark: str | None = None
    minor_tick_mark: str | None = None
    title: str | None = None
    deleted: bool | None = None


@dataclass(slots=True)
class ChartLegend:
    position: str | None = None
    overlay: bool | None = None
    deleted: bool | None = None


@dataclass(slots=True)
class ChartAnchor:
    anchor_type: str
    from_col: int | None = None
    from_row: int | None = None
    from_col_offset: int | None = None
    from_row_offset: int | None = None
    to_col: int | None = None
    to_row: int | None = None
    to_col_offset: int | None = None
    to_row_offset: int | None = None
    x: int | None = None
    y: int | None = None
    width: int | None = None
    height: int | None = None

    @property
    def signature(self) -> str:
        values = (
            self.anchor_type,
            self.from_col,
            self.from_row,
            self.from_col_offset,
            self.from_row_offset,
            self.to_col,
            self.to_row,
            self.to_col_offset,
            self.to_row_offset,
            self.x,
            self.y,
            self.width,
            self.height,
        )
        return ":".join("" if value is None else str(value) for value in values)


@dataclass(slots=True)
class ChartDescriptor:
    sheet: str
    title: str | None
    chart_type: str
    series: list[ChartSeries] = field(default_factory=list)
    source_index: int = 0
    source_id: str = ""
    source_name: str | None = None
    source_part: str | None = None
    plots: list[ChartPlot] = field(default_factory=list)
    axes: list[ChartAxis] = field(default_factory=list)
    legend: ChartLegend | None = None
    anchor: ChartAnchor | None = None
    display_blanks_as: str | None = None


@dataclass(slots=True)
class PivotDescriptor:
    name: str
    location_ref: str | None
    source_sheet: str | None
    source_ref: str | None


@dataclass(slots=True)
class TableDescriptor:
    sheet: str
    name: str
    display_name: str
    cell_range: str
    columns: list[str] = field(default_factory=list)
    header_row_count: int = 1
    totals_row_count: int = 0
    source_id: int | None = None


@dataclass(slots=True)
class DataValidationDescriptor:
    sheet: str
    source_index: int
    source_id: str
    target_ranges: tuple[str, ...]
    validation_type: str | None = None
    operator: str | None = None
    formula1: str | None = None
    formula2: str | None = None
    allow_blank: bool | None = None
    dropdown_suppressed: bool | None = None
    show_error_message: bool | None = None
    show_input_message: bool | None = None
    error_style: str | None = None
    prompt_title: str | None = None
    prompt: str | None = None
    error_title: str | None = None
    error: str | None = None


@dataclass(slots=True)
class ConditionalFormatDescriptor:
    sheet: str
    source_index: int
    source_id: str
    target_ranges: tuple[str, ...]
    rule_type: str
    operator: str | None = None
    formulas: tuple[str, ...] = ()
    priority: int | None = None
    stop_if_true: bool | None = None
    text: str | None = None
    time_period: str | None = None
    rank: int | None = None
    percent: bool | None = None
    bottom: bool | None = None
    above_average: bool | None = None
    equal_average: bool | None = None
    standard_deviations: int | None = None
    style_key: str | None = None
    semantic_supported: bool = True
    semantic_detail: str = ""
    style_supported: bool = True
    style_detail: str = ""


@dataclass(slots=True)
class SheetSnapshot:
    name: str
    visibility: str  # "visible" | "hidden" | "veryHidden"
    max_row: int
    max_column: int
    cells: dict[tuple[int, int], CellRecord] = field(default_factory=dict)
    hidden_rows: frozenset[int] = frozenset()
    hidden_columns: frozenset[int] = frozenset()

    def cell(self, ref: str) -> CellRecord | None:
        """Look up a cell by A1 reference (e.g. ``"C7"``)."""
        from openpyxl.utils.cell import coordinate_to_tuple

        row, column = coordinate_to_tuple(ref)
        return self.cells.get((row, column))


@dataclass(slots=True)
class WorkbookSnapshot:
    source_name: str
    file_format: str  # "xlsx" | "xlsm" | "xlsb"
    formulas_available: bool
    styles_available: bool
    formula_presence_available: bool = False
    tables_available: bool = True
    charts_available: bool = True
    interaction_rules_available: bool = True
    interaction_rules_supported: bool = True
    conditional_format_styles_supported: bool = True
    formula_source: str | None = None
    formula_detail: str = ""
    chart_detail: str = ""
    interaction_rule_detail: str = ""
    conditional_format_style_detail: str = ""
    sheets: list[SheetSnapshot] = field(default_factory=list)
    named_ranges: list[NamedRange] = field(default_factory=list)
    charts: list[ChartDescriptor] = field(default_factory=list)
    pivots: list[PivotDescriptor] = field(default_factory=list)
    tables: list[TableDescriptor] = field(default_factory=list)
    data_validations: list[DataValidationDescriptor] = field(default_factory=list)
    conditional_formats: list[ConditionalFormatDescriptor] = field(default_factory=list)
    calculation_mode: str | None = None
    full_calc_on_load: bool | None = None
    external_links: list[str] = field(default_factory=list)

    @property
    def sheet_names(self) -> list[str]:
        return [s.name for s in self.sheets]

    def sheet(self, name: str) -> SheetSnapshot:
        for sheet in self.sheets:
            if sheet.name == name:
                return sheet
        raise KeyError(f"no sheet named {name!r} in {self.source_name}")
