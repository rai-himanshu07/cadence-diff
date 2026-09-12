"""Unified read-only snapshot model for workbooks across xlsx/xlsm/xlsb.

Snapshots are the single contract the diff engines consume; loaders for
every format populate the same structures. Fields a format cannot provide
stay ``None`` and workbook-level capability flags distinguish formula-record
presence from decoded formula text.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from enum import StrEnum
from typing import Literal, TypeGuard

from qc_tool.io.ooxml_metadata import WorkbookMetadataScan
from qc_tool.io.vba import VbaProjectScan

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
    #: Canonical R1C1 text (``=``-prefixed, matching ``formula``'s own
    #: convention), populated only when an adapter can supply it per
    #: definition (currently: the cadence-diff native engine). ``None`` means no
    #: adapter-supplied R1C1 is available; consumers fall back to computing
    #: it themselves from ``formula``.
    formula_r1c1: str | None = None
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
class WorkbookWorkload:
    format: Literal["ooxml", "xlsb"] = "ooxml"
    metrics_available: bool = True
    cell_count: int = 0
    worksheet_xml_bytes: int = 0
    worksheet_binary_bytes: int = 0
    shared_string_bytes: int = 0
    styles_bytes: int = 0
    style_count: int = 0
    formula_count: int = 0
    sheet_count: int = 0
    largest_sheet_rows: int = 0
    largest_sheet_columns: int = 0
    largest_sheet_area: int = 0
    warning_reasons: tuple[str, ...] = ()
    override_used: bool = False

    @property
    def degraded(self) -> bool:
        return bool(self.warning_reasons) or self.override_used

    @property
    def detail(self) -> str:
        if not self.metrics_available:
            return "XLSB workload metrics unavailable"
        if self.format == "xlsb":
            metrics = (
                f"{self.cell_count:,} retained cells; "
                f"{self.formula_count:,} formulas; "
                f"{self.worksheet_binary_bytes / (1024 * 1024):.1f} MiB "
                f"BIFF12 worksheets; "
                f"{self.shared_string_bytes / (1024 * 1024):.1f} MiB shared strings; "
                f"{self.styles_bytes / (1024 * 1024):.1f} MiB styles; "
                f"{self.sheet_count:,} sheets; "
                f"largest measured extent {self.largest_sheet_rows:,} x "
                f"{self.largest_sheet_columns:,} "
                f"({self.largest_sheet_area:,} cells)"
            )
            if self.warning_reasons:
                return f"{metrics}; " + "; ".join(self.warning_reasons)
            return metrics
        metrics = (
            f"{self.cell_count:,} physical cells; "
            f"{self.worksheet_xml_bytes / (1024 * 1024):.1f} MiB worksheet XML; "
            f"{self.shared_string_bytes / (1024 * 1024):.1f} MiB shared strings; "
            f"{self.style_count:,} cell styles"
        )
        if self.warning_reasons:
            return f"{metrics}; " + "; ".join(self.warning_reasons)
        return metrics


class WorkbookRiskKind(StrEnum):
    EXTERNAL_WORKBOOK_LINK = "external_workbook_link"
    EXTERNAL_RELATIONSHIP = "external_relationship"
    EXTERNAL_DATA_CONNECTION = "external_data_connection"
    QUERY_TABLE = "query_table"
    VBA_PROJECT = "vba_project"
    EXCEL4_MACRO_SHEET = "excel4_macro_sheet"
    ACTIVEX_CONTROL = "activex_control"
    EMBEDDED_OLE = "embedded_ole"
    CONTROL_CONTENT = "control_content"
    DIALOG_SHEET = "dialog_sheet"
    CUSTOM_OFFICE_UI = "custom_office_ui"
    UNREADABLE_RELATIONSHIP_METADATA = "unreadable_relationship_metadata"


@dataclass(frozen=True, slots=True)
class WorkbookRisk:
    kind: WorkbookRiskKind
    count: int = 1

    def __post_init__(self) -> None:
        if self.count < 1:
            raise ValueError("workbook risk count must be positive")


@dataclass(slots=True)
class NamedRange:
    name: str
    target: str
    sheet: str | None = None
    hidden: bool = False

    @property
    def scope_label(self) -> str:
        return "workbook" if self.sheet is None else self.sheet

    @property
    def qualified_name(self) -> str:
        """Identity that stays distinct when one name exists in several scopes."""
        return self.name if self.sheet is None else f"{self.sheet}!{self.name}"


@dataclass(frozen=True, slots=True)
class FormulaRangeDescriptor:
    """Authoritative range declared by one anchor formula element."""

    sheet: str
    anchor_row: int
    anchor_column: int
    cell_range: str
    formula_type: str
    always_calculate: bool | None = None


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


@dataclass(frozen=True, slots=True)
class FormulaTextCoverage:
    """Compact, coordinate-set-free formula-text completeness disclosure.

    ``state`` starts at ``"none"`` (no external-engine formula text merged,
    either because no adapter ran or because it was refused). A validated
    merge with zero missing coordinates moves it to ``"complete"``, matching
    the historical all-or-nothing meaning of ``formulas_available``. A
    validated merge missing at least one expected coordinate moves it to
    ``"partial"`` -- coordinate-aware comparison may still use whichever
    pairs both sides actually merged, but dependency/circular/impact claims
    and consistency checks over an incomplete region remain unavailable.
    Any unexpected coordinate stays fatal before mutation and never produces
    a ``"partial"`` result.
    """

    state: Literal["none", "complete", "partial"] = "none"
    expected_count: int = 0
    merged_count: int = 0
    missing_count: int = 0
    detail: str = ""


@dataclass(frozen=True, slots=True)
class ExternalLinkReachability:
    """Bounded, aggregate-only verdict for one workbook's passive links.

    Never retains formula text, defined-name text, coordinates, or paths --
    only counts and a boolean verdict. ``proven`` requires complete trusted
    formula text plus a complete, bounded defined-name collection; a passive
    link is only ever reported ``live=False`` when ``proven`` is also True.
    """

    proven: bool = False
    live: bool = False
    direct_reference_count: int = 0
    transitive_reference_count: int = 0
    detail: str = ""


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
    #: Resolved cached-values decoder identity. Unlike
    #: ``values_engine_fallback_detail``, this always names the engine that
    #: actually supplied values after any ``auto`` fallback.
    values_source: str | None = field(default=None, compare=False)
    formula_detail: str = ""
    chart_detail: str = ""
    interaction_rule_detail: str = ""
    conditional_format_style_detail: str = ""
    defined_name_scope_available: bool = False
    defined_name_scope_detail: str = ""
    #: Whether XLSB per-cell number-format/date-system metadata (`xl/styles.bin`
    #: cell-XF table plus the workbook date-system flag) was read completely.
    #: Distinct from `styles_available`, which stays False for xlsb because
    #: fonts/fills/borders and named cell styles remain unread.
    number_formats_available: bool = False
    number_format_detail: str = ""
    #: XLSB-only completeness disclosure for the current external-engine
    #: formula-text merge; OOXML formats leave this at its default ("none")
    #: because `formulas_available` alone already proves complete text there.
    formula_text_coverage: FormulaTextCoverage = field(
        default_factory=FormulaTextCoverage
    )
    #: Set only when a passive external-workbook link was classified during
    #: loading; `None` means no passive link was present to evaluate.
    external_link_reachability: ExternalLinkReachability | None = None
    vba: VbaProjectScan = field(default_factory=VbaProjectScan)
    metadata: WorkbookMetadataScan = field(default_factory=WorkbookMetadataScan)
    sheets: list[SheetSnapshot] = field(default_factory=list)
    named_ranges: list[NamedRange] = field(default_factory=list)
    formula_ranges: list[FormulaRangeDescriptor] = field(default_factory=list)
    charts: list[ChartDescriptor] = field(default_factory=list)
    pivots: list[PivotDescriptor] = field(default_factory=list)
    tables: list[TableDescriptor] = field(default_factory=list)
    data_validations: list[DataValidationDescriptor] = field(default_factory=list)
    conditional_formats: list[ConditionalFormatDescriptor] = field(default_factory=list)
    calculation_mode: str | None = None
    full_calc_on_load: bool | None = None
    intrinsic_risks: list[WorkbookRisk] = field(default_factory=list)
    workload: WorkbookWorkload = field(
        default_factory=WorkbookWorkload,
        compare=False,
    )
    #: Set only when `_xlsb_values_engine="auto"` resolved to the native
    #: kernel but it then failed at runtime, so values decoding fell back to
    #: pyxlsb; "" means no fallback occurred (native was not used, ran
    #: cleanly, or was explicitly requested and failed closed instead of
    #: falling back). Fixed, content-free text -- never a path or value.
    values_engine_fallback_detail: str = ""

    @property
    def sheet_names(self) -> list[str]:
        return [s.name for s in self.sheets]

    def sheet(self, name: str) -> SheetSnapshot:
        for sheet in self.sheets:
            if sheet.name == name:
                return sheet
        raise KeyError(f"no sheet named {name!r} in {self.source_name}")
