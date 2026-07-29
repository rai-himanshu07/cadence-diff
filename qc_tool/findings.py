"""Unified findings model — the single contract feeding UI, reports, history.

Diff engines emit `Finding` records without severity; the triage rule
engine (severity module) classifies them. Reports must never re-derive
diff logic from raw artifacts.
"""

from collections import Counter
from dataclasses import dataclass
from enum import StrEnum

from pydantic import BaseModel, Field


class FindingClass(StrEnum):
    # cell-level
    VALUE_CHANGED = "value_changed"
    FORMULA_ERROR = "formula_error"
    FORMULA_HARDCODED = "formula_hardcoded"
    FORMULA_REMOVED = "formula_removed"
    FORMULA_MISSING = "formula_missing"
    FORMULA_CACHE_MISSING = "formula_cache_missing"
    FORMULA_NOT_EXTENDED = "formula_not_extended"
    FORMULA_LOGIC_CHANGED = "formula_logic_changed"
    FORMULA_INCONSISTENT = "formula_inconsistent"
    NUMBER_FORMAT_CHANGED = "number_format_changed"
    STYLE_CHANGED = "style_changed"
    # axis / structure
    ROW_DELETED = "row_deleted"
    COLUMN_DELETED = "column_deleted"
    ROW_INSERTED = "row_inserted"
    COLUMN_INSERTED = "column_inserted"
    ROW_GROWTH = "row_growth"
    COLUMN_GROWTH = "column_growth"
    SHEET_ADDED = "sheet_added"
    SHEET_REMOVED = "sheet_removed"
    HIDDEN_CHANGED = "hidden_changed"
    NAMED_RANGE_CHANGED = "named_range_changed"
    TABLE_STRUCTURE_CHANGED = "table_structure_changed"
    DATA_VALIDATION_CHANGED = "data_validation_changed"
    CONDITIONAL_FORMAT_CHANGED = "conditional_format_changed"
    CHART_STRUCTURE_CHANGED = "chart_structure_changed"
    CHART_PLOT_CHANGED = "chart_plot_changed"
    CHART_SERIES_CHANGED = "chart_series_changed"
    CHART_AXIS_CHANGED = "chart_axis_changed"
    CHART_LEGEND_CHANGED = "chart_legend_changed"
    CHART_LABELS_CHANGED = "chart_labels_changed"
    CHART_GEOMETRY_CHANGED = "chart_geometry_changed"
    PIVOT_SOURCE_CHANGED = "pivot_source_changed"
    REGION_UNPAIRED = "region_unpaired"
    ALIGNMENT_LOW_CONFIDENCE = "alignment_low_confidence"
    FINDINGS_CAPPED = "findings_capped"
    PERIOD_DUPLICATE = "period_duplicate"
    PERIOD_OUT_OF_ORDER = "period_out_of_order"
    PERIOD_GAP = "period_gap"
    CALCULATION_MODE = "calculation_mode"
    EXTERNAL_LINK = "external_link"
    NAMED_RANGE_INVALID = "named_range_invalid"
    CHART_REFERENCE_INVALID = "chart_reference_invalid"
    CHART_LENGTH_MISMATCH = "chart_length_mismatch"
    PIVOT_SOURCE_INVALID = "pivot_source_invalid"
    HIDDEN_CONTENT = "hidden_content"
    REQUIRED_VALUE_MISSING = "required_value_missing"
    DUPLICATE_KEY = "duplicate_key"
    NUMERIC_BOUND_VIOLATION = "numeric_bound_violation"
    TIE_OUT_MISMATCH = "tie_out_mismatch"
    CONTROL_INVALID = "control_invalid"
    WAIVER_EXPIRED = "waiver_expired"
    # ppt
    SLIDE_ADDED = "slide_added"
    SLIDE_REMOVED = "slide_removed"
    SLIDE_REORDERED = "slide_reordered"
    SLIDE_TEXT_CHANGED = "slide_text_changed"
    TABLE_VALUE_CHANGED = "table_value_changed"
    CHART_VALUE_CHANGED = "chart_value_changed"
    PPT_TABLE_STRUCTURE_CHANGED = "ppt_table_structure_changed"
    PPT_CHART_STRUCTURE_CHANGED = "ppt_chart_structure_changed"
    PPT_CHART_PLOT_CHANGED = "ppt_chart_plot_changed"
    PPT_CHART_SERIES_CHANGED = "ppt_chart_series_changed"
    PPT_CHART_AXIS_CHANGED = "ppt_chart_axis_changed"
    PPT_CHART_LEGEND_CHANGED = "ppt_chart_legend_changed"
    PPT_CHART_LABELS_CHANGED = "ppt_chart_labels_changed"
    PPT_SHAPE_GEOMETRY_CHANGED = "ppt_shape_geometry_changed"
    PPT_DRAFT_TOKEN = "ppt_draft_token"
    PPT_EMPTY_SLIDE = "ppt_empty_slide"
    PPT_DUPLICATE_TITLE = "ppt_duplicate_title"
    PPT_REQUIRED_SLIDE_MISSING = "ppt_required_slide_missing"
    PPT_PERIOD_INCONSISTENT = "ppt_period_inconsistent"
    PPT_TABLE_BLANK = "ppt_table_blank"
    PPT_CHART_LENGTH_MISMATCH = "ppt_chart_length_mismatch"
    PPT_CHART_VALUE_MISSING = "ppt_chart_value_missing"
    # crosscheck
    CROSSCHECK_MISMATCH = "crosscheck_mismatch"
    CROSSCHECK_UNRESOLVED = "crosscheck_unresolved"
    PACKAGE_PERIOD_MISMATCH = "package_period_mismatch"


class Severity(StrEnum):
    CRITICAL = "critical"
    WARNING = "warning"
    INFO = "info"
    EXPECTED = "expected"


class GridExcerpt(BaseModel):
    """A small neighborhood of sheet cells around a finding, for in-UI context."""

    cols: list[str]  # column letters
    rows: list[int]  # row numbers
    cells: list[list[str]]  # display values (formula text when uncached)
    hit_row: int | None = None  # index into rows
    hit_col: int | None = None  # index into cols


class Finding(BaseModel):
    finding_id: str = ""  # assigned when a run collects findings
    artifact: str  # "excel" | "ppt" | "crosscheck"
    finding_class: FindingClass
    severity: Severity | None = None  # assigned by the triage rule engine
    expected_growth: bool = False
    sheet: str | None = None
    location: str | None = None  # current-side A1 ref / range / axis span
    baseline_location: str | None = None
    element: str | None = None  # named range / chart / pivot / slide element
    slide: str | None = None
    baseline_value: str | None = None
    current_value: str | None = None
    message: str
    impacts: list[str] = Field(default_factory=list)
    baseline_excerpt: GridExcerpt | None = None
    current_excerpt: GridExcerpt | None = None
    #: Analyst review: free-text note and manual severity override.
    analyst_comment: str = ""
    severity_overridden: bool = False
    root_cause_key: str = ""
    waiver_reason: str = ""
    waiver_expires: str = ""


@dataclass(slots=True)
class FindingsBudgetResult:
    findings: list[Finding]
    omitted_by_artifact: dict[str, int]
    global_omitted: int = 0


def limit_findings(
    findings: list[Finding],
    *,
    max_per_class_scope: int = 500,
    max_total: int = 10_000,
) -> FindingsBudgetResult:
    """Bound output volume and disclose every omitted finding."""
    if max_per_class_scope < 1 or max_total < 2:
        raise ValueError("findings budgets must retain at least one detail and summary")
    retained: list[Finding] = []
    kept: Counter[tuple[str, FindingClass, str]] = Counter()
    omitted: Counter[tuple[str, FindingClass, str]] = Counter()
    for finding in findings:
        scope = finding.sheet or finding.slide or "workbook/package"
        key = (finding.artifact, finding.finding_class, scope)
        if kept[key] < max_per_class_scope:
            retained.append(finding)
            kept[key] += 1
        else:
            omitted[key] += 1

    omitted_by_artifact: Counter[str] = Counter()
    for (artifact, finding_class, scope), omitted_count in sorted(
        omitted.items(),
        key=lambda item: (
            item[0][0],
            item[0][2],
            item[0][1].value,
        ),
    ):
        omitted_by_artifact[artifact] += omitted_count
        retained.append(
            Finding(
                artifact=artifact,
                finding_class=FindingClass.FINDINGS_CAPPED,
                sheet=scope if artifact == "excel" else None,
                slide=scope if artifact == "ppt" else None,
                element=finding_class.value,
                current_value=(
                    f"{max_per_class_scope} retained; {omitted_count} omitted"
                ),
                message=(
                    f"{scope}: output budget retained the first "
                    f"{max_per_class_scope} {finding_class.value} findings and "
                    f"omitted {omitted_count}; affected coverage is degraded"
                ),
            )
        )

    global_omitted = 0
    if len(retained) > max_total:
        keep_count = max_total - 1
        dropped = retained[keep_count:]
        retained = retained[:keep_count]
        global_omitted = len(dropped)
        for finding in dropped:
            omitted_by_artifact[finding.artifact] += 1
        retained.append(
            Finding(
                artifact="run",
                finding_class=FindingClass.FINDINGS_CAPPED,
                element="global",
                current_value=f"{keep_count} retained; {global_omitted} omitted",
                message=(
                    f"Run output budget retained {keep_count} findings and omitted "
                    f"{global_omitted}; all affected coverage is degraded"
                ),
            )
        )
    return FindingsBudgetResult(
        findings=retained,
        omitted_by_artifact=dict(omitted_by_artifact),
        global_omitted=global_omitted,
    )
