"""Unified findings model — the single contract feeding UI, reports, history.

Diff engines emit `Finding` records without severity; the triage rule
engine (severity module) classifies them. Reports must never re-derive
diff logic from raw artifacts.
"""

from collections import Counter
from dataclasses import dataclass
from enum import StrEnum
from typing import Self

from pydantic import BaseModel, Field, field_serializer, model_validator


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
    ROW_KEY_CHANGED = "row_key_changed"
    COLUMN_KEY_CHANGED = "column_key_changed"
    ROW_GROWTH = "row_growth"
    COLUMN_GROWTH = "column_growth"
    SHEET_ADDED = "sheet_added"
    SHEET_REMOVED = "sheet_removed"
    HIDDEN_CHANGED = "hidden_changed"
    NAMED_RANGE_CHANGED = "named_range_changed"
    VBA_MODULE_CHANGED = "vba_module_changed"
    CELL_COMMENT_CHANGED = "cell_comment_changed"
    POWER_QUERY_CHANGED = "power_query_changed"
    CONNECTION_CHANGED = "connection_changed"
    EXTERNAL_CONNECTION = "external_connection"
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
    ACTIVE_CONTENT = "active_content"
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


class FindingProvenance(StrEnum):
    """Cycle history of a current-state observation, proven against the baseline.

    Only assigned when the cell has a confidently aligned baseline counterpart;
    otherwise it stays unset rather than guessing a regression.
    """

    NEW = "new"
    CHANGED = "changed"
    INHERITED = "inherited"
    HISTORICAL_PATTERN = "historical_pattern"


class Materiality(StrEnum):
    """Numeric materiality tier of a value delta.

    New findings use only magnitude/acceptance states. ``recent_restatement``
    remains solely so stored evidence from earlier versions can rehydrate.
    """

    NOISE = "noise"
    WITHIN_TOLERANCE = "within_tolerance"
    RECENT_RESTATEMENT = "recent_restatement"
    MATERIAL = "material"


class FindingTemporalContext(StrEnum):
    """Where a finding sits relative to its proved local period edge."""

    CURRENT_PERIOD = "current_period"
    RECENT_WINDOW = "recent_window"
    HISTORICAL = "historical"


class FindingExpectedReason(StrEnum):
    """Closed evidence reasons that may produce an Expected finding."""

    PERIOD_PROGRESSION = "period_progression"
    CADENCE_EXTENSION = "cadence_extension"
    ROLLING_WINDOW = "rolling_window"
    PROFILE_REFRESH = "profile_refresh"
    FIGURE_REFRESH = "figure_refresh"
    PRESENTATION_REORDER = "presentation_reorder"
    WAIVER = "waiver"


class FindingEvidenceTag(StrEnum):
    """Bounded, additive evidence carried outside finding identity."""

    DISPLAY_EQUIVALENT = "display_equivalent"
    ULP_SCALE = "ulp_scale"
    EXPLICIT_NA = "explicit_na"
    FORMULA_TEXT = "formula_text"
    FORMULA_PRESENCE = "formula_presence"
    CACHED_VALUE_ONLY = "cached_value_only"
    CONCENTRATED_POPULATION = "concentrated_population"
    CONTIGUOUS_POPULATION = "contiguous_population"
    SPARSE_MASS_POPULATION = "sparse_mass_population"
    STRUCTURAL_ERROR = "structural_error"
    EXACT_WRAPPER = "exact_wrapper"
    SHAPE_WRAPPER = "shape_wrapper"
    ADDED_REFERENCE = "added_reference"
    RESOLVED_DRIVER = "resolved_driver"
    EXACT_COLOCATION = "exact_colocation"


class FindingSubtype(StrEnum):
    """Precise mechanic behind a finding, orthogonal to its class."""

    # constant-value mechanics
    VALUE_ADDED_POPULATION = "added_population"
    VALUE_CLEARED_POPULATION = "cleared_population"
    VALUE_REPLACEMENT = "replacement"
    # formula-logic mechanics
    FORMULA_WRAPPED = "wrapped"
    FORMULA_UNWRAPPED = "unwrapped"
    #: Same saved error dominating one column: one systemic lookup-gap
    #: population, not N independent incidents.
    COLUMNAR_ERROR_POPULATION = "columnar_error_population"
    # region-level row/column events shared by every atomic of one region axis
    AXIS_ROLLING_TURNOVER = "rolling_turnover"
    AXIS_KEY_REPLACEMENT = "key_replacement"
    AXIS_KEY_DERIVED_LABEL = "derived_label_change"
    AXIS_PHYSICAL_INSERTION = "physical_insertion"
    AXIS_PHYSICAL_DELETION = "physical_deletion"
    AXIS_EXTENT_GROWTH = "extent_growth"
    # structural object events shared by every atomic of one rule or table edit
    OBJECT_ADDED = "object_added"
    OBJECT_REMOVED = "object_removed"
    OBJECT_RENAMED = "object_renamed"
    OBJECT_TARGET_CHANGED = "object_target_changed"
    OBJECT_CONDITION_CHANGED = "object_condition_changed"
    OBJECT_DISPLAY_CHANGED = "object_display_changed"
    OBJECT_STYLE_CHANGED = "object_style_changed"
    OBJECT_ORDER_CHANGED = "object_order_changed"
    OBJECT_COLUMNS_CHANGED = "object_columns_changed"
    OBJECT_SETTINGS_CHANGED = "object_settings_changed"


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
    #: Compatibility output for earlier history/report consumers. New producers
    #: set ``expected_reason`` and the validator derives this boolean.
    expected_growth: bool = False
    expected_reason: FindingExpectedReason | None = None
    #: Additive evidence detail; never part of the cross-run identity tuple.
    provenance: FindingProvenance | None = None
    subtype: FindingSubtype | None = None
    #: Numeric materiality tier for value deltas; additive, never identity.
    materiality: Materiality | None = None
    temporal_context: FindingTemporalContext | None = None
    evidence_tags: set[FindingEvidenceTag] = Field(
        default_factory=set,
        max_length=len(FindingEvidenceTag),
    )
    #: Stable identity of the underlying edit, shared by its fan-out atomics.
    event_key: str = ""
    sheet: str | None = None
    location: str | None = None  # current-side A1 ref / range / axis span
    baseline_location: str | None = None
    element: str | None = None  # named range / chart / pivot / slide element
    slide: str | None = None
    slide_index: int | None = Field(default=None, ge=1)
    baseline_slide_index: int | None = Field(default=None, ge=1)
    #: Transient producer provenance copied into the private focus sidecar
    #: before findings are serialized to reports, JSON, or history.
    focus_shape_id: int | None = Field(default=None, ge=1, exclude=True)
    baseline_focus_shape_id: int | None = Field(default=None, ge=1, exclude=True)
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

    @model_validator(mode="after")
    def derive_expected_growth(self) -> Self:
        if self.expected_reason is not None:
            self.expected_growth = True
        return self

    def mark_expected(self, reason: FindingExpectedReason) -> None:
        """Set the canonical reason and its legacy compatibility projection."""
        self.expected_reason = reason
        self.expected_growth = True

    @field_serializer("evidence_tags")
    def serialize_evidence_tags(
        self, evidence_tags: set[FindingEvidenceTag]
    ) -> list[str]:
        return sorted(tag.value for tag in evidence_tags)


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
        scope = (
            finding.sheet
            or (f"slide {finding.slide_index}" if finding.slide_index is not None else None)
            or finding.slide
            or "workbook/package"
        )
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
