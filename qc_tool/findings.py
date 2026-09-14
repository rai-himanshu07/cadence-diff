"""Unified findings model — the single contract feeding UI, reports, history.

Diff engines emit `Finding` records without severity; the triage rule
engine (severity module) classifies them. Reports must never re-derive
diff logic from raw artifacts.
"""

import math
from enum import StrEnum
from typing import Any, Literal, Self

from pydantic import (
    BaseModel,
    Field,
    field_serializer,
    field_validator,
    model_validator,
)

from qc_tool.package import MEMBER_ID_PATTERN


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
    CIRCULAR_REFERENCE = "circular_reference"
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
    #: One analyst-confirmed logical rename, replacing what would otherwise
    #: be a SHEET_REMOVED+SHEET_ADDED pair (plan-20260913, Step 3).
    SHEET_RENAMED = "sheet_renamed"
    #: One analyst-confirmed logical column move (a mapped column's
    #: baseline-side letter genuinely differs from its current-side
    #: letter), replacing what would otherwise be spurious cell-level
    #: VALUE_CHANGED noise from positional pairing (plan-20260913, Step 12).
    COLUMN_MOVED = "column_moved"
    WORKBOOK_ADDED = "workbook_added"
    WORKBOOK_REMOVED = "workbook_removed"
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
    # Retained for stored runs recorded while output budgets existed (pre-0.2.0a2).
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
    PPT_MEDIA_CHANGED = "ppt_media_changed"
    PPT_DRAFT_TOKEN = "ppt_draft_token"
    PPT_REPEATED_CLAIM_MISMATCH = "ppt_repeated_claim_mismatch"
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



class NumericCounterfactualBasis(BaseModel):
    """Typed, versioned private basis for numeric WHAT-IF previews.

    This model is intentionally frozen/immutable and excluded from public
    serializations when attached to a `Finding`.
    """

    version: Literal[1] = 1
    baseline: float
    current: float
    number_format: str | None = None
    sheet: str = Field(min_length=1)
    location: str = Field(min_length=1)

    model_config = {"frozen": True}

    @field_validator("baseline", "current", mode="before")
    @classmethod
    def validate_number(cls, value: object) -> float:
        if isinstance(value, bool) or not isinstance(value, int | float):
            raise ValueError("counterfactual values must be numeric non-bools")
        converted = float(value)
        if not math.isfinite(converted):
            raise ValueError("counterfactual values must be finite")
        return converted


class SeriesAnchorV1(BaseModel):
    """Structural-only private anchor tying a finding to one logical series.

    Carries no value, formula, header, or metric text — only the coordinates a
    producer already proved while diffing an aligned period region.
    """

    version: Literal[1] = 1
    sheet: str = Field(min_length=1)
    current_region_id: str = Field(min_length=1)
    period_axis: Literal["rows", "columns"]
    series_index: int = Field(ge=1)
    period_index: int = Field(ge=1)

    model_config = {"frozen": True}


class SeriesAnchorV2(BaseModel):
    """A series anchor that also names which segment of the series it belongs to.

    ``new_period`` covers a period populated for the first time this cycle;
    ``cleared_period`` covers a historical cell cleared to blank while sibling
    measures keep the period alive. Neither has a baseline/current number pair,
    so neither carries a materiality tier or temporal context.
    """

    version: Literal[2] = 2
    sheet: str = Field(min_length=1)
    current_region_id: str = Field(min_length=1)
    period_axis: Literal["rows", "columns"]
    series_index: int = Field(ge=1)
    period_index: int = Field(ge=1)
    segment: Literal["restatement", "new_period", "cleared_period"] = "restatement"

    model_config = {"frozen": True}


class LogicalFindingAddress(BaseModel):
    """Bounded, content-free private address tying a finding to the saved
    logical member/sheet/region/column bindings that produced it
    (plan-20260913, Step 3).

    Never a sheet name, range, formula, or cell value -- only stable logical
    ids and, for a keyed row, a one-way digest of its identity key (never
    the raw key values themselves).
    """

    version: Literal[1] = 1
    member_id: str = Field(min_length=1, max_length=64)
    sheet_id: str = Field(min_length=1, max_length=64)
    region_id: str | None = Field(default=None, max_length=64)
    column_id: str | None = Field(default=None, max_length=64)
    row_key_digest: str | None = Field(default=None, min_length=64, max_length=64)

    model_config = {"frozen": True}


#: Readable anchor versions. V1 payloads keep their exact stored digest.
SeriesAnchor = SeriesAnchorV1 | SeriesAnchorV2


class MembershipCodec(BaseModel):
    """Dual-sided population membership: current rectangles + baseline pairing.

    ``shift`` mode covers a constant ``(dr, dc)`` offset shared by every
    member (positional/translated alignment); ``pairs`` mode covers the
    general case with an explicit ``(current, baseline)`` list. Exact
    ``(baseline, current)`` reconstruction from either mode is a unit-tested
    invariant, not merely documented behavior.
    """

    version: Literal[1] = 1
    current_rectangles: tuple[str, ...] = Field(default=(), max_length=2000)
    baseline_mode: Literal["shift", "pairs"]
    shift: tuple[int, int] | None = None
    pairs: tuple[tuple[str, str], ...] | None = Field(default=None, max_length=50_000)
    member_count: int = Field(ge=1)

    model_config = {"frozen": True}

    @model_validator(mode="after")
    def validate_mode_fields(self) -> Self:
        if self.baseline_mode == "shift" and self.shift is None:
            raise ValueError("shift mode requires a shift offset")
        if self.baseline_mode == "pairs" and not self.pairs:
            raise ValueError("pairs mode requires a non-empty explicit pairs list")
        return self


class PopulationSample(BaseModel):
    """One sampled member used for impacts/story/excerpt evidence only."""

    current_location: str = Field(min_length=1)
    baseline_location: str | None = None

    model_config = {"frozen": True}


class PopulationEvidence(BaseModel):
    """Group-first population evidence: membership, samples, shape digests.

    Root cause, story context, and impacts for a population follow the
    explicit population-level rules (root cause = the population identity
    key; impacts/story/excerpts come from ``samples`` only, labelled
    ``sampled``) rather than claiming full per-member atomic evidence.
    """

    version: Literal[1] = 1
    member_count: int = Field(ge=1)
    membership: MembershipCodec
    first: str = Field(min_length=1)
    last: str = Field(min_length=1)
    samples: tuple[PopulationSample, ...] = Field(default=(), max_length=5)
    shape_before_digest: str = Field(min_length=1)
    shape_after_digest: str = Field(min_length=1)
    #: Impacts computed for `samples` ONLY -- never the population's full
    #: downstream set. Kept distinct from `Finding.impacts` (which stays
    #: empty for population findings) so no renderer, story-evidence signal,
    #: or priority score can mistake a bounded sample for exhaustive
    #: downstream impacts.
    sampled_impacts: tuple[str, ...] = Field(default=())

    model_config = {"frozen": True}


class Finding(BaseModel):
    finding_id: str = ""  # assigned when a run collects findings
    artifact: str  # "excel" | "ppt" | "crosscheck"
    artifact_member: str = Field(
        default="primary",
        pattern=MEMBER_ID_PATTERN,
        exclude_if=lambda value: value == "primary",
    )
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

    #: Private counterfactual basis attached by the diff engine for preview.
    counterfactual_basis: NumericCounterfactualBasis | None = Field(
        default=None, exclude=True
    )

    #: Private logical-series anchor attached by `diff_region_values` only.
    series_anchor: SeriesAnchorV1 | SeriesAnchorV2 | None = Field(
        default=None, exclude=True
    )

    #: Private logical member/sheet/region/column address, attached only
    #: when an execution-bound region produced this finding (Step 3).
    logical_address: LogicalFindingAddress | None = Field(
        default=None, exclude=True
    )

    #: Group-first population evidence. None means atomic (the legacy,
    #: always-on default). When set, ``element`` is "population" and
    #: ``location`` is the bounding range of the current-side rectangles.
    population: PopulationEvidence | None = None

    @model_validator(mode="after")
    def derive_expected_growth(self) -> Self:
        if self.expected_reason is not None:
            self.expected_growth = True
        return self

    @classmethod
    def from_trusted_payload(cls, payload: dict[str, Any]) -> "Finding":
        """Fast constructor for payloads this application wrote itself.

        Stored block lines were validated when produced; re-validating them
        on every read pass dominated monster-run recording. This coerces
        exactly the typed fields and skips constraint checks — equality with
        ``model_validate`` is pinned by tests over every stored field.
        Never use it on external or hand-edited data.
        """
        data: dict[str, Any] = dict(payload)
        data["finding_class"] = FindingClass(data["finding_class"])
        severity = data.get("severity")
        if severity is not None:
            data["severity"] = Severity(severity)
        reason = data.get("expected_reason")
        if reason is not None:
            data["expected_reason"] = FindingExpectedReason(reason)
        provenance = data.get("provenance")
        if provenance is not None:
            data["provenance"] = FindingProvenance(provenance)
        subtype = data.get("subtype")
        if subtype is not None:
            data["subtype"] = FindingSubtype(subtype)
        materiality = data.get("materiality")
        if materiality is not None:
            data["materiality"] = Materiality(materiality)
        temporal = data.get("temporal_context")
        if temporal is not None:
            data["temporal_context"] = FindingTemporalContext(temporal)
        data["evidence_tags"] = {
            FindingEvidenceTag(tag) for tag in data.get("evidence_tags", ())
        }
        data["impacts"] = list(data.get("impacts", ()))
        for key in ("baseline_excerpt", "current_excerpt"):
            excerpt = data.get(key)
            if excerpt is not None:
                data[key] = GridExcerpt.model_construct(**excerpt)
        basis = data.get("counterfactual_basis")
        if basis is not None:
            data["counterfactual_basis"] = (
                NumericCounterfactualBasis.model_construct(**basis)
            )
        anchor = data.get("series_anchor")
        if isinstance(anchor, dict):
            anchor_model = (
                SeriesAnchorV2 if anchor.get("version") == 2 else SeriesAnchorV1
            )
            data["series_anchor"] = anchor_model.model_construct(**anchor)
        address = data.get("logical_address")
        if address is not None:
            data["logical_address"] = LogicalFindingAddress.model_construct(**address)
        population = data.get("population")
        if isinstance(population, dict):
            membership = population["membership"]
            data["population"] = PopulationEvidence.model_construct(
                **{
                    **population,
                    "membership": MembershipCodec.model_construct(
                        **{
                            **membership,
                            "current_rectangles": tuple(
                                membership.get("current_rectangles", ())
                            ),
                            "shift": (
                                tuple(membership["shift"])
                                if membership.get("shift") is not None
                                else None
                            ),
                            "pairs": (
                                tuple(
                                    tuple(pair) for pair in membership["pairs"]
                                )
                                if membership.get("pairs") is not None
                                else None
                            ),
                        }
                    ),
                    "samples": tuple(
                        PopulationSample.model_construct(**sample)
                        for sample in population.get("samples", ())
                    ),
                    "sampled_impacts": tuple(population.get("sampled_impacts", ())),
                }
            )
        finding = cls.model_construct(**data)
        if finding.expected_reason is not None:
            finding.expected_growth = True
        return finding

    def mark_expected(self, reason: FindingExpectedReason) -> None:
        """Set the canonical reason and its legacy compatibility projection."""
        self.expected_reason = reason
        self.expected_growth = True

    @field_serializer("evidence_tags")
    def serialize_evidence_tags(
        self, evidence_tags: set[FindingEvidenceTag]
    ) -> list[str]:
        return sorted(tag.value for tag in evidence_tags)
