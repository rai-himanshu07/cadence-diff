"""Saved logical input contract (``InputContractV1``).

This is the durable, analyst-authored structural contract stored inside a
named profile (``DeliverableProfile.input_contract``) and included in
canonical profile bytes / ``profile_sha256``. It describes logical member,
sheet, region, column, selector, and period-band identities and durable
*preferences* -- never per-run resolved physical coordinates, preview
content, or selector values. See
``docs/plans/plan-20260913-mode-aware-configuration-wizard.md`` Step 1.

A profile with no ``input_contract`` (``None``, the default) behaves exactly
as before this module existed: legacy physical-name profile fields and
automatic detection drive execution, byte-identical canonical hash.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
from typing import Literal, Self

from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
from pydantic import BaseModel, Field, field_validator, model_validator

from qc_tool.package import MEMBER_ID_PATTERN

#: Schema version of the saved logical input contract. Bump only when the
#: shape changes in a way that requires the setup pipeline to re-validate an
#: already-resolved configuration (see ``resolved_input.py``).
INPUT_CONTRACT_VERSION = 1

RegionMode = Literal["automatic", "keyed", "positional", "excluded"]
HeaderIntent = Literal["automatic", "no_header", "first_data_row"]
AlignmentRole = Literal["none", "identity", "ordinal"]
ColumnComparisonPolicy = Literal["normal", "ignore", "expected_refresh"]
DuplicateKeyPolicy = Literal["skip", "occurrence", "position"]
BlankKeyPolicy = Literal["system_default", "tolerate", "block"]
PeriodAxis = Literal["rows", "columns"]
PeriodCadenceKind = Literal["month", "week", "quarter", "date"]
SheetDefaultMode = Literal["automatic", "positional"]

#: Version of the conservative built-in blank-key tolerance that
#: ``"system_default"`` resolves to at run time (Step 8). A region's saved
#: choice of the literal string ``"system_default"`` is itself stable; only
#: what that default *means* is versioned, exactly like
#: ``qc_tool.config.profile.DECISION_MODE_POLICY_VERSION``.
BLANK_KEY_POLICY_DEFAULT_VERSION = 1

_LOGICAL_ID_RE = re.compile(r"^[a-z][a-z0-9_]{0,63}$")


def _validate_logical_id(value: str, *, field_name: str = "id") -> str:
    if not _LOGICAL_ID_RE.fullmatch(value):
        raise ValueError(
            f"{field_name} must be lowercase snake_case, start with a letter, "
            "and be at most 64 characters"
        )
    return value


def _validate_a1_cell(value: str) -> str:
    try:
        coordinate_to_tuple(value)
    except ValueError as exc:
        raise ValueError("must be an A1 cell reference") from exc
    return value.upper()


def _range_bounds(cell_range: str) -> tuple[int, int, int, int]:
    """(min_col, min_row, max_col, max_row) for a plain rectangular range.

    Accepts a single cell (``"A1"``) or a colon range (``"A1:B2"``). Whole
    column/row ranges are a run-time resolution concern (Step 6/8), not a
    saved-contract shape; reject them here rather than guess.
    """
    if ":" in cell_range:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        if min_col is None or min_row is None or max_col is None or max_row is None:
            raise ValueError(
                f"{cell_range!r} must be a bounded rectangular range, not a "
                "whole row/column"
            )
        return (min_col, min_row, max_col, max_row)
    row, col = coordinate_to_tuple(cell_range)
    return (col, row, col, row)


def _ranges_overlap(a: str, b: str) -> bool:
    a_min_col, a_min_row, a_max_col, a_max_row = _range_bounds(a)
    b_min_col, b_min_row, b_max_col, b_max_row = _range_bounds(b)
    return (
        a_min_col <= b_max_col
        and b_min_col <= a_max_col
        and a_min_row <= b_max_row
        and b_min_row <= a_max_row
    )


class StructuralExclusionContract(BaseModel):
    """A durable, expiring decision to exclude a logical scope from QC.

    Never invented automatically: ``reason`` and ``expires_on`` are always
    analyst-supplied. An expired exclusion blocks setup until it is renewed
    or removed (Step 7/8 UI behavior); this model only freezes the shape.
    """

    reason: str = Field(min_length=1, max_length=512)
    expires_on: dt.date

    model_config = {"frozen": True}


class PeriodBandContract(BaseModel):
    """A saved cadence axis for one logical sheet or region.

    Composes with, but never overrides, explicit row-key pairing: a period
    band on the orthogonal axis can still align cadence growth for a keyed
    region (see the plan's End-To-End Flows / Acceptance Criteria).
    """

    band_id: str
    axis: PeriodAxis
    cadence_kind: PeriodCadenceKind
    preferred_current_range: str | None = None
    dynamic_expansion: bool = True

    model_config = {"frozen": True}

    @field_validator("band_id")
    @classmethod
    def _validate_band_id(cls, value: str) -> str:
        return _validate_logical_id(value, field_name="band_id")

    @field_validator("preferred_current_range")
    @classmethod
    def _validate_preferred_current_range(cls, value: str | None) -> str | None:
        if value is None:
            return None
        _range_bounds(value)
        return value.upper()


class SelectorPrerequisiteContract(BaseModel):
    """A saved scenario/selector prerequisite.

    Never stores a business value: at run time the two sides' actual saved
    values are compared for exact non-blank equality, and neither value is
    persisted anywhere (session, resolved configuration, coverage, or
    summary).
    """

    selector_id: str
    label: str = Field(min_length=1, max_length=128)
    owner_sheet_id: str
    preferred_current_cell: str

    model_config = {"frozen": True}

    @field_validator("selector_id", "owner_sheet_id")
    @classmethod
    def _validate_ids(cls, value: str) -> str:
        return _validate_logical_id(value, field_name="selector/owner id")

    @field_validator("preferred_current_cell")
    @classmethod
    def _validate_preferred_current_cell(cls, value: str) -> str:
        return _validate_a1_cell(value)


class LogicalColumnContract(BaseModel):
    """One logical column inside a logical region.

    ``alignment_role`` and ``comparison_policy`` are orthogonal dimensions
    represented as two independent enum fields (never two booleans), so
    identity and ordinal are disjoint by construction.

    Scope boundary (disclosed, plan-20260913 Step 8): ``comparison_policy
    == "ignore"`` reaches the value-diff engine (cached-value/number-
    format findings are suppressed); it does not yet reach the formula-
    diff engine, so a formula-text change in an ignored column may still
    surface a finding.
    """

    column_id: str
    alias: str = ""
    alignment_role: AlignmentRole = "none"
    comparison_policy: ColumnComparisonPolicy = "normal"
    #: Optional outer-whitespace trim for an identity column's equality
    #: check. The only normalization allowed in the first delivery; no
    #: case-folding, numeric coercion, or date coercion.
    trim_outer_whitespace: bool = False
    #: Recorded when the analyst confirms a formula-backed identity column
    #: despite the freshness warning (a formula-backed key cannot prove it
    #: reflects the saved file's calculated state).
    formula_backed_identity_acknowledged: bool = False
    exclusion: StructuralExclusionContract | None = Field(
        default=None, exclude_if=lambda value: value is None
    )

    model_config = {"frozen": True}

    @field_validator("column_id")
    @classmethod
    def _validate_column_id(cls, value: str) -> str:
        return _validate_logical_id(value, field_name="column_id")

    @model_validator(mode="after")
    def _ignore_requires_exclusion(self) -> Self:
        if self.comparison_policy == "ignore" and self.exclusion is None:
            raise ValueError(
                "comparison_policy='ignore' requires a structural exclusion "
                "(reason and expiry)"
            )
        if self.comparison_policy != "ignore" and self.exclusion is not None:
            raise ValueError(
                "only comparison_policy='ignore' columns may carry a "
                "structural exclusion"
            )
        return self

    @model_validator(mode="after")
    def _trim_only_meaningful_for_identity(self) -> Self:
        if self.trim_outer_whitespace and self.alignment_role != "identity":
            raise ValueError(
                "trim_outer_whitespace only applies to alignment_role='identity'"
            )
        return self

    @model_validator(mode="after")
    def _formula_ack_only_meaningful_for_identity(self) -> Self:
        if self.formula_backed_identity_acknowledged and self.alignment_role != "identity":
            raise ValueError(
                "formula_backed_identity_acknowledged only applies to "
                "alignment_role='identity'"
            )
        return self


class LogicalRegionContract(BaseModel):
    """One logical region within one logical sheet.

    A region is a saved structural *preference*: an anchor cell (a stable
    point identifying the same logical block across ordinary growth), a
    preferred current-side range, explicit header intent, and a dynamic
    expansion policy. Exact baseline/current ranges, first-data rows, and
    column letters for one run live only in ``ResolvedInputConfigurationV1``.
    """

    region_id: str
    alias: str = ""
    mode: RegionMode = "automatic"
    header_intent: HeaderIntent = "automatic"
    #: Anchor cell inside the region, stable across ordinary row/column growth.
    anchor_cell: str
    #: Preferred current-side range: a starting point for run-time
    #: resolution, not a permanent two-sided profile coordinate.
    preferred_current_range: str | None = None
    #: Only meaningful when ``header_intent == "first_data_row"``.
    preferred_first_data_row: int | None = Field(default=None, ge=1)
    dynamic_expansion: bool = True
    columns: tuple[LogicalColumnContract, ...] = ()
    period_bands: tuple[PeriodBandContract, ...] = ()
    blank_key_policy: BlankKeyPolicy = "system_default"
    duplicate_key_policy: DuplicateKeyPolicy = "skip"
    exclusion: StructuralExclusionContract | None = Field(
        default=None, exclude_if=lambda value: value is None
    )

    model_config = {"frozen": True}

    @field_validator("region_id")
    @classmethod
    def _validate_region_id(cls, value: str) -> str:
        return _validate_logical_id(value, field_name="region_id")

    @field_validator("anchor_cell")
    @classmethod
    def _validate_anchor_cell(cls, value: str) -> str:
        return _validate_a1_cell(value)

    @field_validator("preferred_current_range")
    @classmethod
    def _validate_preferred_current_range(cls, value: str | None) -> str | None:
        if value is None:
            return None
        _range_bounds(value)
        return value.upper()

    @model_validator(mode="after")
    def _excluded_mode_requires_exclusion(self) -> Self:
        if self.mode == "excluded" and self.exclusion is None:
            raise ValueError("mode='excluded' requires a structural exclusion")
        if self.mode != "excluded" and self.exclusion is not None:
            raise ValueError("only mode='excluded' regions may carry an exclusion")
        return self

    @model_validator(mode="after")
    def _header_intent_and_first_data_row_agree(self) -> Self:
        if self.header_intent == "first_data_row" and self.preferred_first_data_row is None:
            raise ValueError(
                "header_intent='first_data_row' requires preferred_first_data_row"
            )
        if self.header_intent != "first_data_row" and self.preferred_first_data_row is not None:
            raise ValueError(
                "preferred_first_data_row only applies to "
                "header_intent='first_data_row'"
            )
        return self

    @model_validator(mode="after")
    def _column_ids_unique(self) -> Self:
        ids = [column.column_id for column in self.columns]
        if len(ids) != len(set(ids)):
            raise ValueError("column_id values must be unique within a region")
        return self

    @model_validator(mode="after")
    def _period_band_ids_unique(self) -> Self:
        ids = [band.band_id for band in self.period_bands]
        if len(ids) != len(set(ids)):
            raise ValueError("band_id values must be unique within a region")
        return self


class LogicalSheetContract(BaseModel):
    """One logical sheet, stable across a physical rename.

    ``preferred_sheet_name`` is a hint used to pre-pair setup analysis on a
    later run; it is never authoritative. The authoritative per-run pairing
    lives in ``ResolvedInputConfigurationV1``.
    """

    sheet_id: str
    alias: str = ""
    preferred_sheet_name: str | None = None
    #: Default region-mode inherited by populated cells with no explicit
    #: region rule. ``"automatic"`` is the un-migrated default; explicit
    #: region rules always override it.
    default_mode: SheetDefaultMode = "automatic"
    regions: tuple[LogicalRegionContract, ...] = ()
    period_bands: tuple[PeriodBandContract, ...] = ()
    selectors: tuple[SelectorPrerequisiteContract, ...] = ()
    exclusion: StructuralExclusionContract | None = Field(
        default=None, exclude_if=lambda value: value is None
    )

    model_config = {"frozen": True}

    @field_validator("sheet_id")
    @classmethod
    def _validate_sheet_id(cls, value: str) -> str:
        return _validate_logical_id(value, field_name="sheet_id")

    @model_validator(mode="after")
    def _excluded_sheet_requires_exclusion(self) -> Self:
        # A sheet has no "mode" field (only regions/columns do), but an
        # attached exclusion still requires a reason/expiry when present;
        # nothing else to cross-check here.
        return self

    @model_validator(mode="after")
    def _region_ids_unique(self) -> Self:
        ids = [region.region_id for region in self.regions]
        if len(ids) != len(set(ids)):
            raise ValueError("region_id values must be unique within a sheet")
        return self

    @model_validator(mode="after")
    def _regions_do_not_overlap(self) -> Self:
        anchors = [region.anchor_cell for region in self.regions]
        if len(anchors) != len(set(anchors)):
            raise ValueError("two regions cannot share the same anchor cell")
        ranged = [
            (region.region_id, region.preferred_current_range)
            for region in self.regions
            if region.preferred_current_range is not None
        ]
        for index, (region_id, cell_range) in enumerate(ranged):
            for other_id, other_range in ranged[index + 1 :]:
                if _ranges_overlap(cell_range, other_range):
                    raise ValueError(
                        f"regions {region_id!r} and {other_id!r} have "
                        "overlapping preferred_current_range values"
                    )
        return self

    @model_validator(mode="after")
    def _period_band_ids_unique(self) -> Self:
        ids = [band.band_id for band in self.period_bands]
        if len(ids) != len(set(ids)):
            raise ValueError("band_id values must be unique within a sheet")
        return self

    @model_validator(mode="after")
    def _selector_ids_unique(self) -> Self:
        ids = [selector.selector_id for selector in self.selectors]
        if len(ids) != len(set(ids)):
            raise ValueError("selector_id values must be unique within a sheet")
        return self


class LogicalMemberContract(BaseModel):
    """One logical Excel package member (matches the existing stable
    ``MEMBER_ID_PATTERN`` member-identity system already used by
    ``qc_tool.package``; members are not renamed the way sheets are, so no
    separate alias-hint field is needed beyond presentation ``alias``).
    """

    member_id: str
    alias: str = ""
    sheets: tuple[LogicalSheetContract, ...] = ()

    model_config = {"frozen": True}

    @field_validator("member_id")
    @classmethod
    def _validate_member_id(cls, value: str) -> str:
        if re.fullmatch(MEMBER_ID_PATTERN, value) is None:
            raise ValueError(f"invalid member id: {value!r}")
        return value

    @model_validator(mode="after")
    def _sheet_ids_unique(self) -> Self:
        ids = [sheet.sheet_id for sheet in self.sheets]
        if len(ids) != len(set(ids)):
            raise ValueError("sheet_id values must be unique within a member")
        return self


class WorkbookInputContract(BaseModel):
    """``InputContractV1``: the saved logical Excel input contract.

    Embedded as ``DeliverableProfile.input_contract``. Absent (``None``)
    means every scope uses legacy physical-name profile fields and
    automatic detection, unchanged -- see ``qc_tool.config.compat`` (Step 2)
    for the compatibility projector.
    """

    version: Literal[1] = 1
    members: tuple[LogicalMemberContract, ...] = ()

    model_config = {"frozen": True}

    @model_validator(mode="after")
    def _member_ids_unique(self) -> Self:
        ids = [member.member_id for member in self.members]
        if len(ids) != len(set(ids)):
            raise ValueError("member_id values must be unique within a contract")
        return self

    def canonical_json(self) -> str:
        payload = self.model_dump(mode="json", by_alias=True)
        return json.dumps(payload, sort_keys=True, separators=(",", ":"))

    def canonical_sha256(self) -> str:
        return hashlib.sha256(self.canonical_json().encode("utf-8")).hexdigest()
