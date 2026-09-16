"""Per-run resolved input configuration (``ResolvedInputConfigurationV1``).

Separate and independently versioned/digested from the saved
``InputContractV1`` (``qc_tool.config.input_contract``). This is the exact
configuration one run actually used after applying files, setup-analysis
observations, analyst confirmations, explicit overrides, and saved-contract
preferences. It is an engine execution input (Step 3 makes logical bindings
native to alignment/structural/formula comparison) and is persisted in
history, reports, JSON, HTML, Excel, and attestation with its own canonical
digest -- never sharing ``profile_sha256``'s hash space.

Never stores: filesystem paths, preview values, revealed formula text, or
selector values. Source identity is bound by content hash only.
"""

from __future__ import annotations

import hashlib
import json
from typing import Literal

from pydantic import BaseModel, Field, model_validator

from qc_tool.coverage import QCRunMode
from qc_tool.package import MEMBER_ID_PATTERN

#: Schema version of the resolved per-run configuration payload.
RESOLVED_INPUT_CONFIGURATION_VERSION = 1

ConfigurationCoverageState = Literal[
    "confirmed",
    "inherited",
    "automatic_confirmed",
    "positional",
    "excluded",
    "degraded_acknowledged",
]

#: Where a resolved configuration came from. ``legacy_default`` means the
#: profile carried no ``input_contract``: legacy physical-name fields and
#: automatic detection drove the run, and there is no logical resolution to
#: record beyond this placeholder.
ResolutionSource = Literal["input_contract", "legacy_default"]


class StaleResolvedConfigurationError(ValueError):
    """A resolved configuration no longer matches the sources it names."""


class ResolvedSelector(BaseModel):
    """One selector's exact per-run resolution. Never the actual value."""

    selector_id: str
    baseline_cell: str | None = None
    current_cell: str | None = None
    equal: bool = False
    baseline_formula_backed: bool = False
    current_formula_backed: bool = False
    coverage: ConfigurationCoverageState = "confirmed"

    model_config = {"frozen": True}


class ResolvedColumn(BaseModel):
    """One logical column's exact per-side physical resolution."""

    column_id: str
    baseline_letter: str | None = None
    current_letter: str | None = None
    alignment_role: Literal["none", "identity", "ordinal"] = "none"
    comparison_policy: Literal["normal", "ignore", "expected_refresh"] = "normal"
    #: Optional outer-whitespace trim for an identity column's equality
    #: check (mirrors ``LogicalColumnContract.trim_outer_whitespace``); only
    #: meaningful when ``alignment_role == "identity"``.
    trim_outer_whitespace: bool = False
    coverage: ConfigurationCoverageState = "automatic_confirmed"

    model_config = {"frozen": True}


class ResolvedRegion(BaseModel):
    """One logical region's exact per-run, per-side boundary resolution.

    Outer/data/preamble/footer bounds are all per-side facts: a region's
    notes/header rows above the data and footer rows below it remain
    visible in separate positional comparisons even when the data itself is
    keyed.
    """

    region_id: str
    mode: Literal["automatic", "keyed", "positional", "excluded"] = "automatic"
    header_intent: Literal["automatic", "no_header", "first_data_row"] = "automatic"
    baseline_outer_range: str | None = None
    current_outer_range: str | None = None
    baseline_data_range: str | None = None
    current_data_range: str | None = None
    baseline_first_data_row: int | None = None
    current_first_data_row: int | None = None
    baseline_preamble_rows: int = Field(default=0, ge=0)
    current_preamble_rows: int = Field(default=0, ge=0)
    baseline_footer_rows: int = Field(default=0, ge=0)
    current_footer_rows: int = Field(default=0, ge=0)
    columns: tuple[ResolvedColumn, ...] = ()
    #: Resolved duplicate-key policy for a keyed region (mirrors the saved
    #: `LogicalRegionContract.duplicate_key_policy`); irrelevant otherwise.
    duplicate_key_policy: Literal["skip", "occurrence", "position"] = "skip"
    #: Resolved blank-identity-key policy for a keyed region (mirrors the
    #: saved `LogicalRegionContract.blank_key_policy`); irrelevant otherwise.
    #: ``"system_default"``/``"tolerate"`` both mean today's existing
    #: behavior (a blank-keyed row is excluded from key matching, surfacing
    #: as an ordinary insert/delete); only ``"block"`` changes anything --
    #: it refuses the run before alignment instead of silently excluding.
    blank_key_policy: Literal["system_default", "tolerate", "block"] = "system_default"
    coverage: ConfigurationCoverageState = "automatic_confirmed"
    degraded_reason: str = ""

    model_config = {"frozen": True}

    @model_validator(mode="after")
    def _column_ids_unique(self) -> ResolvedRegion:
        ids = [column.column_id for column in self.columns]
        if len(ids) != len(set(ids)):
            raise ValueError("column_id values must be unique within a resolved region")
        return self


class ResolvedSheet(BaseModel):
    """One logical sheet's exact per-run baseline/current physical pairing."""

    sheet_id: str
    baseline_sheet_name: str | None = None
    current_sheet_name: str | None = None
    regions: tuple[ResolvedRegion, ...] = ()
    selectors: tuple[ResolvedSelector, ...] = ()
    coverage: ConfigurationCoverageState = "automatic_confirmed"

    model_config = {"frozen": True}

    @model_validator(mode="after")
    def _region_ids_unique(self) -> ResolvedSheet:
        ids = [region.region_id for region in self.regions]
        if len(ids) != len(set(ids)):
            raise ValueError("region_id values must be unique within a resolved sheet")
        return self


class ResolvedMember(BaseModel):
    """One logical package member's exact per-run source-hash-bound identity.

    Filesystem paths are deliberately absent; only content hashes bind
    identity, matching ``PackageManifest``'s existing path-free convention.
    """

    member_id: str = Field(pattern=MEMBER_ID_PATTERN)
    baseline_source_sha256: str | None = None
    current_source_sha256: str | None = None
    sheets: tuple[ResolvedSheet, ...] = ()

    model_config = {"frozen": True}

    @model_validator(mode="after")
    def _sheet_ids_unique(self) -> ResolvedMember:
        ids = [sheet.sheet_id for sheet in self.sheets]
        if len(ids) != len(set(ids)):
            raise ValueError("sheet_id values must be unique within a resolved member")
        return self

    @model_validator(mode="after")
    def _sheet_names_are_one_to_one(self) -> ResolvedMember:
        baseline_names = [
            sheet.baseline_sheet_name for sheet in self.sheets if sheet.baseline_sheet_name
        ]
        if len(baseline_names) != len(set(baseline_names)):
            raise ValueError(
                f"member {self.member_id!r}: baseline sheet names must resolve "
                "one-to-one"
            )
        current_names = [
            sheet.current_sheet_name for sheet in self.sheets if sheet.current_sheet_name
        ]
        if len(current_names) != len(set(current_names)):
            raise ValueError(
                f"member {self.member_id!r}: current sheet names must resolve "
                "one-to-one"
            )
        return self


class ResolvedInputConfigurationV1(BaseModel):
    """The exact configuration one run actually used.

    Independent of ``DeliverableProfile``/``profile_sha256``: this model has
    its own canonical JSON form and digest (``canonical_sha256``). A profile
    edit after a run never changes an already-recorded resolved
    configuration or its digest.
    """

    version: Literal[1] = 1
    source: ResolutionSource = "input_contract"
    mode: QCRunMode | None = None
    profile_name: str = ""
    profile_sha256: str = ""
    contract_id: str = ""
    #: Which ``INPUT_CONTRACT_VERSION`` the setup pipeline understood when
    #: this resolution was produced; used to reject a stale resolution
    #: before analysis (see ``validate_freshness``).
    inspection_contract_version: int = 1
    members: tuple[ResolvedMember, ...] = ()
    #: Free-form stable warning codes the analyst explicitly acknowledged
    #: (e.g. formula-backed key, low key overlap); never prose containing
    #: workbook content.
    warnings_acknowledged: tuple[str, ...] = ()
    degradations_accepted: tuple[str, ...] = ()
    exclusions_renewed: tuple[str, ...] = ()
    override_reasons: tuple[str, ...] = ()

    model_config = {"frozen": True}

    @model_validator(mode="after")
    def _member_ids_unique(self) -> ResolvedInputConfigurationV1:
        ids = [member.member_id for member in self.members]
        if len(ids) != len(set(ids)):
            raise ValueError("member_id values must be unique within a resolved configuration")
        return self

    @classmethod
    def legacy_default(
        cls,
        *,
        profile_name: str = "",
        profile_sha256: str = "",
        mode: QCRunMode | None = None,
    ) -> ResolvedInputConfigurationV1:
        """The placeholder resolution for a run whose profile carries no
        ``input_contract``: legacy physical-name fields and automatic
        detection drove execution, unchanged, and there is no logical
        resolution to record.
        """
        return cls(
            source="legacy_default",
            mode=mode,
            profile_name=profile_name,
            profile_sha256=profile_sha256,
        )

    def canonical_json(self) -> str:
        payload = self.model_dump(mode="json", by_alias=True)
        return json.dumps(payload, sort_keys=True, separators=(",", ":"))

    def canonical_sha256(self) -> str:
        return hashlib.sha256(self.canonical_json().encode("utf-8")).hexdigest()


def validate_freshness(
    resolved: ResolvedInputConfigurationV1,
    *,
    current_source_sha256: dict[str, tuple[str | None, str | None]],
) -> None:
    """Refuse a resolved configuration whose sources or schema drifted.

    ``current_source_sha256`` maps ``member_id`` to a freshly computed
    ``(baseline_sha256, current_sha256)`` pair. Raises
    ``StaleResolvedConfigurationError`` on any mismatch or schema-version
    drift; does not attempt to repair or partially apply a stale
    resolution.
    """
    from qc_tool.config.input_contract import INPUT_CONTRACT_VERSION

    if resolved.inspection_contract_version != INPUT_CONTRACT_VERSION:
        raise StaleResolvedConfigurationError(
            "resolved configuration was produced under input-contract "
            f"version {resolved.inspection_contract_version}, but the "
            f"current version is {INPUT_CONTRACT_VERSION}"
        )
    resolved_member_ids = {member.member_id for member in resolved.members}
    missing_member_ids = set(current_source_sha256) - resolved_member_ids
    if missing_member_ids:
        member_id = sorted(missing_member_ids)[0]
        raise StaleResolvedConfigurationError(
            f"resolved configuration is missing member {member_id!r}"
        )
    for member in resolved.members:
        fresh = current_source_sha256.get(member.member_id)
        if fresh is None:
            raise StaleResolvedConfigurationError(
                f"member {member.member_id!r} is no longer part of this run"
            )
        fresh_baseline, fresh_current = fresh
        if member.baseline_source_sha256 != fresh_baseline:
            raise StaleResolvedConfigurationError(
                f"member {member.member_id!r}: baseline source changed since "
                "this configuration was resolved"
            )
        if member.current_source_sha256 != fresh_current:
            raise StaleResolvedConfigurationError(
                f"member {member.member_id!r}: current source changed since "
                "this configuration was resolved"
            )
