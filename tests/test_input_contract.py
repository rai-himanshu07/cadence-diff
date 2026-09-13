"""Frozen contract tests for the saved logical input contract
(``InputContractV1``, ``qc_tool.config.input_contract``).

Plan: docs/plans/plan-20260913-mode-aware-configuration-wizard.md, Step 1.
These tests freeze every accepted architectural decision before any engine
wiring exists (Step 3+). All fixtures are synthetic; no real workbook data.
"""

from __future__ import annotations

import datetime as dt

import pytest
from pydantic import BaseModel, ValidationError

from qc_tool.config.input_contract import (
    INPUT_CONTRACT_VERSION,
    AlignmentRole,
    DuplicateKeyPolicy,
    LogicalColumnContract,
    LogicalMemberContract,
    LogicalRegionContract,
    LogicalSheetContract,
    PeriodAxis,
    PeriodBandContract,
    RegionMode,
    SelectorPrerequisiteContract,
    StructuralExclusionContract,
    WorkbookInputContract,
)
from qc_tool.config.profile import DeliverableProfile, profile_sha256

_TODAY = dt.date(2026, 9, 13)
_FUTURE = dt.date(2027, 1, 1)

#: Regression pin for `DeliverableProfile(name="monthly")`'s canonical hash,
#: reused from `tests/test_package.py`'s own established pin. Adding
#: `input_contract` must never perturb it while the field stays `None`.
_DEFAULT_PROFILE_HASH = (
    "e7259f869ad0381dad15d726fab606b2c7438f9c8cf9fa9332596f75dfa1ace2"
)


# --- logical id validation ---------------------------------------------


def test_logical_ids_reject_uppercase_and_non_snake_case() -> None:
    with pytest.raises(ValidationError):
        LogicalMemberContract(member_id="Primary")  # member_id has its own pattern
    with pytest.raises(ValidationError):
        LogicalSheetContract(sheet_id="Data-Sheet")
    with pytest.raises(ValidationError):
        LogicalRegionContract(region_id="Region 1", anchor_cell="A1")
    with pytest.raises(ValidationError):
        LogicalColumnContract(column_id="Column1")
    with pytest.raises(ValidationError):
        PeriodBandContract(band_id="Band!", axis="rows", cadence_kind="month")
    with pytest.raises(ValidationError):
        SelectorPrerequisiteContract(
            selector_id="Selector",
            label="Scenario",
            owner_sheet_id="data",
            preferred_current_cell="B2",
        )


def test_logical_ids_accept_lowercase_snake_case() -> None:
    region = LogicalRegionContract(region_id="revenue_table", anchor_cell="A1")
    assert region.region_id == "revenue_table"


# --- aliases are optional and presentation-only -------------------------


def test_aliases_default_empty_and_do_not_gate_anything() -> None:
    sheet = LogicalSheetContract(sheet_id="data")
    region = LogicalRegionContract(region_id="r1", anchor_cell="A1")
    column = LogicalColumnContract(column_id="c1")
    member = LogicalMemberContract(member_id="primary")
    assert sheet.alias == region.alias == column.alias == member.alias == ""


def test_alias_rename_changes_contract_digest() -> None:
    base = WorkbookInputContract(
        members=(LogicalMemberContract(member_id="primary", sheets=(
            LogicalSheetContract(sheet_id="data", alias="Data"),
        )),)
    )
    renamed = WorkbookInputContract(
        members=(LogicalMemberContract(member_id="primary", sheets=(
            LogicalSheetContract(sheet_id="data", alias="Ledger"),
        )),)
    )
    assert base.canonical_sha256() != renamed.canonical_sha256()


# --- all four region modes ------------------------------------------------


@pytest.mark.parametrize("mode", ["automatic", "keyed", "positional"])
def test_region_modes_not_requiring_exclusion(mode: RegionMode) -> None:
    region = LogicalRegionContract(region_id="r1", anchor_cell="A1", mode=mode)
    assert region.mode == mode
    assert region.exclusion is None


def test_excluded_region_mode_requires_exclusion() -> None:
    with pytest.raises(ValidationError):
        LogicalRegionContract(region_id="r1", anchor_cell="A1", mode="excluded")
    region = LogicalRegionContract(
        region_id="r1",
        anchor_cell="A1",
        mode="excluded",
        exclusion=StructuralExclusionContract(reason="pilot only", expires_on=_FUTURE),
    )
    assert region.mode == "excluded"
    assert region.exclusion is not None


def test_non_excluded_region_rejects_attached_exclusion() -> None:
    with pytest.raises(ValidationError):
        LogicalRegionContract(
            region_id="r1",
            anchor_cell="A1",
            mode="automatic",
            exclusion=StructuralExclusionContract(reason="n/a", expires_on=_FUTURE),
        )


# --- explicit header intent -------------------------------------------


def test_header_intent_automatic_and_no_header_reject_first_data_row() -> None:
    LogicalRegionContract(region_id="r1", anchor_cell="A1", header_intent="automatic")
    LogicalRegionContract(region_id="r1", anchor_cell="A1", header_intent="no_header")
    with pytest.raises(ValidationError):
        LogicalRegionContract(
            region_id="r1",
            anchor_cell="A1",
            header_intent="automatic",
            preferred_first_data_row=3,
        )


def test_header_intent_first_data_row_requires_the_row_number() -> None:
    with pytest.raises(ValidationError):
        LogicalRegionContract(
            region_id="r1", anchor_cell="A1", header_intent="first_data_row"
        )
    region = LogicalRegionContract(
        region_id="r1",
        anchor_cell="A1",
        header_intent="first_data_row",
        preferred_first_data_row=3,
    )
    assert region.preferred_first_data_row == 3


def test_no_header_is_a_distinct_state_from_a_missing_header_row_number() -> None:
    """`no_header` must not be encoded as `header_row is None` -- it is its
    own explicit literal, distinguishable from `automatic`.
    """
    no_header = LogicalRegionContract(
        region_id="r1", anchor_cell="A1", header_intent="no_header"
    )
    automatic = LogicalRegionContract(
        region_id="r1", anchor_cell="A1", header_intent="automatic"
    )
    assert no_header.header_intent != automatic.header_intent
    assert no_header.model_dump()["header_intent"] == "no_header"


# --- side-specific resolutions live only in the resolved model ---------


def test_saved_region_contract_never_stores_baseline_current_split() -> None:
    """Exact per-side ranges/first-data-rows are run resolutions, never a
    permanent two-sided profile coordinate.
    """
    fields = set(LogicalRegionContract.model_fields)
    forbidden_fields = (
        "baseline_range",
        "current_range",
        "baseline_first_data_row",
        "current_first_data_row",
    )
    for forbidden in forbidden_fields:
        assert forbidden not in fields
    assert "preferred_current_range" in fields
    assert "preferred_first_data_row" in fields


# --- column alignment role / comparison policy orthogonality -----------


@pytest.mark.parametrize("role", ["none", "identity", "ordinal"])
def test_alignment_role_and_comparison_policy_are_independent_fields(role: AlignmentRole) -> None:
    column = LogicalColumnContract(column_id="c1", alignment_role=role)
    assert column.alignment_role == role
    assert column.comparison_policy == "normal"


def test_identity_and_ordinal_are_disjoint_by_a_single_enum_field() -> None:
    """A column cannot be both identity and ordinal because alignment_role
    is one Literal field, not two independent booleans.
    """
    fields = LogicalColumnContract.model_fields
    assert "alignment_role" in fields
    assert "is_identity" not in fields
    assert "is_ordinal" not in fields


def test_ignore_policy_requires_exclusion_reason_and_expiry() -> None:
    with pytest.raises(ValidationError):
        LogicalColumnContract(column_id="c1", comparison_policy="ignore")
    column = LogicalColumnContract(
        column_id="c1",
        comparison_policy="ignore",
        exclusion=StructuralExclusionContract(reason="legacy helper column", expires_on=_FUTURE),
    )
    assert column.comparison_policy == "ignore"


def test_non_ignore_policy_rejects_attached_exclusion() -> None:
    with pytest.raises(ValidationError):
        LogicalColumnContract(
            column_id="c1",
            comparison_policy="normal",
            exclusion=StructuralExclusionContract(reason="n/a", expires_on=_FUTURE),
        )


def test_expected_refresh_and_normal_are_valid_comparison_policies() -> None:
    LogicalColumnContract(column_id="c1", comparison_policy="normal")
    LogicalColumnContract(column_id="c1", comparison_policy="expected_refresh")


# --- exact/trim key semantics --------------------------------------------


def test_trim_outer_whitespace_only_applies_to_identity_columns() -> None:
    with pytest.raises(ValidationError):
        LogicalColumnContract(
            column_id="c1", alignment_role="none", trim_outer_whitespace=True
        )
    with pytest.raises(ValidationError):
        LogicalColumnContract(
            column_id="c1", alignment_role="ordinal", trim_outer_whitespace=True
        )
    identity = LogicalColumnContract(
        column_id="c1", alignment_role="identity", trim_outer_whitespace=True
    )
    assert identity.trim_outer_whitespace is True


def test_identity_key_equality_defaults_to_exact_no_trim() -> None:
    identity = LogicalColumnContract(column_id="c1", alignment_role="identity")
    assert identity.trim_outer_whitespace is False


def test_formula_backed_identity_acknowledgement_only_applies_to_identity() -> None:
    with pytest.raises(ValidationError):
        LogicalColumnContract(
            column_id="c1",
            alignment_role="none",
            formula_backed_identity_acknowledged=True,
        )
    identity = LogicalColumnContract(
        column_id="c1",
        alignment_role="identity",
        formula_backed_identity_acknowledged=True,
    )
    assert identity.formula_backed_identity_acknowledged is True


# --- blank-key / duplicate-key policy -------------------------------------


def test_blank_key_policy_defaults_to_versioned_system_default() -> None:
    region = LogicalRegionContract(region_id="r1", anchor_cell="A1")
    assert region.blank_key_policy == "system_default"


def test_duplicate_key_policy_defaults_to_skip() -> None:
    region = LogicalRegionContract(region_id="r1", anchor_cell="A1")
    assert region.duplicate_key_policy == "skip"


@pytest.mark.parametrize("policy", ["skip", "occurrence", "position"])
def test_duplicate_key_policy_accepts_all_three_advanced_values(policy: DuplicateKeyPolicy) -> None:
    region = LogicalRegionContract(
        region_id="r1", anchor_cell="A1", duplicate_key_policy=policy
    )
    assert region.duplicate_key_policy == policy


# --- period-axis composition ----------------------------------------------


@pytest.mark.parametrize("axis", ["rows", "columns"])
def test_period_band_axis_accepts_rows_and_columns(axis: PeriodAxis) -> None:
    band = PeriodBandContract(band_id="monthly", axis=axis, cadence_kind="month")
    assert band.axis == axis


def test_period_band_ids_must_be_unique_within_a_region() -> None:
    band = PeriodBandContract(band_id="monthly", axis="columns", cadence_kind="month")
    with pytest.raises(ValidationError):
        LogicalRegionContract(
            region_id="r1", anchor_cell="A1", period_bands=(band, band)
        )


def test_a_region_may_carry_both_a_keyed_row_axis_and_a_period_column_axis() -> None:
    """Composing a keyed region with an orthogonal period band is valid --
    period bands do not require positional mode.
    """
    region = LogicalRegionContract(
        region_id="r1",
        anchor_cell="A1",
        mode="keyed",
        columns=(LogicalColumnContract(column_id="id", alignment_role="identity"),),
        period_bands=(
            PeriodBandContract(band_id="monthly", axis="columns", cadence_kind="month"),
        ),
    )
    assert region.mode == "keyed"
    assert region.period_bands[0].axis == "columns"


# --- structural exclusion reason/expiry -----------------------------------


def test_structural_exclusion_requires_reason_and_expiry() -> None:
    with pytest.raises(ValidationError):
        StructuralExclusionContract(reason="", expires_on=_FUTURE)
    with pytest.raises(ValidationError):
        StructuralExclusionContract.model_validate({"reason": "vendor pilot"})
    exclusion = StructuralExclusionContract(reason="vendor pilot", expires_on=_FUTURE)
    assert exclusion.reason == "vendor pilot"
    assert exclusion.expires_on == _FUTURE


def test_structural_exclusion_never_invents_an_expiry() -> None:
    """No default/auto-generated expiry exists; omitting it is a validation
    error, not a silently-chosen date.
    """
    with pytest.raises(ValidationError):
        StructuralExclusionContract.model_validate({"reason": "vendor pilot"})


# --- one-to-one member/sheet identity within the saved contract ---------


def test_member_ids_must_be_unique_within_a_contract() -> None:
    member = LogicalMemberContract(member_id="primary")
    with pytest.raises(ValidationError):
        WorkbookInputContract(members=(member, member))


def test_sheet_ids_must_be_unique_within_a_member() -> None:
    sheet = LogicalSheetContract(sheet_id="data")
    with pytest.raises(ValidationError):
        LogicalMemberContract(member_id="primary", sheets=(sheet, sheet))


def test_region_ids_must_be_unique_within_a_sheet() -> None:
    region = LogicalRegionContract(region_id="r1", anchor_cell="A1")
    other = LogicalRegionContract(region_id="r1", anchor_cell="C1")
    with pytest.raises(ValidationError):
        LogicalSheetContract(sheet_id="data", regions=(region, other))


def test_column_ids_must_be_unique_within_a_region() -> None:
    column = LogicalColumnContract(column_id="c1")
    with pytest.raises(ValidationError):
        LogicalRegionContract(
            region_id="r1", anchor_cell="A1", columns=(column, column)
        )


# --- region-level granularity / overlap rejection -------------------------


def test_two_regions_cannot_share_an_anchor_cell() -> None:
    with pytest.raises(ValidationError):
        LogicalSheetContract(
            sheet_id="data",
            regions=(
                LogicalRegionContract(region_id="r1", anchor_cell="A1"),
                LogicalRegionContract(region_id="r2", anchor_cell="A1"),
            ),
        )


def test_two_regions_with_overlapping_preferred_ranges_are_rejected() -> None:
    with pytest.raises(ValidationError):
        LogicalSheetContract(
            sheet_id="data",
            regions=(
                LogicalRegionContract(
                    region_id="r1", anchor_cell="A1", preferred_current_range="A1:C10"
                ),
                LogicalRegionContract(
                    region_id="r2", anchor_cell="B5", preferred_current_range="B5:D15"
                ),
            ),
        )


def test_two_regions_with_disjoint_preferred_ranges_are_accepted() -> None:
    sheet = LogicalSheetContract(
        sheet_id="data",
        regions=(
            LogicalRegionContract(
                region_id="r1", anchor_cell="A1", preferred_current_range="A1:C10"
            ),
            LogicalRegionContract(
                region_id="r2", anchor_cell="E1", preferred_current_range="E1:G10"
            ),
        ),
    )
    assert len(sheet.regions) == 2


# --- canonical digests -----------------------------------------------------


def test_canonical_digest_is_deterministic_and_content_sensitive() -> None:
    contract = WorkbookInputContract(
        members=(LogicalMemberContract(member_id="primary", sheets=(
            LogicalSheetContract(sheet_id="data"),
        )),)
    )
    again = WorkbookInputContract(
        members=(LogicalMemberContract(member_id="primary", sheets=(
            LogicalSheetContract(sheet_id="data"),
        )),)
    )
    changed = WorkbookInputContract(
        members=(LogicalMemberContract(member_id="primary", sheets=(
            LogicalSheetContract(sheet_id="other"),
        )),)
    )
    assert contract.canonical_sha256() == again.canonical_sha256()
    assert contract.canonical_sha256() != changed.canonical_sha256()
    assert len(contract.canonical_sha256()) == 64


def test_contract_version_is_frozen_at_one() -> None:
    assert INPUT_CONTRACT_VERSION == 1
    assert WorkbookInputContract().version == 1


# --- absent-contract legacy defaults --------------------------------------


def test_profile_with_no_input_contract_preserves_the_pinned_canonical_hash() -> None:
    implicit = DeliverableProfile(name="monthly")
    explicit_none = DeliverableProfile(name="monthly", input_contract=None)
    assert profile_sha256(implicit) == profile_sha256(explicit_none) == _DEFAULT_PROFILE_HASH


def test_profile_with_an_input_contract_gets_a_different_canonical_hash() -> None:
    profile = DeliverableProfile(
        name="monthly",
        input_contract=WorkbookInputContract(
            members=(LogicalMemberContract(member_id="primary"),)
        ),
    )
    assert profile_sha256(profile) != _DEFAULT_PROFILE_HASH


# --- simple profile lifecycle: no version/draft/verified/promotion -----


def _all_input_contract_models() -> list[type[BaseModel]]:
    return [
        WorkbookInputContract,
        LogicalMemberContract,
        LogicalSheetContract,
        LogicalRegionContract,
        LogicalColumnContract,
        SelectorPrerequisiteContract,
        PeriodBandContract,
        StructuralExclusionContract,
    ]


@pytest.mark.parametrize("model_cls", _all_input_contract_models())
def test_no_saved_contract_model_carries_a_lifecycle_status_field(
    model_cls: type[BaseModel],
) -> None:
    """Profiles stay simple save/reuse/update; the saved contract must never
    grow a draft/verified/executable/promoted/status field.
    """
    forbidden_substrings = ("status", "draft", "verified", "executable", "promot")
    for field_name in model_cls.model_fields:
        lowered = field_name.lower()
        for token in forbidden_substrings:
            assert token not in lowered, f"{model_cls.__name__}.{field_name}"
