"""Tests for engine-native execution bindings (``qc_tool.config.execution``).

Plan: docs/plans/plan-20260913-mode-aware-configuration-wizard.md, Step 3.
"""

from __future__ import annotations

from qc_tool.config.execution import (
    ExecutionBindings,
    build_execution_bindings,
    region_as_row_identity_rule,
)
from qc_tool.config.resolved_input import (
    ResolvedColumn,
    ResolvedInputConfigurationV1,
    ResolvedMember,
    ResolvedRegion,
    ResolvedSheet,
)


def _member(**overrides: object) -> ResolvedMember:
    defaults: dict[str, object] = {
        "member_id": "primary",
        "baseline_source_sha256": "a" * 64,
        "current_source_sha256": "b" * 64,
    }
    defaults.update(overrides)
    return ResolvedMember(**defaults)  # type: ignore[arg-type]


# --- build_execution_bindings guard ----------------------------------------


def test_build_execution_bindings_is_none_for_none_input() -> None:
    assert build_execution_bindings(None) is None


def test_build_execution_bindings_is_none_for_legacy_default() -> None:
    resolved = ResolvedInputConfigurationV1.legacy_default(profile_name="default")
    assert build_execution_bindings(resolved) is None


def test_build_execution_bindings_returns_bindings_for_a_real_resolution() -> None:
    resolved = ResolvedInputConfigurationV1(members=(_member(),))
    bindings = build_execution_bindings(resolved)
    assert isinstance(bindings, ExecutionBindings)


# --- confirmed_sheet_renames -------------------------------------------------


def test_no_rename_when_baseline_and_current_names_match() -> None:
    resolved = ResolvedInputConfigurationV1(
        members=(
            _member(
                sheets=(
                    ResolvedSheet(
                        sheet_id="data",
                        baseline_sheet_name="Data",
                        current_sheet_name="Data",
                    ),
                )
            ),
        )
    )
    bindings = ExecutionBindings(resolved)
    assert bindings.confirmed_sheet_renames("primary") == {}


def test_a_confirmed_rename_is_reported() -> None:
    resolved = ResolvedInputConfigurationV1(
        members=(
            _member(
                sheets=(
                    ResolvedSheet(
                        sheet_id="data",
                        baseline_sheet_name="2025 Data",
                        current_sheet_name="2026 Data",
                    ),
                )
            ),
        )
    )
    bindings = ExecutionBindings(resolved)
    assert bindings.confirmed_sheet_renames("primary") == {"2026 Data": "2025 Data"}


def test_unresolved_member_returns_no_renames() -> None:
    resolved = ResolvedInputConfigurationV1(members=(_member(),))
    bindings = ExecutionBindings(resolved)
    assert bindings.confirmed_sheet_renames("ops") == {}


def test_a_sheet_missing_one_side_name_is_never_treated_as_a_rename() -> None:
    resolved = ResolvedInputConfigurationV1(
        members=(
            _member(
                sheets=(
                    ResolvedSheet(sheet_id="data", current_sheet_name="Data"),
                )
            ),
        )
    )
    bindings = ExecutionBindings(resolved)
    assert bindings.confirmed_sheet_renames("primary") == {}


# --- region_as_row_identity_rule --------------------------------------------


def _keyed_region(**overrides: object) -> ResolvedRegion:
    defaults: dict[str, object] = {
        "region_id": "r1",
        "mode": "keyed",
        "current_data_range": "A2:C10",
        "columns": (
            ResolvedColumn(column_id="id", current_letter="B", alignment_role="identity"),
            ResolvedColumn(column_id="qty", current_letter="C", alignment_role="ordinal"),
        ),
    }
    defaults.update(overrides)
    return ResolvedRegion(**defaults)  # type: ignore[arg-type]


def test_non_keyed_region_never_produces_a_rule() -> None:
    region = _keyed_region(mode="automatic")
    assert region_as_row_identity_rule(region) is None


def test_keyed_region_with_no_identity_column_produces_no_rule() -> None:
    region = _keyed_region(
        columns=(ResolvedColumn(column_id="qty", current_letter="C", alignment_role="ordinal"),)
    )
    assert region_as_row_identity_rule(region) is None


def test_keyed_region_adapts_into_a_valid_row_identity_rule() -> None:
    region = _keyed_region()
    rule = region_as_row_identity_rule(region)
    assert rule is not None
    assert rule.anchor_cell == "A2"
    assert rule.identity_columns == ["B"]
    assert rule.ordinal_columns == ["C"]
    assert rule.duplicate_policy == "skip"
    assert rule.header_row is None


def test_first_data_row_header_intent_yields_the_row_before_it() -> None:
    region = _keyed_region(header_intent="first_data_row", current_first_data_row=3)
    rule = region_as_row_identity_rule(region)
    assert rule is not None
    assert rule.header_row == 2


def test_no_header_intent_yields_no_header_row_the_same_as_automatic() -> None:
    """Known Step 3 scope boundary: the adapted RowIdentityRule cannot
    distinguish 'no_header' from 'automatic' (a legacy-shape limitation);
    both behave identically today (Step 8 sharpens this further upstream).
    """
    no_header = region_as_row_identity_rule(_keyed_region(header_intent="no_header"))
    automatic = region_as_row_identity_rule(_keyed_region(header_intent="automatic"))
    assert no_header is not None and automatic is not None
    assert no_header.header_row == automatic.header_row is None


def test_duplicate_key_policy_is_carried_through() -> None:
    region = _keyed_region(duplicate_key_policy="occurrence")
    rule = region_as_row_identity_rule(region)
    assert rule is not None
    assert rule.duplicate_policy == "occurrence"


def test_region_with_no_resolvable_anchor_range_produces_no_rule() -> None:
    region = _keyed_region(current_data_range=None, current_outer_range=None)
    assert region_as_row_identity_rule(region) is None


def test_current_preamble_rows_yields_a_header_row_relative_to_the_outer_range() -> None:
    """Step 8: an analyst-entered preamble ROW COUNT (not just an explicit
    first-data-row) also produces a header_row, computed from the region's
    own outer-range top.
    """
    region = _keyed_region(current_outer_range="A1:C10", current_preamble_rows=2)
    rule = region_as_row_identity_rule(region)
    assert rule is not None
    assert rule.header_row == 2  # rows 1-2 preamble, data starts row 3


def test_current_footer_rows_yields_a_footer_row_relative_to_the_outer_range() -> None:
    region = _keyed_region(current_outer_range="A1:C10", current_footer_rows=2)
    rule = region_as_row_identity_rule(region)
    assert rule is not None
    assert rule.footer_row == 9  # rows 9-10 footer


def test_baseline_side_facts_produce_independent_overrides() -> None:
    """Step 8: genuinely different baseline preamble/footer sizes produce
    real baseline_header_row/baseline_footer_row overrides, not the shared
    current-side value.
    """
    region = _keyed_region(
        current_outer_range="A1:C10",
        current_preamble_rows=2,
        current_footer_rows=1,
        baseline_outer_range="A1:C9",
        baseline_preamble_rows=1,
        baseline_footer_rows=1,
    )
    rule = region_as_row_identity_rule(region)
    assert rule is not None
    assert rule.header_row == 2
    assert rule.footer_row == 10
    assert rule.baseline_header_row == 1
    assert rule.baseline_footer_row == 9


def test_absent_baseline_outer_range_leaves_the_per_side_overrides_unset() -> None:
    """When this region has no independently-detected baseline facts at
    all, the overrides stay None so the alignment layer's own fallback
    (share the current-side value) applies -- never a wrong guess.
    """
    region = _keyed_region(current_outer_range="A1:C10", current_preamble_rows=2)
    rule = region_as_row_identity_rule(region)
    assert rule is not None
    assert rule.baseline_header_row is None
    assert rule.baseline_footer_row is None


# --- row_identity_rules lookup ----------------------------------------------


def test_row_identity_rules_lookup_by_member_and_current_sheet_name() -> None:
    resolved = ResolvedInputConfigurationV1(
        members=(
            _member(
                sheets=(
                    ResolvedSheet(
                        sheet_id="data",
                        current_sheet_name="Data",
                        regions=(_keyed_region(),),
                    ),
                )
            ),
        )
    )
    bindings = ExecutionBindings(resolved)
    rules = bindings.row_identity_rules("primary", "Data")
    assert len(rules) == 1
    assert rules[0].identity_columns == ["B"]
    assert bindings.row_identity_rules("primary", "Other") == ()
    assert bindings.row_identity_rules("ops", "Data") == ()


def test_logical_sheet_id_lookup() -> None:
    resolved = ResolvedInputConfigurationV1(
        members=(
            _member(
                sheets=(ResolvedSheet(sheet_id="data", current_sheet_name="Data"),)
            ),
        )
    )
    bindings = ExecutionBindings(resolved)
    assert bindings.logical_sheet_id("primary", "Data") == "data"
    assert bindings.logical_sheet_id("primary", "Other") is None
