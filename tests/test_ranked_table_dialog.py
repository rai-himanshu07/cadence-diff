"""Step 9: pure view model and validation for the ranked-table review dialog.

No NiceGUI import anywhere in this file -- these are ordinary unit tests on
plain dataclasses and functions (plan-20260909, Criteria 12-15).
"""

from __future__ import annotations

from qc_tool.config.profile import DeliverableProfile, RowIdentityRule
from qc_tool.ui.ranked_table_dialog import (
    DUPLICATE_POLICY_COPY,
    DialogViewModel,
    apply_view_model,
    region_draft_from_item,
    region_to_row_identity_rule,
    view_model_from_action,
    with_row_identity_rule,
)

_V2_ITEM: dict[str, object] = {
    "member_id": "primary",
    "sheet": "Panel",
    "cell": "A1",
    "label": "Possible ranked/sorted table: columns B",
    "ranked_table_evidence": {
        "version": 2,
        "member_id": "primary",
        "sheet": "Panel",
        "current_range": "A1:E6001",
        "data_row_count": 6000,
        "header_row": 1,
        "available_columns": ["A", "B", "C", "D", "E"],
        "column_headers": ["Rank", "Record ID", "Value", "Value 2", "Value 3"],
        "suggested_identity_columns": ["B"],
        "suggested_ordinal_columns": ["A"],
        "non_blank_coverage": 0.999,
        "unique_ratio": 0.998,
        "key_overlap": 0.95,
        "formula_ratio": 0.0,
        "displaced_ratio": 1.0,
        "mismatch_reduction": 0.995,
        "projected_positional_mismatches": 500_000,
        "projected_avoided_mismatches": 497_500,
    },
}

_V1_ITEM: dict[str, object] = {
    "member_id": "primary",
    "sheet": "Legacy",
    "cell": "A1",
    "detail": "non_blank_coverage=99.9% unique_ratio=99.8%",
    "suggested_identity_columns": ["B"],
    "suggested_ordinal_columns": ["A"],
}


# --- RegionDraft construction ------------------------------------------------


def test_region_draft_from_v2_item_carries_typed_evidence() -> None:
    region = region_draft_from_item(_V2_ITEM)
    assert region is not None
    assert region.has_typed_evidence
    assert region.member_id == "primary"
    assert region.sheet == "Panel"
    assert region.anchor_cell == "A1"
    assert region.current_range == "A1:E6001"
    assert region.available_columns == ("A", "B", "C", "D", "E")
    assert region.header_row == 1
    assert region.column_headers == (
        "Rank",
        "Record ID",
        "Value",
        "Value 2",
        "Value 3",
    )
    assert region.identity_columns == ("B",)
    assert region.ordinal_columns == ("A",)
    assert region.data_row_count == 6000


def test_region_draft_from_v1_item_is_the_bounded_compatibility_path() -> None:
    region = region_draft_from_item(_V1_ITEM)
    assert region is not None
    assert not region.has_typed_evidence
    assert region.available_columns == ()
    assert region.identity_columns == ("B",)
    assert region.legacy_detail == "non_blank_coverage=99.9% unique_ratio=99.8%"


def test_region_draft_from_item_returns_none_for_non_ranked_items() -> None:
    assert region_draft_from_item({"sheet": "Config", "cell": "B2"}) is None
    assert region_draft_from_item({"sheet": "", "cell": "A1"}) is None
    assert region_draft_from_item({}) is None


# --- RegionDraft display -----------------------------------------------------


def test_region_label_is_member_qualified_only_when_not_primary() -> None:
    primary = region_draft_from_item(_V2_ITEM)
    assert primary is not None
    assert primary.label == "Panel / A1:E6001"

    member_item = dict(_V2_ITEM, member_id="ops")
    member_region = region_draft_from_item(member_item)
    assert member_region is not None
    assert member_region.label == "ops / Panel / A1:E6001"


def test_region_label_omits_range_for_a_legacy_item() -> None:
    region = region_draft_from_item(_V1_ITEM)
    assert region is not None
    assert region.label == "Legacy"


def test_column_options_keep_letters_and_add_detected_headers() -> None:
    region = region_draft_from_item(_V2_ITEM)
    assert region is not None
    assert region.column_options == {
        "A": "A · Rank",
        "B": "B · Record ID",
        "C": "C · Value",
        "D": "D · Value 2",
        "E": "E · Value 3",
    }


def test_noise_summary_and_why_paused_use_typed_evidence_not_telemetry() -> None:
    region = region_draft_from_item(_V2_ITEM)
    assert region is not None
    assert "497,500" in region.noise_summary
    assert "500,000" in region.noise_summary
    assert "non_blank_coverage=" not in region.noise_summary
    detail = region.why_paused_detail
    assert "6,000 data rows" in detail
    assert "100%" in detail  # non-blank coverage rounds to 100% at .0%
    assert "non_blank_coverage=" not in detail


def test_noise_summary_and_why_paused_fall_back_to_legacy_detail() -> None:
    region = region_draft_from_item(_V1_ITEM)
    assert region is not None
    assert region.noise_summary == "non_blank_coverage=99.9% unique_ratio=99.8%"
    assert region.why_paused_detail == "non_blank_coverage=99.9% unique_ratio=99.8%"


# --- RegionDraft validation ---------------------------------------------------


def test_region_errors_require_at_least_one_identity_column() -> None:
    region = region_draft_from_item(_V2_ITEM)
    assert region is not None
    empty = region.with_columns(identity=(), ordinal=())
    assert not empty.is_valid
    assert empty.status == "error"
    assert "Choose at least one column" in empty.errors()[0]


def test_region_errors_reject_identity_ordinal_overlap() -> None:
    region = region_draft_from_item(_V2_ITEM)
    assert region is not None
    overlapping = region.with_columns(identity=("B",), ordinal=("B",))
    assert not overlapping.is_valid
    assert "cannot both match identity" in overlapping.errors()[0]


def test_region_errors_reject_selections_outside_the_region() -> None:
    region = region_draft_from_item(_V2_ITEM)
    assert region is not None
    outside = region.with_columns(identity=("Z",), ordinal=())
    assert not outside.is_valid
    assert "Outside this region" in outside.errors()[0]


def test_region_is_valid_with_a_clean_selection() -> None:
    region = region_draft_from_item(_V2_ITEM)
    assert region is not None
    assert region.is_valid
    assert region.status == "ready"


def test_legacy_region_skips_the_outside_region_check() -> None:
    """A v1 item has no ``available_columns`` to validate against; the
    identity/ordinal-disjoint checks still apply."""
    region = region_draft_from_item(_V1_ITEM)
    assert region is not None
    retyped = region.with_columns(identity=("Q",), ordinal=())
    assert retyped.is_valid


def test_duplicate_policy_copy_covers_all_three_with_skip_recommended() -> None:
    assert set(DUPLICATE_POLICY_COPY) == {"skip", "occurrence", "position"}
    for label, consequence, risk in DUPLICATE_POLICY_COPY.values():
        assert label
        assert consequence
        assert risk
    assert DUPLICATE_POLICY_COPY["skip"][2] == "recommended"


# --- DialogViewModel: multi-region state -------------------------------------


def _view_model() -> DialogViewModel:
    action: dict[str, object] = {
        "items": [
            _V2_ITEM,
            dict(_V2_ITEM, sheet="Panel2", member_id="ops"),
        ]
    }
    view_model = view_model_from_action(
        action,
        source_profile="ops-profile",
        opened_source_existed=True,
    )
    assert view_model is not None
    return view_model


def test_view_model_from_action_builds_one_region_per_ranked_item() -> None:
    view_model = _view_model()
    assert view_model.region_count == 2
    assert view_model.regions[0].sheet == "Panel"
    assert view_model.regions[1].sheet == "Panel2"
    assert view_model.profile_name == "ops-profile"
    assert view_model.opened_profile_name == "ops-profile"


def test_view_model_from_action_returns_none_without_ranked_regions() -> None:
    action: dict[str, object] = {"items": [{"sheet": "Config", "cell": "B2", "detail": "mismatch"}]}
    assert view_model_from_action(action, source_profile="default") is None


def test_view_model_from_action_blanks_destination_when_source_is_default() -> None:
    action: dict[str, object] = {"items": [_V2_ITEM]}
    view_model = view_model_from_action(action, source_profile="default")
    assert view_model is not None
    assert view_model.profile_name == ""
    assert view_model.opened_profile_name == ""
    assert not view_model.persist_profile
    assert view_model.is_valid


def test_unsaved_temporary_source_defaults_to_run_once() -> None:
    view_model = view_model_from_action(
        {"items": [_V2_ITEM]},
        source_profile="default (temporary)",
        source_profile_saved=False,
    )

    assert view_model is not None
    assert view_model.profile_name == ""
    assert not view_model.persist_profile
    assert view_model.is_valid


def test_with_active_index_clamps_to_the_region_bounds() -> None:
    view_model = _view_model()
    assert view_model.with_active_index(5).active_index == 1
    assert view_model.with_active_index(-5).active_index == 0
    assert view_model.with_active_index(1).active_index == 1


def test_with_region_and_with_active_region_update_only_the_target() -> None:
    view_model = _view_model()
    updated_region = view_model.regions[0].with_duplicate_policy("occurrence")
    updated = view_model.with_region(0, updated_region)
    assert updated.regions[0].duplicate_policy == "occurrence"
    assert updated.regions[1].duplicate_policy == "skip"

    moved = updated.with_active_index(1)
    active_updated = moved.with_active_region(
        moved.active_region.with_duplicate_policy("position")
    )
    assert active_updated.regions[1].duplicate_policy == "position"
    assert active_updated.regions[0].duplicate_policy == "occurrence"


def test_first_invalid_index_finds_the_first_invalid_region() -> None:
    view_model = _view_model()
    assert view_model.first_invalid_index is None
    broken = view_model.with_region(1, view_model.regions[1].with_columns(identity=(), ordinal=()))
    assert broken.first_invalid_index == 1


# --- DialogViewModel: destination validation and concurrency -----------------


def test_destination_errors_requires_a_non_default_name() -> None:
    view_model = _view_model()
    assert view_model.with_profile_name("").destination_errors()
    assert view_model.with_profile_name("default").destination_errors()
    assert not view_model.with_profile_name("ops-profile-2").destination_errors()


def test_temporary_mode_never_requires_a_profile_destination() -> None:
    view_model = _view_model().with_persist_profile(False).with_profile_name("")
    assert not view_model.destination_errors()
    assert view_model.is_valid
    assert view_model.initial_focus_target == "region_identity"


def test_saved_mode_requires_a_valid_existing_or_new_profile_name() -> None:
    view_model = view_model_from_action(
        {"items": [_V2_ITEM]}, source_profile="default"
    )
    assert view_model is not None
    saved = view_model.with_persist_profile(True)
    assert saved.destination_errors()
    assert saved.initial_focus_target == "profile_name"
    existing = saved.with_profile_destination(
        "existing",
        opened_hash="abc123",
        opened_source_existed=True,
    )
    assert not existing.destination_errors()
    assert not existing.is_creating_profile


def test_is_valid_combines_destination_and_region_validity() -> None:
    view_model = _view_model()
    assert view_model.is_valid
    assert not view_model.with_profile_name("default").is_valid
    broken = view_model.with_region(0, view_model.regions[0].with_columns(identity=(), ordinal=()))
    assert not broken.is_valid


def test_is_creating_profile_reflects_a_changed_destination() -> None:
    view_model = _view_model()
    assert not view_model.is_creating_profile
    assert view_model.with_profile_name("brand-new").is_creating_profile


def test_initial_focus_target_depends_on_whether_default_was_immutable() -> None:
    from_default = view_model_from_action({"items": [_V2_ITEM]}, source_profile="default")
    assert from_default is not None
    assert from_default.initial_focus_target == "region_identity"

    saving_from_default = from_default.with_persist_profile(True)
    assert saving_from_default.initial_focus_target == "profile_name"

    from_named = view_model_from_action({"items": [_V2_ITEM]}, source_profile="ops")
    assert from_named is not None
    assert from_named.initial_focus_target == "region_identity"


def test_has_profile_conflict_only_applies_to_the_originally_opened_name() -> None:
    view_model = DialogViewModel(
        regions=(),
        profile_name="ops-profile",
        opened_profile_name="ops-profile",
        opened_hash="abc123",
        opened_source_existed=True,
    )
    # Unchanged on disk: no conflict.
    assert not view_model.has_profile_conflict(current_hash="abc123", current_exists=True)
    # Hash differs: conflict.
    assert view_model.has_profile_conflict(current_hash="def456", current_exists=True)
    # Existed before, now missing: conflict.
    assert view_model.has_profile_conflict(current_hash=None, current_exists=False)
    # A retyped destination was never snapshotted: never conflicts.
    retyped = view_model.with_profile_name("a-brand-new-name")
    assert not retyped.has_profile_conflict(current_hash="anything", current_exists=True)


def test_has_profile_conflict_for_a_freshly_created_destination() -> None:
    view_model = DialogViewModel(
        regions=(),
        profile_name="",
        opened_profile_name="",
        opened_hash=None,
        opened_source_existed=False,
    )
    assert not view_model.has_profile_conflict(current_hash=None, current_exists=False)
    named = view_model.with_profile_name("brand-new")
    # "brand-new" was never the opened name ("" was), so no conflict applies
    # even if something unrelated happens to exist at that path already --
    # the create-vs-update flow (target already exists) is handled
    # separately by the save path, not by this concurrency guard.
    assert not named.has_profile_conflict(current_hash="anything", current_exists=True)


# --- profile application ------------------------------------------------------


def test_region_to_row_identity_rule_converts_fields_directly() -> None:
    region = region_draft_from_item(_V2_ITEM)
    assert region is not None
    rule = region_to_row_identity_rule(region)
    assert rule == RowIdentityRule(
        anchor_cell="A1",
        header_row=1,
        identity_columns=["B"],
        ordinal_columns=["A"],
        duplicate_policy="skip",
    )


def test_with_row_identity_rule_targets_single_and_member_profiles() -> None:
    rule = RowIdentityRule(
        anchor_cell="A1",
        identity_columns=["B"],
        ordinal_columns=["A"],
    )
    single = with_row_identity_rule(
        DeliverableProfile(name="single"),
        member_id="primary",
        workbook_count=1,
        sheet="Panel",
        rule=rule,
    )
    assert single.excel.sheets["Panel"].row_identity_rules == [rule]

    package = with_row_identity_rule(
        DeliverableProfile(name="package"),
        member_id="ops",
        workbook_count=2,
        sheet="Panel",
        rule=rule,
    )
    assert package.excel.members["ops"].sheets["Panel"].row_identity_rules == [rule]


def test_apply_view_model_upserts_every_region_into_its_own_member_scope() -> None:
    view_model = _view_model()
    profile = apply_view_model(
        DeliverableProfile(name="ops-profile"), view_model, workbook_count=2
    )
    assert profile.excel.members["primary"].sheets["Panel"].row_identity_rules[
        0
    ].identity_columns == ["B"]
    assert profile.excel.members["ops"].sheets["Panel2"].row_identity_rules[
        0
    ].identity_columns == ["B"]
