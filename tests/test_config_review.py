"""Pure view-model/diff/resolved-configuration builder tests for the
mode-aware configuration workspace shell (plan-20260913, Step 7).
"""

from __future__ import annotations

from dataclasses import replace

from qc_tool.config.profile import DeliverableProfile, ExcelMemberProfile, ExcelProfile
from qc_tool.coverage import QCRunMode
from qc_tool.excel.regions import TableRegion
from qc_tool.setup.models import (
    DetectedRegion,
    MemberSetupProfile,
    SetupAnalysisResult,
    SheetSetupProfile,
    XlsbRiskProfile,
)
from qc_tool.ui.config_review import (
    ConfigWorkspaceState,
    apply_anchor_click,
    apply_manual_range,
    apply_region_transform,
    build_input_contract,
    build_resolved_configuration,
    compute_warnings,
    confirm_all_regions,
    diff_profile_against_scan,
    format_a1_range,
    is_clean_profile_diff,
    member_review_from_scan,
    parse_a1_cell,
    parse_a1_range,
    sheet_pairing,
    slugify,
    unresolved_blockers,
    update_region_decision,
)


def _region(sheet: str, *, min_row=1, min_col=1, max_row=5, max_col=3) -> DetectedRegion:
    return DetectedRegion(
        region=TableRegion(
            sheet=sheet,
            min_row=min_row,
            min_col=min_col,
            max_row=max_row,
            max_col=max_col,
            orientation="long",
            header_row=min_row,
            key_col=min_col,
            period_axis="none",
        )
    )


def _member_profile(*, with_ranked_candidate: bool = False) -> MemberSetupProfile:
    from qc_tool.excel.ranked_identity import RankedTableCandidate

    region = _region("Data")
    if with_ranked_candidate:
        region = DetectedRegion(
            region=region.region,
            ranked_candidate=RankedTableCandidate(
                columns=(1,),
                non_blank_coverage=0.99,
                unique_ratio=0.99,
                key_overlap=0.95,
                formula_ratio=0.0,
                displaced_ratio=0.5,
                mismatch_reduction=0.9,
                projected_positional_mismatches=100,
            ),
        )
    return MemberSetupProfile(
        member_id="primary",
        baseline_hash="a" * 64,
        current_hash="b" * 64,
        current_sheets=(
            SheetSetupProfile(sheet_name="Data", regions=(region,)),
            SheetSetupProfile(sheet_name="NewSheet", regions=()),
        ),
        baseline_sheets=(
            SheetSetupProfile(sheet_name="Data", regions=()),
            SheetSetupProfile(sheet_name="OldSheet", regions=()),
        ),
    )


def _scan_result(**kwargs) -> SetupAnalysisResult:
    return SetupAnalysisResult(members={"primary": _member_profile(**kwargs)})


def test_slugify_produces_a_stable_snake_case_id() -> None:
    assert slugify("Sheet One!A1:B5") == "sheet_one_a1_b5"
    assert slugify("123") == "s_123"
    assert slugify("") == "s"


def test_member_review_from_scan_defaults_every_region_to_automatic() -> None:
    review = member_review_from_scan("primary", _member_profile())
    data_sheet = next(s for s in review.current_sheets if s.sheet_name == "Data")
    assert len(data_sheet.regions) == 1
    assert data_sheet.regions[0].mode == "automatic"
    assert data_sheet.regions[0].confirmed is False
    assert data_sheet.regions[0].anchor_cell == "A1"
    assert data_sheet.regions[0].available_columns == ("A", "B", "C")


def test_sheet_pairing_auto_pairs_same_name_and_flags_added_removed() -> None:
    review = member_review_from_scan("primary", _member_profile())
    pairs, added, removed = sheet_pairing(review)
    assert pairs == (("Data", "Data"),)
    assert added == ("NewSheet",)
    assert removed == ("OldSheet",)


def test_compute_warnings_flags_a_pending_ranked_candidate() -> None:
    result = _scan_result(with_ranked_candidate=True)
    warnings = compute_warnings(result)
    codes = [w.code for w in warnings]
    assert any(code.startswith("ranked_candidate:") for code in codes)


def test_compute_warnings_flags_an_unsafe_xlsb_source() -> None:
    profile = _member_profile()
    profile = MemberSetupProfile(
        member_id="primary",
        baseline_hash=profile.baseline_hash,
        current_hash=profile.current_hash,
        current_sheets=profile.current_sheets,
        baseline_sheets=profile.baseline_sheets,
        current_xlsb_risk=XlsbRiskProfile(safe_for_external_engine=False),
    )
    result = SetupAnalysisResult(members={"primary": profile})
    warnings = compute_warnings(result)
    assert any(w.code == "xlsb_risk:primary" for w in warnings)


def test_compute_warnings_surfaces_a_blocking_scan_failure() -> None:
    result = SetupAnalysisResult(failure_detail="could not read the shared password")
    warnings = compute_warnings(result)
    assert any(w.severity == "block" for w in warnings)


def test_diff_profile_against_scan_classifies_every_scope() -> None:
    result = _scan_result()
    profile = DeliverableProfile(
        name="default",
        excel=ExcelProfile(ignore_sheets=["NewSheet"]),
    )
    diff = diff_profile_against_scan(profile, result, mode=QCRunMode.CYCLE_COMPARISON)
    kinds = {entry.scope: entry.kind for entry in diff}
    assert kinds["primary/NewSheet"] == "conflict"
    assert kinds["primary/Data"] == "new_in_scan"


def test_is_clean_profile_diff_true_only_when_everything_matches() -> None:
    result = _scan_result()
    matching_profile = DeliverableProfile(name="default")
    diff = diff_profile_against_scan(
        matching_profile, result, mode=QCRunMode.CYCLE_COMPARISON
    )
    assert not is_clean_profile_diff(diff)  # nothing saved yet -> new_in_scan
    empty_result = SetupAnalysisResult(members={})
    assert is_clean_profile_diff(
        diff_profile_against_scan(
            matching_profile, empty_result, mode=QCRunMode.CYCLE_COMPARISON
        )
    )


def test_unresolved_blockers_flags_an_invalid_keyed_region_without_identity_columns() -> None:
    review = member_review_from_scan("primary", _member_profile())
    updated = update_region_decision(
        review, "Data", review.current_sheets[0].regions[0].region_id, mode="keyed"
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    blockers = unresolved_blockers(state, ())
    assert blockers


def test_unresolved_blockers_clears_once_identity_columns_are_set() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review, "Data", region_id, mode="keyed", identity_columns=("A",), confirmed=True
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    assert unresolved_blockers(state, ()) == ()


def test_build_resolved_configuration_carries_keyed_identity_columns() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review, "Data", region_id, mode="keyed", identity_columns=("A",), confirmed=True
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    resolved = build_resolved_configuration(
        state, profile=DeliverableProfile(name="default"), profile_sha256="deadbeef"
    )
    assert resolved.source == "input_contract"
    sheet = resolved.members[0].sheets[0]
    assert sheet.baseline_sheet_name == "Data"  # same-name auto-pair
    region = sheet.regions[0]
    assert region.mode == "keyed"
    identity_letters = [c.current_letter for c in region.columns if c.alignment_role == "identity"]
    assert identity_letters == ["A"]


def test_build_resolved_configuration_leaves_untouched_regions_automatic() -> None:
    review = member_review_from_scan("primary", _member_profile())
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(review,))
    resolved = build_resolved_configuration(
        state, profile=DeliverableProfile(name="default"), profile_sha256="deadbeef"
    )
    region = resolved.members[0].sheets[0].regions[0]
    assert region.mode == "automatic"
    assert region.columns == ()


def test_build_input_contract_only_saves_touched_regions() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review, "Data", region_id, mode="keyed", identity_columns=("A",), confirmed=True
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    contract = build_input_contract(state)
    member = contract.members[0]
    data_sheet = next(s for s in member.sheets if s.preferred_sheet_name == "Data")
    assert len(data_sheet.regions) == 1
    assert data_sheet.regions[0].mode == "keyed"
    new_sheet = next(s for s in member.sheets if s.preferred_sheet_name == "NewSheet")
    assert new_sheet.regions == ()


def test_build_input_contract_requires_exclusion_reason_and_expiry_for_excluded_mode() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review,
        "Data",
        region_id,
        mode="excluded",
        exclusion_reason="not part of this cycle",
        exclusion_expires_on="2027-01-01",
        confirmed=True,
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    contract = build_input_contract(state)
    data_sheet = next(
        s for s in contract.members[0].sheets if s.preferred_sheet_name == "Data"
    )
    assert data_sheet.regions[0].exclusion is not None
    assert data_sheet.regions[0].exclusion.reason == "not part of this cycle"


def test_member_review_isolates_a_sheet_failure_without_dropping_other_sheets() -> None:
    profile = MemberSetupProfile(
        member_id="primary",
        baseline_hash="a" * 64,
        current_hash="b" * 64,
        current_sheets=(
            SheetSetupProfile(sheet_name="Broken", failure_detail="unexpected shape"),
            SheetSetupProfile(sheet_name="Fine", regions=(_region("Fine"),)),
        ),
    )
    review = member_review_from_scan("primary", profile)
    broken = next(s for s in review.current_sheets if s.sheet_name == "Broken")
    fine = next(s for s in review.current_sheets if s.sheet_name == "Fine")
    assert broken.failure_detail == "unexpected shape"
    assert len(fine.regions) == 1


def test_excel_member_profile_ignore_sheets_used_for_diff() -> None:
    # sanity: ExcelMemberProfile is importable and constructible the way the
    # diff helper expects for a package member.
    member_profile = ExcelMemberProfile(ignore_sheets=["X"])
    assert "X" in member_profile.ignore_sheets


def test_parse_a1_cell_accepts_a_plain_cell_reference() -> None:
    assert parse_a1_cell("b4") == (4, 2)
    assert parse_a1_cell("$C$7") == (7, 3)


def test_parse_a1_cell_rejects_a_range() -> None:
    try:
        parse_a1_cell("B4:F20")
    except ValueError:
        pass
    else:
        raise AssertionError("expected ValueError")


def test_parse_a1_range_treats_a_single_cell_as_a_1x1_range() -> None:
    assert parse_a1_range("B4") == (4, 2, 4, 2)


def test_parse_a1_range_parses_a_normal_range() -> None:
    assert parse_a1_range("B4:F20") == (4, 2, 20, 6)


def test_parse_a1_range_rejects_reversed_corners() -> None:
    try:
        parse_a1_range("F20:B4")
    except ValueError:
        pass
    else:
        raise AssertionError("expected ValueError")


def test_format_a1_range_round_trips_with_parse_a1_range() -> None:
    assert format_a1_range(*parse_a1_range("B4:F20")) == "B4:F20"
    assert format_a1_range(*parse_a1_range("B4")) == "B4"


def test_apply_anchor_click_shifts_the_region_preserving_size() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = review.current_sheets[0].regions[0]
    assert region.current_range == "A1:C5"
    moved = apply_anchor_click(region, 3, 2)
    assert moved.anchor_cell == "B3"
    assert moved.current_range == "B3:D7"  # same 3x5 footprint, shifted
    assert moved.available_columns == ("B", "C", "D")


def test_apply_anchor_click_prunes_identity_columns_that_fall_out_of_range() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = replace(review.current_sheets[0].regions[0], identity_columns=("A",))
    moved = apply_anchor_click(region, 3, 2)  # shifts away from column A
    assert moved.identity_columns == ()


def test_apply_manual_range_resizes_the_region_and_available_columns() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = review.current_sheets[0].regions[0]
    resized = apply_manual_range(region, "B2:D10")
    assert resized.anchor_cell == "B2"
    assert resized.current_range == "B2:D10"
    assert resized.available_columns == ("B", "C", "D")


def test_apply_manual_range_rejects_unparseable_text() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = review.current_sheets[0].regions[0]
    try:
        apply_manual_range(region, "not a range")
    except ValueError:
        pass
    else:
        raise AssertionError("expected ValueError")


def test_apply_region_transform_updates_only_the_targeted_region() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = apply_region_transform(
        review, "Data", region_id, lambda region: apply_manual_range(region, "B2:D10")
    )
    assert updated.current_sheets[0].regions[0].current_range == "B2:D10"
    # sibling sheet is untouched
    other = next(s for s in updated.current_sheets if s.sheet_name == "NewSheet")
    assert other.regions == ()


def test_apply_region_transform_propagates_a_transform_value_error() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    try:
        apply_region_transform(
            review, "Data", region_id, lambda region: apply_manual_range(region, "garbage")
        )
    except ValueError:
        pass
    else:
        raise AssertionError("expected ValueError")


def test_confirm_all_regions_only_touches_valid_regions() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    # make it invalid: keyed with no identity columns
    invalid_member = update_region_decision(review, "Data", region_id, mode="keyed")
    confirmed = confirm_all_regions(invalid_member)
    region = confirmed.current_sheets[0].regions[0]
    assert region.confirmed is False  # still invalid, left alone

    valid_member = update_region_decision(
        review, "Data", region_id, mode="keyed", identity_columns=("A",)
    )
    confirmed_valid = confirm_all_regions(valid_member)
    assert confirmed_valid.current_sheets[0].regions[0].confirmed is True
