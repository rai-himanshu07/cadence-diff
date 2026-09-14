"""Pure view-model/diff/resolved-configuration builder tests for the
mode-aware configuration workspace shell (plan-20260913, Step 7).
"""

from __future__ import annotations

from dataclasses import replace

from qc_tool.config.input_contract import INPUT_CONTRACT_VERSION
from qc_tool.config.profile import DeliverableProfile, ExcelMemberProfile, ExcelProfile
from qc_tool.config.resolved_input import (
    ResolvedColumn,
    ResolvedInputConfigurationV1,
    ResolvedMember,
    ResolvedRegion,
    ResolvedSelector,
    ResolvedSheet,
)
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
    compute_required_slide_warnings,
    compute_sheet_pairing_warnings,
    compute_slide_pairing_warnings,
    compute_warnings,
    confirm_all_regions,
    deck_review_from_titles,
    diff_profile_against_scan,
    diff_resolved_configurations,
    effective_sheet_pairing,
    effective_slide_pairing,
    format_a1_range,
    is_clean_profile_diff,
    member_review_from_scan,
    parse_a1_cell,
    parse_a1_range,
    regions_overlap,
    resolve_slide_anchors,
    set_expected_refresh_columns,
    set_identity_columns,
    set_ignore_columns,
    set_ordinal_columns,
    set_sheet_rename,
    set_slide_included,
    set_slide_rename,
    sheet_pairing,
    slide_pairing,
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


def _resolved_region(region_id: str = "r1", **overrides: object) -> ResolvedRegion:
    defaults: dict[str, object] = {"region_id": region_id}
    defaults.update(overrides)
    return ResolvedRegion(**defaults)  # type: ignore[arg-type]


def _resolved_sheet(sheet_id: str = "s1", **overrides: object) -> ResolvedSheet:
    defaults: dict[str, object] = {
        "sheet_id": sheet_id,
        "baseline_sheet_name": "Data",
        "current_sheet_name": "Data",
        "regions": (_resolved_region(),),
    }
    defaults.update(overrides)
    return ResolvedSheet(**defaults)  # type: ignore[arg-type]


def _resolved_config(**overrides: object) -> ResolvedInputConfigurationV1:
    defaults: dict[str, object] = {
        "inspection_contract_version": INPUT_CONTRACT_VERSION,
        "members": (
            ResolvedMember(member_id="primary", sheets=(_resolved_sheet(),)),
        ),
    }
    defaults.update(overrides)
    return ResolvedInputConfigurationV1(**defaults)  # type: ignore[arg-type]


def test_diff_resolved_configurations_is_empty_when_both_sides_are_none() -> None:
    assert diff_resolved_configurations(None, None) == ()


def test_diff_resolved_configurations_is_empty_for_an_identical_snapshot() -> None:
    config = _resolved_config()
    assert diff_resolved_configurations(config, config) == ()


def test_diff_resolved_configurations_discloses_a_legacy_predecessor_mismatch() -> None:
    diff = diff_resolved_configurations(None, _resolved_config())
    assert len(diff) == 1
    assert diff[0].scope == "configuration"
    assert diff[0].kind == "conflict"


def test_diff_resolved_configurations_ignores_pure_range_drift() -> None:
    """Concise by design: a routine range/first-data-row shift between two
    runs over growing data is NOT disclosed -- only a real resolution
    (mode/policy/pairing/column-role) change is.
    """
    previous = _resolved_config(
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(
                    _resolved_sheet(
                        regions=(
                            _resolved_region(current_data_range="A2:B10", current_first_data_row=2),
                        )
                    ),
                ),
            ),
        )
    )
    current = _resolved_config(
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(
                    _resolved_sheet(
                        regions=(
                            _resolved_region(current_data_range="A2:B40", current_first_data_row=2),
                        )
                    ),
                ),
            ),
        )
    )
    assert diff_resolved_configurations(previous, current) == ()


def test_diff_resolved_configurations_flags_a_region_mode_change() -> None:
    previous = _resolved_config()
    current = _resolved_config(
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(_resolved_sheet(regions=(_resolved_region(mode="keyed"),)),),
            ),
        )
    )
    diff = diff_resolved_configurations(previous, current)
    assert len(diff) == 1
    assert diff[0].scope == "primary/s1/r1"
    assert diff[0].kind == "conflict"
    assert "mode" in diff[0].description


def test_diff_resolved_configurations_flags_an_added_and_a_removed_sheet() -> None:
    previous = _resolved_config()
    current = _resolved_config(
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(_resolved_sheet(sheet_id="s2"),),
            ),
        )
    )
    diff = diff_resolved_configurations(previous, current)
    kinds = {entry.scope: entry.kind for entry in diff}
    assert kinds["primary/s1"] == "missing_from_profile"
    assert kinds["primary/s2"] == "new_in_scan"


def test_diff_resolved_configurations_flags_a_sheet_pairing_change() -> None:
    previous = _resolved_config()
    current = _resolved_config(
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(_resolved_sheet(baseline_sheet_name="Data (FY25)"),),
            ),
        )
    )
    diff = diff_resolved_configurations(previous, current)
    assert any(
        entry.scope == "primary/s1" and entry.kind == "conflict" for entry in diff
    )


def test_diff_resolved_configurations_flags_a_column_role_change() -> None:
    previous = _resolved_config(
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(
                    _resolved_sheet(
                        regions=(
                            _resolved_region(
                                columns=(ResolvedColumn(column_id="c1"),)
                            ),
                        )
                    ),
                ),
            ),
        )
    )
    current = _resolved_config(
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(
                    _resolved_sheet(
                        regions=(
                            _resolved_region(
                                columns=(
                                    ResolvedColumn(
                                        column_id="c1", alignment_role="identity"
                                    ),
                                )
                            ),
                        )
                    ),
                ),
            ),
        )
    )
    diff = diff_resolved_configurations(previous, current)
    assert len(diff) == 1
    assert "column role/policy changed for 1 column(s)" in diff[0].description


def test_diff_resolved_configurations_flags_an_added_and_a_removed_selector() -> None:
    previous = _resolved_config(
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(
                    _resolved_sheet(
                        selectors=(ResolvedSelector(selector_id="scenario"),)
                    ),
                ),
            ),
        )
    )
    current = _resolved_config(
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(
                    _resolved_sheet(
                        selectors=(ResolvedSelector(selector_id="fx_rate"),)
                    ),
                ),
            ),
        )
    )
    diff = diff_resolved_configurations(previous, current)
    kinds = {entry.scope: entry.kind for entry in diff}
    assert kinds["primary/s1/scenario"] == "missing_from_profile"
    assert kinds["primary/s1/fx_rate"] == "new_in_scan"


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


def test_baseline_regions_are_detected_alongside_current_regions() -> None:
    """Step 8: baseline-side regions are no longer discarded -- they seed
    the baseline auto-match used by build_resolved_configuration.
    """
    profile = MemberSetupProfile(
        member_id="primary",
        baseline_hash="a" * 64,
        current_hash="b" * 64,
        baseline_sheets=(SheetSetupProfile(sheet_name="Data", regions=(_region("Data"),)),),
        current_sheets=(SheetSetupProfile(sheet_name="Data", regions=(_region("Data"),)),),
    )
    review = member_review_from_scan("primary", profile)
    baseline_sheet = review.baseline_sheets[0]
    assert len(baseline_sheet.regions) == 1
    assert baseline_sheet.regions[0].current_range == "A1:C5"


def test_set_sheet_rename_declares_and_clears_a_pairing() -> None:
    review = member_review_from_scan(
        "primary",
        MemberSetupProfile(
            member_id="primary",
            baseline_hash="a" * 64,
            current_hash="b" * 64,
            baseline_sheets=(SheetSetupProfile(sheet_name="2025"),),
            current_sheets=(SheetSetupProfile(sheet_name="2026"),),
        ),
    )
    renamed = set_sheet_rename(review, "2026", "2025")
    assert renamed.sheet_renames == {"2026": "2025"}
    cleared = set_sheet_rename(renamed, "2026", None)
    assert cleared.sheet_renames == {}


def test_set_sheet_rename_enforces_one_to_one_baseline_claim() -> None:
    review = member_review_from_scan(
        "primary",
        MemberSetupProfile(
            member_id="primary",
            baseline_hash="a" * 64,
            current_hash="b" * 64,
            baseline_sheets=(SheetSetupProfile(sheet_name="Old"),),
            current_sheets=(
                SheetSetupProfile(sheet_name="New1"),
                SheetSetupProfile(sheet_name="New2"),
            ),
        ),
    )
    first = set_sheet_rename(review, "New1", "Old")
    second = set_sheet_rename(first, "New2", "Old")
    # New2 claiming "Old" releases New1's earlier claim.
    assert second.sheet_renames == {"New2": "Old"}


def test_effective_sheet_pairing_promotes_a_declared_rename() -> None:
    review = member_review_from_scan(
        "primary",
        MemberSetupProfile(
            member_id="primary",
            baseline_hash="a" * 64,
            current_hash="b" * 64,
            baseline_sheets=(SheetSetupProfile(sheet_name="2025 Data"),),
            current_sheets=(SheetSetupProfile(sheet_name="2026 Data"),),
        ),
    )
    _pairs, added, removed = sheet_pairing(review)
    assert added == ("2026 Data",)
    assert removed == ("2025 Data",)

    renamed = set_sheet_rename(review, "2026 Data", "2025 Data")
    pairs, added, removed = effective_sheet_pairing(renamed)
    assert pairs == (("2025 Data", "2026 Data"),)
    assert added == ()
    assert removed == ()


def test_compute_sheet_pairing_warnings_flags_unacknowledged_added_and_removed() -> None:
    review = member_review_from_scan(
        "primary",
        MemberSetupProfile(
            member_id="primary",
            baseline_hash="a" * 64,
            current_hash="b" * 64,
            baseline_sheets=(SheetSetupProfile(sheet_name="Old"),),
            current_sheets=(SheetSetupProfile(sheet_name="New"),),
        ),
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(review,))
    warnings = compute_sheet_pairing_warnings(state)
    codes = {w.code for w in warnings}
    assert "sheet_added:primary:New" in codes
    assert "sheet_removed:primary:Old" in codes


def test_compute_sheet_pairing_warnings_excludes_a_declared_rename() -> None:
    review = member_review_from_scan(
        "primary",
        MemberSetupProfile(
            member_id="primary",
            baseline_hash="a" * 64,
            current_hash="b" * 64,
            baseline_sheets=(SheetSetupProfile(sheet_name="Old"),),
            current_sheets=(SheetSetupProfile(sheet_name="New"),),
        ),
    )
    renamed = set_sheet_rename(review, "New", "Old")
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(renamed,))
    assert compute_sheet_pairing_warnings(state) == ()


def test_column_roles_are_mutually_exclusive() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = review.current_sheets[0].regions[0]
    region = set_identity_columns(region, ("A", "B"))
    assert region.identity_columns == ("A", "B")
    # Assigning B as ordinal evicts it from identity.
    region = set_ordinal_columns(region, ("B",))
    assert region.identity_columns == ("A",)
    assert region.ordinal_columns == ("B",)
    # Assigning A to ignore evicts it from identity.
    region = set_ignore_columns(region, ("A",))
    assert region.identity_columns == ()
    assert region.ignore_columns == ("A",)
    # Assigning A to expected_refresh evicts it from ignore.
    region = set_expected_refresh_columns(region, ("A",))
    assert region.ignore_columns == ()
    assert region.expected_refresh_columns == ("A",)


def test_is_valid_requires_first_data_row_when_header_intent_is_first_data_row() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review, "Data", region_id, header_intent="first_data_row"
    )
    region = updated.current_sheets[0].regions[0]
    assert region.is_valid is False
    fixed = update_region_decision(
        updated, "Data", region_id, header_intent="first_data_row", first_data_row=2
    )
    assert fixed.current_sheets[0].regions[0].is_valid is True


def test_is_valid_requires_reason_and_expiry_for_ignored_columns() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(review, "Data", region_id, ignore_columns=("A",))
    assert updated.current_sheets[0].regions[0].is_valid is False
    fixed = update_region_decision(
        review,
        "Data",
        region_id,
        ignore_columns=("A",),
        ignore_columns_reason="legacy scratch column",
        ignore_columns_expires_on="2027-01-01",
    )
    assert fixed.current_sheets[0].regions[0].is_valid is True


def test_regions_overlap_detects_overlapping_and_non_overlapping_ranges() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_a = review.current_sheets[0].regions[0]
    region_b = replace(region_a, region_id="other", current_range="B3:D8")
    region_c = replace(region_a, region_id="far", current_range="F1:G2")
    assert regions_overlap(region_a, region_b) is True
    assert regions_overlap(region_a, region_c) is False


def test_apply_region_transform_rejects_an_overlapping_manual_range() -> None:
    profile = MemberSetupProfile(
        member_id="primary",
        baseline_hash="a" * 64,
        current_hash="b" * 64,
        current_sheets=(
            SheetSetupProfile(
                sheet_name="Data",
                regions=(
                    _region("Data", min_row=1, min_col=1, max_row=5, max_col=3),
                    _region("Data", min_row=10, min_col=1, max_row=15, max_col=3),
                ),
            ),
        ),
    )
    review = member_review_from_scan("primary", profile)
    first_id = review.current_sheets[0].regions[0].region_id
    second_range = review.current_sheets[0].regions[1].current_range
    try:
        apply_region_transform(
            review, "Data", first_id, lambda r: apply_manual_range(r, second_range)
        )
    except ValueError:
        pass
    else:
        raise AssertionError("expected ValueError for an overlapping range")


def test_build_resolved_configuration_auto_matches_a_baseline_region() -> None:
    profile = MemberSetupProfile(
        member_id="primary",
        baseline_hash="a" * 64,
        current_hash="b" * 64,
        baseline_sheets=(
            SheetSetupProfile(
                sheet_name="Data",
                regions=(_region("Data", min_row=1, min_col=1, max_row=4, max_col=3),),
            ),
        ),
        current_sheets=(
            SheetSetupProfile(
                sheet_name="Data",
                regions=(_region("Data", min_row=1, min_col=1, max_row=5, max_col=3),),
            ),
        ),
    )
    review = member_review_from_scan("primary", profile)
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(review,))
    resolved = build_resolved_configuration(
        state, profile=DeliverableProfile(name="default"), profile_sha256="deadbeef"
    )
    region = resolved.members[0].sheets[0].regions[0]
    assert region.current_outer_range == "A1:C5"
    assert region.baseline_outer_range == "A1:C4"  # auto-matched, smaller baseline table


def test_build_resolved_configuration_carries_preamble_footer_and_first_data_row() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review,
        "Data",
        region_id,
        header_intent="first_data_row",
        first_data_row=2,
        preamble_rows=1,
        footer_rows=1,
        blank_key_policy="block",
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    resolved = build_resolved_configuration(
        state, profile=DeliverableProfile(name="default"), profile_sha256="deadbeef"
    )
    region = resolved.members[0].sheets[0].regions[0]
    assert region.current_first_data_row == 2
    assert region.current_preamble_rows == 1
    assert region.current_footer_rows == 1


def test_build_input_contract_carries_blank_key_policy_and_first_data_row() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review,
        "Data",
        region_id,
        mode="keyed",
        identity_columns=("A",),
        header_intent="first_data_row",
        first_data_row=2,
        blank_key_policy="tolerate",
        confirmed=True,
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    contract = build_input_contract(state)
    region_contract = contract.members[0].sheets[0].regions[0]
    assert region_contract.blank_key_policy == "tolerate"
    assert region_contract.preferred_first_data_row == 2


def test_build_input_contract_carries_ignore_column_exclusion() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review,
        "Data",
        region_id,
        ignore_columns=("A",),
        ignore_columns_reason="legacy scratch column",
        ignore_columns_expires_on="2027-01-01",
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    contract = build_input_contract(state)
    region_contract = contract.members[0].sheets[0].regions[0]
    ignore_column = next(c for c in region_contract.columns if c.column_id.endswith("_a"))
    assert ignore_column.comparison_policy == "ignore"
    assert ignore_column.exclusion is not None
    assert ignore_column.exclusion.reason == "legacy scratch column"


def test_deck_review_from_titles_seeds_every_slide_included() -> None:
    deck = deck_review_from_titles(
        baseline_titles=[(1, "Cover"), (2, "Revenue")],
        current_titles=[(1, "Cover"), (2, "Revenue"), (3, "Appendix")],
    )
    assert [s.title for s in deck.baseline_slides] == ["Cover", "Revenue"]
    assert [s.title for s in deck.current_slides] == ["Cover", "Revenue", "Appendix"]
    assert all(s.included for s in deck.current_slides)


def test_set_slide_included_toggles_only_the_target_current_slide() -> None:
    deck = deck_review_from_titles(
        baseline_titles=[], current_titles=[(1, "Cover"), (2, "Revenue")]
    )
    updated = set_slide_included(deck, 2, False)
    statuses = {s.slide_index: s.included for s in updated.current_slides}
    assert statuses == {1: True, 2: False}


def test_slide_pairing_pairs_by_exact_title() -> None:
    deck = deck_review_from_titles(
        baseline_titles=[(1, "Cover"), (2, "Old Name")],
        current_titles=[(1, "Cover"), (2, "New Name")],
    )
    pairs, added, removed = slide_pairing(deck)
    assert pairs == (("Cover", "Cover"),)
    assert added == ("New Name",)
    assert removed == ("Old Name",)


def test_set_slide_rename_enforces_one_to_one_baseline_claim() -> None:
    deck = deck_review_from_titles(
        baseline_titles=[(1, "Old")],
        current_titles=[(1, "New1"), (2, "New2")],
    )
    first = set_slide_rename(deck, "New1", "Old")
    second = set_slide_rename(first, "New2", "Old")
    assert second.slide_renames == {"New2": "Old"}


def test_effective_slide_pairing_promotes_a_declared_rename() -> None:
    deck = deck_review_from_titles(
        baseline_titles=[(1, "2025 Overview")],
        current_titles=[(1, "2026 Overview")],
    )
    _pairs, added, removed = slide_pairing(deck)
    assert added == ("2026 Overview",)
    assert removed == ("2025 Overview",)

    renamed = set_slide_rename(deck, "2026 Overview", "2025 Overview")
    pairs, added, removed = effective_slide_pairing(renamed)
    assert pairs == (("2025 Overview", "2026 Overview"),)
    assert added == ()
    assert removed == ()


def test_compute_slide_pairing_warnings_flags_unacknowledged_added_and_removed() -> None:
    deck = deck_review_from_titles(
        baseline_titles=[(1, "Old")], current_titles=[(1, "New")]
    )
    warnings = compute_slide_pairing_warnings(deck)
    codes = {w.code for w in warnings}
    assert "slide_added:New" in codes
    assert "slide_removed:Old" in codes
    assert all(w.severity == "block" for w in warnings)


def test_compute_slide_pairing_warnings_excludes_a_declared_rename() -> None:
    deck = deck_review_from_titles(
        baseline_titles=[(1, "Old")], current_titles=[(1, "New")]
    )
    renamed = set_slide_rename(deck, "New", "Old")
    assert compute_slide_pairing_warnings(renamed) == ()


def test_compute_slide_pairing_warnings_is_silent_for_a_single_sided_deck() -> None:
    # Preflight/final-package: no baseline deck at all.
    deck = deck_review_from_titles(baseline_titles=[], current_titles=[(1, "Cover")])
    assert compute_slide_pairing_warnings(deck) == ()
    assert compute_slide_pairing_warnings(None) == ()


def test_compute_required_slide_warnings_flags_a_missing_required_slide() -> None:
    deck = deck_review_from_titles(
        baseline_titles=[], current_titles=[(1, "Cover"), (2, "Revenue")]
    )
    warnings = compute_required_slide_warnings(deck, ("Revenue", "Appendix"))
    assert len(warnings) == 1
    assert warnings[0].code == "required_slide_missing:Appendix"
    assert warnings[0].severity == "caution"


def test_compute_required_slide_warnings_empty_when_all_present() -> None:
    deck = deck_review_from_titles(baseline_titles=[], current_titles=[(1, "Cover")])
    assert compute_required_slide_warnings(deck, ("Cover",)) == ()


def test_resolve_slide_anchors_flags_a_missing_anchor() -> None:
    deck = deck_review_from_titles(baseline_titles=[], current_titles=[(1, "Revenue")])
    warnings = resolve_slide_anchors(deck, ("Old Summary",))
    assert len(warnings) == 1
    assert warnings[0].code == "slide_anchor_missing:Old Summary"
    assert warnings[0].severity == "caution"


def test_resolve_slide_anchors_flags_an_ambiguous_duplicate_title_never_guessing() -> None:
    deck = deck_review_from_titles(
        baseline_titles=[],
        current_titles=[(1, "Regional Summary"), (2, "Regional Summary")],
    )
    warnings = resolve_slide_anchors(deck, ("Regional Summary",))
    assert len(warnings) == 1
    assert warnings[0].code == "slide_anchor_ambiguous:Regional Summary"
    assert warnings[0].severity == "caution"


def test_resolve_slide_anchors_is_silent_for_a_unique_resolved_anchor() -> None:
    deck = deck_review_from_titles(baseline_titles=[], current_titles=[(1, "Revenue")])
    assert resolve_slide_anchors(deck, ("Revenue",)) == ()


def test_resolve_slide_anchors_deduplicates_repeated_anchor_titles() -> None:
    # Two saved mappings on the SAME slide must not double-report.
    deck = deck_review_from_titles(baseline_titles=[], current_titles=[])
    warnings = resolve_slide_anchors(deck, ("Missing", "Missing"))
    assert len(warnings) == 1
