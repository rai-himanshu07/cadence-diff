"""Pure view-model/diff/resolved-configuration builder tests for the
mode-aware configuration workspace shell (plan-20260913, Step 7).
"""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest

from qc_tool.config.input_contract import INPUT_CONTRACT_VERSION
from qc_tool.config.profile import (
    DeliverableProfile,
    ExcelMemberProfile,
    ExcelProfile,
    profile_sha256,
)
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
    LOW_KEY_OVERLAP_THRESHOLD,
    ConfigWorkspaceState,
    add_selector,
    apply_anchor_click,
    apply_manual_range,
    apply_region_transform,
    build_input_contract,
    build_resolved_configuration,
    compute_key_overlap_warnings,
    compute_required_slide_warnings,
    compute_sheet_pairing_warnings,
    compute_slide_pairing_warnings,
    compute_warnings,
    compute_workspace_readiness,
    confirm_all_regions,
    deck_review_from_titles,
    diff_profile_against_scan,
    diff_resolved_configurations,
    effective_region_data_range,
    effective_sheet_pairing,
    effective_slide_pairing,
    format_a1_range,
    input_contract_from_resolved_configuration,
    is_clean_profile_diff,
    low_key_overlap_warning_code,
    member_review_from_scan,
    parse_a1_cell,
    parse_a1_range,
    region_key_overlap_query_bounds,
    regions_overlap,
    resolve_slide_anchors,
    set_column_baseline_letter,
    set_expected_refresh_columns,
    set_identity_columns,
    set_ignore_columns,
    set_ordinal_columns,
    set_region_data_start,
    set_region_footer_rows,
    set_selector_baseline_cell,
    set_sheet_rename,
    set_slide_included,
    set_slide_rename,
    sheet_pairing,
    slide_pairing,
    slugify,
    summarize_contract_promotion,
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


def test_compute_warnings_removes_ranked_caution_after_explicit_choice() -> None:
    result = _scan_result(with_ranked_candidate=True)
    review = member_review_from_scan("primary", result.members["primary"])
    region_id = review.current_sheets[0].regions[0].region_id
    review = update_region_decision(
        review,
        "Data",
        region_id,
        mode="positional",
        confirmed=True,
    )
    state = ConfigWorkspaceState(
        mode=QCRunMode.CYCLE_COMPARISON,
        member_reviews=(review,),
    )

    assert not any(
        warning.code.startswith("ranked_candidate:")
        for warning in compute_warnings(result, state=state)
    )


def test_unresolved_blockers_requires_an_explicit_ranked_region_choice() -> None:
    review = member_review_from_scan(
        "primary", _member_profile(with_ranked_candidate=True)
    )
    state = ConfigWorkspaceState(
        mode=QCRunMode.CYCLE_COMPARISON,
        member_reviews=(review,),
    )

    blockers = unresolved_blockers(state, compute_warnings(_scan_result(
        with_ranked_candidate=True
    )))

    assert any("choose Match rows by key" in blocker for blocker in blockers)

    region_id = review.current_sheets[0].regions[0].region_id
    positional = update_region_decision(
        review,
        "Data",
        region_id,
        mode="positional",
        confirmed=True,
    )
    resolved = replace(state, member_reviews=(positional,))
    assert not any(
        "choose Match rows by key" in blocker
        for blocker in unresolved_blockers(
            resolved, compute_warnings(_scan_result(with_ranked_candidate=True))
        )
    )


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


def test_workspace_readiness_rejects_done_scan_before_expected_reviews_are_seeded() -> None:
    readiness = compute_workspace_readiness(
        ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON),
        (),
        setup_complete=True,
        expected_excel_member_ids=frozenset({"primary"}),
        ppt_required=False,
    )

    assert not readiness.ready
    assert "missing required Excel member review: primary" in readiness.blockers


def test_workspace_readiness_requires_exact_members_and_required_deck() -> None:
    primary = member_review_from_scan("primary", _member_profile())
    state = ConfigWorkspaceState(
        mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
        member_reviews=(primary,),
    )

    without_deck = compute_workspace_readiness(
        state,
        (),
        setup_complete=True,
        expected_excel_member_ids=frozenset({"primary"}),
        ppt_required=True,
    )
    with_deck = compute_workspace_readiness(
        replace(state, deck_review=deck_review_from_titles([], [(1, "Cover")])),
        (),
        setup_complete=True,
        expected_excel_member_ids=frozenset({"primary"}),
        ppt_required=True,
    )

    assert not without_deck.ready
    assert "PowerPoint review is not ready" in without_deck.blockers
    assert with_deck.ready
    assert with_deck.blockers == ()


def test_unresolved_blockers_flags_an_expired_region_exclusion() -> None:
    """``input_contract.py``'s own docstring promises "an expired exclusion
    blocks setup until it is renewed or removed" -- this is that promise's
    only enforcement point.
    """
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review,
        "Data",
        region_id,
        mode="excluded",
        exclusion_reason="known bad legacy tab",
        exclusion_expires_on="2000-01-01",
        confirmed=True,
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    blockers = unresolved_blockers(state, ())
    assert blockers
    assert "expired on 2000-01-01" in blockers[0]


def test_unresolved_blockers_allows_an_unexpired_region_exclusion() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review,
        "Data",
        region_id,
        mode="excluded",
        exclusion_reason="known bad legacy tab",
        exclusion_expires_on="2099-01-01",
        confirmed=True,
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    assert unresolved_blockers(state, ()) == ()


def test_unresolved_blockers_flags_an_expired_ignore_columns_exclusion() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review,
        "Data",
        region_id,
        ignore_columns=("B",),
        ignore_columns_reason="volatile helper column",
        ignore_columns_expires_on="2000-01-01",
        confirmed=True,
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    blockers = unresolved_blockers(state, ())
    assert blockers
    assert "ignored-column exclusion expired on 2000-01-01" in blockers[0]


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


def test_set_column_baseline_letter_stores_and_clears_an_override() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = review.current_sheets[0].regions[0]

    moved = set_column_baseline_letter(region, "A", "c")
    assert moved.baseline_letter_for("A") == "C"  # normalized upper
    assert moved.baseline_letter_for("B") == "B"  # untouched column unaffected

    cleared = set_column_baseline_letter(moved, "A", "")
    assert cleared.baseline_letter_for("A") == "A"
    assert cleared.column_baseline_letters == ()

    same_letter = set_column_baseline_letter(region, "A", "A")
    assert same_letter.column_baseline_letters == ()  # no-op override is not stored


def test_build_resolved_configuration_carries_a_baseline_letter_override() -> None:
    """plan-20260913 Step 12 fix: identity columns resolve separately on
    each side instead of always mirroring the current-side letter.
    """
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review, "Data", region_id, mode="keyed", identity_columns=("A",), confirmed=True
    )
    updated = apply_region_transform(
        updated,
        "Data",
        region_id,
        lambda r: set_column_baseline_letter(r, "A", "C"),
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    resolved = build_resolved_configuration(
        state, profile=DeliverableProfile(name="default"), profile_sha256="deadbeef"
    )
    region = resolved.members[0].sheets[0].regions[0]
    identity_column = next(c for c in region.columns if c.alignment_role == "identity")
    assert identity_column.current_letter == "A"
    assert identity_column.baseline_letter == "C"


def test_set_selector_baseline_cell_stores_and_clears_an_override() -> None:
    review = member_review_from_scan("primary", _member_profile())
    review = add_selector(review, "Data", label="Scenario", cell="B1")
    selector_id = review.current_sheets[0].selectors[0].selector_id

    moved = set_selector_baseline_cell(review, "Data", selector_id, "b2")
    assert moved.current_sheets[0].selectors[0].baseline_cell == "B2"  # normalized upper

    cleared = set_selector_baseline_cell(moved, "Data", selector_id, "")
    assert cleared.current_sheets[0].selectors[0].baseline_cell == ""

    with pytest.raises(ValueError):
        set_selector_baseline_cell(review, "Data", selector_id, "not-a-cell")


def test_build_resolved_configuration_carries_a_baseline_cell_override_for_a_selector() -> None:
    review = member_review_from_scan("primary", _member_profile())
    review = add_selector(review, "Data", label="Scenario", cell="B1")
    selector_id = review.current_sheets[0].selectors[0].selector_id
    review = set_selector_baseline_cell(review, "Data", selector_id, "B5")
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(review,))
    resolved = build_resolved_configuration(
        state, profile=DeliverableProfile(name="default"), profile_sha256="deadbeef"
    )
    selector = resolved.members[0].sheets[0].selectors[0]
    assert selector.current_cell == "B1"
    assert selector.baseline_cell == "B5"


def test_build_resolved_configuration_leaves_untouched_regions_automatic() -> None:
    review = member_review_from_scan("primary", _member_profile())
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(review,))
    resolved = build_resolved_configuration(
        state, profile=DeliverableProfile(name="default"), profile_sha256="deadbeef"
    )
    region = resolved.members[0].sheets[0].regions[0]
    assert region.mode == "automatic"
    assert region.columns == ()


def test_build_resolved_configuration_reports_excluded_coverage_for_an_excluded_region() -> None:
    """An intentional content exclusion is its own distinct coverage state
    -- never collapsed into "degraded_acknowledged", which names a
    different concept (a forced capability limitation the analyst
    accepted, not a deliberate scope choice).
    """
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review,
        "Data",
        region_id,
        mode="excluded",
        exclusion_reason="notes sheet, not comparable",
        exclusion_expires_on="2099-01-01",
        confirmed=True,
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    resolved = build_resolved_configuration(
        state, profile=DeliverableProfile(name="default"), profile_sha256="deadbeef"
    )
    region = resolved.members[0].sheets[0].regions[0]
    assert region.mode == "excluded"
    assert region.coverage == "excluded"


def test_build_resolved_configuration_without_file_hashes_leaves_source_hashes_none() -> None:
    """Documents the DEFAULT (no ``file_hashes`` argument) behavior -- this
    is exactly the shape that, before the fix, silently made every wizard-
    submitted run fail `validate_freshness()` as "stale" the instant that
    check was actually wired into `perform_run()` (plan-20260913 Step 10).
    A caller building a resolved configuration that will reach
    `perform_run()` MUST pass `file_hashes`; this test exists to make that
    contract explicit, not to bless the omission.
    """
    review = member_review_from_scan("primary", _member_profile())
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(review,))
    resolved = build_resolved_configuration(
        state, profile=DeliverableProfile(name="default"), profile_sha256="deadbeef"
    )
    assert resolved.members[0].baseline_source_sha256 is None
    assert resolved.members[0].current_source_sha256 is None


def test_build_resolved_configuration_populates_source_hashes_from_file_hashes() -> None:
    review = member_review_from_scan("primary", _member_profile())
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(review,))
    resolved = build_resolved_configuration(
        state,
        profile=DeliverableProfile(name="default"),
        profile_sha256="deadbeef",
        file_hashes={"baseline_excel": "a" * 64, "current_excel": "b" * 64},
    )
    assert resolved.members[0].baseline_source_sha256 == "a" * 64
    assert resolved.members[0].current_source_sha256 == "b" * 64


def test_build_resolved_configuration_populates_member_scoped_source_hashes() -> None:
    review = member_review_from_scan("secondary", _member_profile())
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(review,))
    resolved = build_resolved_configuration(
        state,
        profile=DeliverableProfile(name="default"),
        profile_sha256="deadbeef",
        file_hashes={
            "baseline_excel:secondary": "c" * 64,
            "current_excel:secondary": "d" * 64,
        },
    )
    assert resolved.members[0].baseline_source_sha256 == "c" * 64
    assert resolved.members[0].current_source_sha256 == "d" * 64


def test_build_resolved_configuration_output_survives_perform_run_freshness_check(
    tmp_path: Path,
) -> None:
    """Integration proof for a REAL regression: a resolved configuration
    built the exact way the workspace's own ``_finalize`` builds it (via
    ``build_resolved_configuration``, not hand-constructed with already-
    correct hashes) must actually survive `perform_run()`'s
    `validate_freshness()` call -- this is the precise gap an independent
    review caught: every wizard-submitted run was silently rejected as
    stale because `file_hashes` was never threaded through.
    """
    from openpyxl import Workbook

    from qc_tool.history.store import sha256_file
    from qc_tool.run_service import perform_run

    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    for path, value in ((baseline_path, 1), (current_path, 2)):
        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "Data"
        sheet["A1"] = value
        workbook.save(path)

    review = member_review_from_scan(
        "primary",
        MemberSetupProfile(
            member_id="primary",
            baseline_hash="a" * 64,
            current_hash="b" * 64,
            current_sheets=(SheetSetupProfile(sheet_name="Data", regions=()),),
            baseline_sheets=(SheetSetupProfile(sheet_name="Data", regions=()),),
        ),
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(review,))
    profile = DeliverableProfile(name="default")
    file_hashes = {
        "baseline_excel": sha256_file(baseline_path),
        "current_excel": sha256_file(current_path),
    }
    resolved = build_resolved_configuration(
        state,
        profile=profile,
        profile_sha256=profile_sha256(profile),
        file_hashes=file_hashes,
    )

    artifacts = perform_run(
        tmp_path / "work",
        {"baseline_excel": baseline_path, "current_excel": current_path},
        {},
        profile,
        mode=QCRunMode.CYCLE_COMPARISON,
        resolved_input_configuration=resolved,
    )
    assert artifacts.result is not None


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


def test_region_data_start_is_canonical_and_clears_stale_boundary_state() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = replace(
        review.current_sheets[0].regions[0],
        header_intent="first_data_row",
        first_data_row=3,
        preamble_rows=2,
    )

    automatic = set_region_data_start(region, None)
    assert automatic.header_intent == "automatic"
    assert automatic.first_data_row is None
    assert automatic.preamble_rows == 0
    assert effective_region_data_range(automatic) is None

    no_header = set_region_data_start(region, 1)
    assert no_header.header_intent == "no_header"
    assert no_header.first_data_row is None
    assert no_header.preamble_rows == 0
    assert effective_region_data_range(no_header) == "A1:C5"

    explicit = set_region_data_start(region, 3)
    assert explicit.header_intent == "first_data_row"
    assert explicit.first_data_row == 3
    assert explicit.preamble_rows == 0
    assert effective_region_data_range(explicit) == "A3:C5"


def test_region_footer_must_leave_a_data_row() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = set_region_data_start(review.current_sheets[0].regions[0], 3)

    assert effective_region_data_range(set_region_footer_rows(region, 1)) == "A3:C4"
    with pytest.raises(ValueError, match="leave at least one data row"):
        set_region_footer_rows(region, 3)


def test_effective_data_range_projects_relative_boundaries_to_baseline() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = replace(
        set_region_footer_rows(
            set_region_data_start(review.current_sheets[0].regions[0], 3), 1
        ),
        baseline_range="D10:F20",
    )

    assert effective_region_data_range(region) == "A3:C4"
    assert effective_region_data_range(
        region, outer_range=region.baseline_range
    ) == "D12:F19"


def test_apply_anchor_click_shifts_the_region_preserving_size() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = review.current_sheets[0].regions[0]
    assert region.current_range == "A1:C5"
    moved = apply_anchor_click(region, 3, 2)
    assert moved.anchor_cell == "B3"
    assert moved.current_range == "B3:D7"  # same 3x5 footprint, shifted
    assert moved.available_columns == ("B", "C", "D")


def test_apply_anchor_click_moves_the_explicit_data_start_with_the_region() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = set_region_data_start(review.current_sheets[0].regions[0], 3)

    moved = apply_anchor_click(region, 5, 2)

    assert moved.current_range == "B5:D9"
    assert moved.first_data_row == 7
    assert effective_region_data_range(moved) == "B7:D9"


def test_apply_manual_range_rejects_a_resize_that_consumes_the_data_band() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = set_region_footer_rows(
        set_region_data_start(review.current_sheets[0].regions[0], 3), 1
    )

    with pytest.raises(ValueError, match="leave at least one data row"):
        apply_manual_range(region, "A1:C3")


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


def test_is_valid_is_false_once_an_exclusion_expiry_date_has_passed() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    still_current = update_region_decision(
        review,
        "Data",
        region_id,
        mode="excluded",
        exclusion_reason="known bad legacy tab",
        exclusion_expires_on="2099-01-01",
    )
    assert still_current.current_sheets[0].regions[0].is_valid is True
    expired = update_region_decision(
        review,
        "Data",
        region_id,
        mode="excluded",
        exclusion_reason="known bad legacy tab",
        exclusion_expires_on="2000-01-01",
    )
    assert expired.current_sheets[0].regions[0].is_valid is False


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
        baseline_range="A1:C5",
        blank_key_policy="block",
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    resolved = build_resolved_configuration(
        state, profile=DeliverableProfile(name="default"), profile_sha256="deadbeef"
    )
    region = resolved.members[0].sheets[0].regions[0]
    assert region.current_first_data_row == 2
    assert region.current_preamble_rows == 0
    assert region.current_footer_rows == 1
    assert region.baseline_first_data_row == 2
    assert region.baseline_preamble_rows == 0
    assert region.baseline_footer_rows == 1


def test_build_resolved_configuration_ignores_stale_first_data_row_for_no_header() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    review = update_region_decision(
        review,
        "Data",
        region_id,
        mode="positional",
        confirmed=True,
        header_intent="no_header",
        first_data_row=4,
        preamble_rows=3,
    )
    state = ConfigWorkspaceState(
        mode=QCRunMode.CYCLE_COMPARISON,
        member_reviews=(review,),
    )

    resolved = build_resolved_configuration(
        state,
        profile=DeliverableProfile(name="resolved-boundaries"),
        profile_sha256="a" * 64,
    )

    region = resolved.members[0].sheets[0].regions[0]
    assert region.current_first_data_row is None
    assert region.current_preamble_rows == 0


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


def test_input_contract_from_resolved_configuration_reconstructs_a_region() -> None:
    resolved = ResolvedInputConfigurationV1(
        inspection_contract_version=INPUT_CONTRACT_VERSION,
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(
                    ResolvedSheet(
                        sheet_id="s1",
                        baseline_sheet_name="Data",
                        current_sheet_name="Data",
                        regions=(
                            ResolvedRegion(
                                region_id="r1",
                                mode="keyed",
                                header_intent="first_data_row",
                                current_outer_range="A1:C10",
                                current_first_data_row=2,
                                columns=(
                                    ResolvedColumn(
                                        column_id="c1", alignment_role="identity"
                                    ),
                                ),
                            ),
                        ),
                    ),
                ),
            ),
        ),
    )
    contract = input_contract_from_resolved_configuration(resolved)
    assert len(contract.members) == 1
    sheet = contract.members[0].sheets[0]
    assert sheet.sheet_id == "s1"
    assert sheet.preferred_sheet_name == "Data"
    assert len(sheet.regions) == 1
    region = sheet.regions[0]
    assert region.mode == "keyed"
    assert region.anchor_cell == "A1"
    assert region.preferred_current_range == "A1:C10"
    assert region.preferred_first_data_row == 2
    assert len(region.columns) == 1
    assert region.columns[0].alignment_role == "identity"


def test_input_contract_from_resolved_configuration_skips_an_excluded_region() -> None:
    resolved = ResolvedInputConfigurationV1(
        inspection_contract_version=INPUT_CONTRACT_VERSION,
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(
                    ResolvedSheet(
                        sheet_id="s1",
                        regions=(
                            ResolvedRegion(
                                region_id="r1",
                                mode="excluded",
                                current_outer_range="A1:C10",
                                degraded_reason="not comparable this run",
                            ),
                        ),
                    ),
                ),
            ),
        ),
    )
    contract = input_contract_from_resolved_configuration(resolved)
    assert contract.members[0].sheets[0].regions == ()


def test_input_contract_from_resolved_configuration_skips_an_ignored_column() -> None:
    resolved = ResolvedInputConfigurationV1(
        inspection_contract_version=INPUT_CONTRACT_VERSION,
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(
                    ResolvedSheet(
                        sheet_id="s1",
                        regions=(
                            ResolvedRegion(
                                region_id="r1",
                                current_outer_range="A1:C10",
                                columns=(
                                    ResolvedColumn(
                                        column_id="c1", comparison_policy="ignore"
                                    ),
                                ),
                            ),
                        ),
                    ),
                ),
            ),
        ),
    )
    contract = input_contract_from_resolved_configuration(resolved)
    assert contract.members[0].sheets[0].regions[0].columns == ()


def test_input_contract_from_resolved_configuration_skips_a_valueless_selector() -> None:
    resolved = ResolvedInputConfigurationV1(
        inspection_contract_version=INPUT_CONTRACT_VERSION,
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(
                    ResolvedSheet(
                        sheet_id="s1",
                        selectors=(
                            ResolvedSelector(selector_id="scenario", current_cell=None),
                        ),
                    ),
                ),
            ),
        ),
    )
    contract = input_contract_from_resolved_configuration(resolved)
    assert contract.members[0].sheets[0].selectors == ()


def test_summarize_contract_promotion_describes_a_brand_new_profile() -> None:
    resolved = ResolvedInputConfigurationV1(
        inspection_contract_version=INPUT_CONTRACT_VERSION,
        members=(
            ResolvedMember(
                member_id="primary",
                sheets=(ResolvedSheet(sheet_id="s1"), ResolvedSheet(sheet_id="s2")),
            ),
        ),
    )
    new_contract = input_contract_from_resolved_configuration(resolved)
    lines = summarize_contract_promotion(None, new_contract)
    assert len(lines) == 1
    assert "new profile with 2 configured sheet(s)" in lines[0]


def test_summarize_contract_promotion_reports_a_no_op_when_identical() -> None:
    resolved = ResolvedInputConfigurationV1(
        inspection_contract_version=INPUT_CONTRACT_VERSION,
        members=(ResolvedMember(member_id="primary", sheets=(ResolvedSheet(sheet_id="s1"),)),),
    )
    contract = input_contract_from_resolved_configuration(resolved)
    assert summarize_contract_promotion(contract, contract) == (
        "No configuration changes -- saving would be a no-op.",
    )


def test_summarize_contract_promotion_flags_added_removed_and_changed_sheets() -> None:
    existing = input_contract_from_resolved_configuration(
        ResolvedInputConfigurationV1(
            inspection_contract_version=INPUT_CONTRACT_VERSION,
            members=(
                ResolvedMember(
                    member_id="primary",
                    sheets=(
                        ResolvedSheet(sheet_id="keep_changed", current_sheet_name="A"),
                        ResolvedSheet(sheet_id="removed_only"),
                    ),
                ),
            ),
        )
    )
    new = input_contract_from_resolved_configuration(
        ResolvedInputConfigurationV1(
            inspection_contract_version=INPUT_CONTRACT_VERSION,
            members=(
                ResolvedMember(
                    member_id="primary",
                    sheets=(
                        ResolvedSheet(sheet_id="keep_changed", current_sheet_name="B"),
                        ResolvedSheet(sheet_id="added_only"),
                    ),
                ),
            ),
        )
    )
    lines = summarize_contract_promotion(existing, new)
    joined = " ".join(lines)
    assert "1 sheet(s) gain saved configuration" in joined
    assert "1 sheet(s) lose their saved configuration" in joined
    assert "1 sheet(s) have different saved configuration" in joined


# --- plan-20260913 Step 12 Fix 5: low-key-overlap acknowledgement ----------


def _keyed_state(region_id: str | None = None) -> tuple[ConfigWorkspaceState, str]:
    review = member_review_from_scan("primary", _member_profile())
    rid = region_id or review.current_sheets[0].regions[0].region_id
    updated = update_region_decision(
        review, "Data", rid, mode="keyed", identity_columns=("A",), confirmed=True
    )
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(updated,))
    return state, rid


def test_region_key_overlap_query_bounds_uses_current_range_and_explicit_baseline_range() -> (
    None
):
    review = member_review_from_scan("primary", _member_profile())
    region = review.current_sheets[0].regions[0]
    region = replace(region, current_range="A2:C10", baseline_range="A1:C9")

    bounds = region_key_overlap_query_bounds(region, None)

    assert bounds == (1, 9, 2, 10)


def test_region_key_overlap_query_bounds_falls_back_to_matched_baseline_region() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = replace(review.current_sheets[0].regions[0], current_range="A2:C10")
    baseline_region = replace(region, current_range="A1:C9")

    bounds = region_key_overlap_query_bounds(region, baseline_region)

    assert bounds == (1, 9, 2, 10)


def test_region_key_overlap_query_bounds_is_none_without_a_resolvable_baseline_range() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region = review.current_sheets[0].regions[0]

    assert region_key_overlap_query_bounds(region, None) is None


def test_compute_key_overlap_warnings_flags_a_low_ratio_keyed_region() -> None:
    state, region_id = _keyed_state()
    warnings = compute_key_overlap_warnings(state, {region_id: 0.5})

    assert len(warnings) == 1
    warning = warnings[0]
    assert warning.code == low_key_overlap_warning_code("primary", "Data", region_id)
    assert warning.severity == "block"
    assert "50%" in warning.message


def test_compute_key_overlap_warnings_silent_above_threshold() -> None:
    state, region_id = _keyed_state()
    assert compute_key_overlap_warnings(state, {region_id: LOW_KEY_OVERLAP_THRESHOLD}) == ()
    assert compute_key_overlap_warnings(state, {region_id: 0.99}) == ()


def test_compute_key_overlap_warnings_silent_when_never_queried() -> None:
    state, _region_id = _keyed_state()
    assert compute_key_overlap_warnings(state, {}) == ()


def test_compute_key_overlap_warnings_ignores_non_keyed_regions() -> None:
    review = member_review_from_scan("primary", _member_profile())
    region_id = review.current_sheets[0].regions[0].region_id
    state = ConfigWorkspaceState(mode=QCRunMode.CYCLE_COMPARISON, member_reviews=(review,))
    # mode stays "automatic" -- a low ratio recorded against it (should
    # never happen in practice) still never produces a warning.
    assert compute_key_overlap_warnings(state, {region_id: 0.1}) == ()


def test_build_resolved_configuration_reports_degraded_acknowledged_for_low_overlap() -> None:
    state, region_id = _keyed_state()
    code = low_key_overlap_warning_code("primary", "Data", region_id)

    resolved = build_resolved_configuration(
        state,
        profile=DeliverableProfile(name="default"),
        profile_sha256="deadbeef",
        warnings_acknowledged=(code,),
        key_overlap_ratios={region_id: 0.5},
    )

    region = resolved.members[0].sheets[0].regions[0]
    assert region.coverage == "degraded_acknowledged"


def test_build_resolved_configuration_keeps_confirmed_coverage_when_not_acknowledged() -> None:
    state, region_id = _keyed_state()

    resolved = build_resolved_configuration(
        state,
        profile=DeliverableProfile(name="default"),
        profile_sha256="deadbeef",
        key_overlap_ratios={region_id: 0.5},
    )

    region = resolved.members[0].sheets[0].regions[0]
    assert region.coverage == "confirmed"


def test_build_resolved_configuration_keeps_confirmed_coverage_when_ratio_is_high() -> None:
    state, region_id = _keyed_state()
    code = low_key_overlap_warning_code("primary", "Data", region_id)

    resolved = build_resolved_configuration(
        state,
        profile=DeliverableProfile(name="default"),
        profile_sha256="deadbeef",
        warnings_acknowledged=(code,),
        key_overlap_ratios={region_id: 0.99},
    )

    region = resolved.members[0].sheets[0].regions[0]
    assert region.coverage == "confirmed"
