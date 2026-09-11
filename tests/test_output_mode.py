"""plan-20260910 Step 2: the versioned run-level `FindingOutputMode` contract.

`resolve_output_policy()` is exercised directly for all three modes, then
`run_qc()`/`_run_multi_package()` are exercised end to end to prove the
resolved policy actually governs population output and threads through a
multi-member package -- and that none of this ever perturbs `profile_sha256`.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from pydantic import ValidationError

from qc_tool.config.profile import (
    DeliverableProfile,
    PopulationPolicy,
    ResolvedOutputPolicy,
    ReviewPolicy,
    default_profile,
    profile_sha256,
    resolve_output_policy,
)
from qc_tool.coverage import FindingOutputMode
from qc_tool.engine import run_qc
from qc_tool.findings import FindingClass
from qc_tool.package import (
    PackageArtifact,
    PackageManifest,
    PackageMember,
    PackageSide,
    paths_by_member,
)
from tests.test_population_finalize import _write_formula_pair


def test_profile_mode_returns_the_profiles_own_policy_unchanged() -> None:
    review_policy = ReviewPolicy()  # populations disabled, the default
    resolved = resolve_output_policy(FindingOutputMode.PROFILE, review_policy)

    assert resolved.output_mode is FindingOutputMode.PROFILE
    assert resolved.source == "profile"
    assert resolved.populations.model_dump() == review_policy.populations.model_dump()
    assert resolved.populations.enabled is False


def test_profile_mode_honors_an_explicit_profile_policy_exactly() -> None:
    custom = PopulationPolicy(enabled=True, threshold=25)
    review_policy = ReviewPolicy(populations=custom)
    resolved = resolve_output_policy(FindingOutputMode.PROFILE, review_policy)

    assert resolved.populations.model_dump() == custom.model_dump()
    assert resolved.source == "profile"


def test_atomic_mode_always_disables_populations_even_when_the_profile_enables_them() -> (
    None
):
    review_policy = ReviewPolicy(populations=PopulationPolicy(enabled=True, threshold=1))
    resolved = resolve_output_policy(FindingOutputMode.ATOMIC, review_policy)

    assert resolved.output_mode is FindingOutputMode.ATOMIC
    assert resolved.source == "atomic_forced"
    assert resolved.populations.enabled is False
    # Never mutates the profile's own policy.
    assert review_policy.populations.enabled is True


def test_decision_mode_uses_the_profiles_own_policy_when_already_enabled() -> None:
    custom = PopulationPolicy(enabled=True, threshold=3, max_rectangles=7)
    review_policy = ReviewPolicy(populations=custom)
    resolved = resolve_output_policy(FindingOutputMode.DECISION, review_policy)

    assert resolved.source == "profile"
    assert resolved.populations.model_dump() == custom.model_dump()


def test_decision_mode_falls_back_to_a_conservative_built_in_default() -> None:
    review_policy = ReviewPolicy()  # populations disabled, the default
    resolved = resolve_output_policy(FindingOutputMode.DECISION, review_policy)

    assert resolved.source == "decision_built_in"
    assert resolved.populations.enabled is True
    # Every other field keeps its documented conservative default.
    assert resolved.populations.model_dump() == PopulationPolicy(enabled=True).model_dump()
    # Never mutates the profile's own (disabled) policy.
    assert review_policy.populations.enabled is False


def test_resolved_output_policy_is_frozen_and_versioned() -> None:
    resolved = resolve_output_policy(FindingOutputMode.ATOMIC, ReviewPolicy())
    assert resolved.version == 1
    assert isinstance(resolved, ResolvedOutputPolicy)


def test_population_default_class_order_is_canonical() -> None:
    assert PopulationPolicy().classes == (
        FindingClass.FORMULA_LOGIC_CHANGED,
        FindingClass.NUMBER_FORMAT_CHANGED,
    )


def test_resolved_policy_is_a_deep_immutable_snapshot() -> None:
    custom = PopulationPolicy(enabled=True, threshold=25)
    review_policy = ReviewPolicy(populations=custom)
    resolved = resolve_output_policy(FindingOutputMode.PROFILE, review_policy)

    review_policy.populations.threshold = 99

    assert resolved.populations.threshold == 25
    with pytest.raises(ValidationError):
        resolved.populations.threshold = 77


def test_output_mode_never_perturbs_profile_sha256() -> None:
    """`resolve_output_policy` never touches the profile object, so the exact
    same profile hashes identically regardless of what output mode a run
    later resolves against it -- output mode is a run-level fact, not a
    profile-level one.
    """
    profile = default_profile()
    before = profile_sha256(profile)
    for mode in FindingOutputMode:
        resolve_output_policy(mode, profile.review_policy)
    after = profile_sha256(profile)
    assert before == after


def test_run_qc_atomic_mode_forces_populations_off_despite_a_profile_that_enables_them(
    tmp_path: Path,
) -> None:
    base, curr = _write_formula_pair(tmp_path, rows=19)
    profile = DeliverableProfile(
        name="atomic-override",
        review_policy=ReviewPolicy(
            populations=PopulationPolicy(enabled=True, threshold=10)
        ),
    )

    result = run_qc(
        baseline_excel=base,
        current_excel=curr,
        profile=profile,
        output_mode=FindingOutputMode.ATOMIC,
    )

    assert result.requested_output_mode is FindingOutputMode.ATOMIC
    assert result.resolved_output_policy is not None
    assert result.resolved_output_policy.populations.enabled is False
    assert result.resolved_output_policy.source == "atomic_forced"
    formula_findings = [
        finding
        for finding in result.findings
        if finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
    ]
    assert len(formula_findings) == 19
    assert all(finding.population is None for finding in formula_findings)


def test_run_qc_decision_mode_forces_populations_on_despite_a_profile_that_disables_them(
    tmp_path: Path,
) -> None:
    base, curr = _write_formula_pair(tmp_path, rows=19)
    profile = default_profile()  # populations disabled, the default

    result = run_qc(
        baseline_excel=base,
        current_excel=curr,
        profile=profile,
        output_mode=FindingOutputMode.DECISION,
    )

    assert result.requested_output_mode is FindingOutputMode.DECISION
    assert result.resolved_output_policy is not None
    assert result.resolved_output_policy.populations.enabled is True
    assert result.resolved_output_policy.source == "decision_built_in"
    formula_findings = [
        finding
        for finding in result.findings
        if finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
    ]
    assert len(formula_findings) == 1
    assert formula_findings[0].population is not None
    assert formula_findings[0].population.member_count == 19
    # The profile's own object is untouched -- still disabled.
    assert profile.review_policy.populations.enabled is False


def test_run_qc_profile_mode_matches_omitting_output_mode_entirely(
    tmp_path: Path,
) -> None:
    base, curr = _write_formula_pair(tmp_path, rows=5)

    explicit = run_qc(
        baseline_excel=base,
        current_excel=curr,
        output_mode=FindingOutputMode.PROFILE,
    )
    implicit = run_qc(baseline_excel=base, current_excel=curr)

    assert explicit.requested_output_mode is implicit.requested_output_mode
    assert explicit.requested_output_mode is FindingOutputMode.PROFILE
    assert [f.model_dump(mode="json") for f in explicit.findings] == [
        f.model_dump(mode="json") for f in implicit.findings
    ]


def test_multi_package_threads_output_mode_to_every_member_and_the_merged_result(
    tmp_path: Path,
) -> None:
    # A non-"primary" member id forces the true multi-package path
    # (`_run_multi_package`); an all-"primary" manifest is a legacy
    # projection and `run_qc()` would take the plain single-file path instead.
    base_ops, curr_ops = _write_formula_pair(tmp_path, rows=19)
    manifest = PackageManifest(
        members=(
            PackageMember(
                member_id="ops",
                side=PackageSide.BASELINE,
                artifact=PackageArtifact.EXCEL,
                display_name="ops.xlsx",
            ),
            PackageMember(
                member_id="ops",
                side=PackageSide.CURRENT,
                artifact=PackageArtifact.EXCEL,
                display_name="ops.xlsx",
            ),
        )
    )
    files = {"baseline_excel:ops": base_ops, "current_excel:ops": curr_ops}
    paths_by_member(files, manifest)

    result = run_qc(
        package_manifest=manifest,
        package_files=files,
        output_mode=FindingOutputMode.DECISION,
    )

    assert result.requested_output_mode is FindingOutputMode.DECISION
    assert result.resolved_output_policy is not None
    assert result.resolved_output_policy.populations.enabled is True
    formula_findings = [
        finding
        for finding in result.findings
        if finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
    ]
    assert len(formula_findings) == 1
    assert formula_findings[0].population is not None
    assert formula_findings[0].population.member_count == 19
