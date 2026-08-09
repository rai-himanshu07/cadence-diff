"""Pure counterfactual acceptance and review-floor preview contracts."""

from __future__ import annotations

from collections.abc import Sequence

import pytest

from qc_tool.config.profile import NumericTolerance
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingExpectedReason,
    FindingSubtype,
    NumericCounterfactualBasis,
    Severity,
)
from qc_tool.triage.preview import (
    CounterfactualPolicy,
    PreviewReviewFloor,
    canonical_aggregate_digest,
    canonical_basis_digest,
    preview_policy,
)


def _value_finding(finding_id: str, location: str = "B2") -> Finding:
    return Finding(
        finding_id=finding_id,
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        subtype=FindingSubtype.VALUE_REPLACEMENT,
        sheet="Data",
        location=location,
        baseline_value="1,000",
        current_value="1,005",
        message="Data value changed",
    )


def _basis(
    location: str = "B2",
    baseline: float = 1_000.0,
    current: float = 1_005.0,
) -> NumericCounterfactualBasis:
    return NumericCounterfactualBasis(
        baseline=baseline,
        current=current,
        number_format="#,##0",
        sheet="Data",
        location=location,
    )


def test_canonical_basis_digest_is_stable_and_field_sensitive() -> None:
    basis = _basis()

    assert canonical_basis_digest(basis) == canonical_basis_digest(basis)
    assert canonical_basis_digest(basis) != canonical_basis_digest(
        _basis(current=1_006.0)
    )
    assert canonical_basis_digest(basis) != canonical_basis_digest(
        _basis(location="B3")
    )


def test_canonical_aggregate_digest_is_order_independent_and_membership_sensitive() -> None:
    first = _basis("B2", 1.0, 2.0)
    second = _basis("B3", 3.0, 4.0)

    assert canonical_aggregate_digest({"F1": first, "F2": second}) == (
        canonical_aggregate_digest({"F2": second, "F1": first})
    )
    assert canonical_aggregate_digest({"F1": first}) != canonical_aggregate_digest(
        {"F1": first, "F2": second}
    )


def test_preview_acceptance_moves_finding_to_visible_info() -> None:
    finding = _value_finding("F1")
    preview = preview_policy(
        [finding],
        {"F1": _basis()},
        None,
        CounterfactualPolicy(
            acceptance=NumericTolerance(absolute=10.0, relative=0.0)
        ),
    )

    assert preview.affected_atomics == 1
    assert preview.accepted_atomics == 1
    assert preview.after_atomic[Severity.INFO.value] == 1
    assert preview.after_atomic[Severity.CRITICAL.value] == 0
    assert len(preview.affected_decisions) == 1
    assert preview.affected_decisions[0].reasons == ("within_acceptance",)


def test_preview_review_floor_moves_noise_below_action_queue() -> None:
    finding = _value_finding("F1")
    preview = preview_policy(
        [finding],
        {"F1": _basis(baseline=1_000_000.0, current=1_000_000.1)},
        None,
        CounterfactualPolicy(
            review_floor=PreviewReviewFloor.WITHIN_TOLERANCE
        ),
    )

    assert preview.affected_atomics == 1
    assert preview.accepted_atomics == 0
    assert preview.review_floor_atomics == 1
    assert preview.after_atomic[Severity.INFO.value] == 1
    assert preview.affected_decisions[0].reasons == ("below_review_floor",)


def test_preview_preserves_existing_expected_findings() -> None:
    finding = _value_finding("F1")
    finding.mark_expected(FindingExpectedReason.CADENCE_EXTENSION)

    preview = preview_policy(
        [finding],
        {"F1": _basis()},
        None,
        CounterfactualPolicy(
            acceptance=NumericTolerance(absolute=1_000.0, relative=1.0),
            review_floor=PreviewReviewFloor.MATERIAL,
        ),
    )

    assert preview.affected_atomics == 0
    assert preview.accepted_atomics == 0
    assert preview.before_atomic == preview.after_atomic
    assert preview.after_atomic[Severity.EXPECTED.value] == 1


def test_preview_does_not_mutate_stored_findings() -> None:
    finding = _value_finding("F1")
    finding.severity = Severity.WARNING
    finding.severity_overridden = True
    finding.analyst_comment = "keep this decision"

    preview_policy(
        [finding],
        {"F1": _basis()},
        None,
        CounterfactualPolicy(
            acceptance=NumericTolerance(absolute=10.0, relative=0.0)
        ),
    )

    assert finding.severity is Severity.WARNING
    assert finding.severity_overridden is True
    assert finding.analyst_comment == "keep this decision"
    assert finding.materiality is None


def test_preview_with_no_bases_is_an_identity_view() -> None:
    preview = preview_policy(
        [_value_finding("F1")],
        {},
        None,
        CounterfactualPolicy(),
    )

    assert preview.before_atomic == preview.after_atomic
    assert preview.before_decision == preview.after_decision
    assert preview.affected_atomics == 0
    assert preview.affected_decisions == ()


def test_oversized_preview_refuses_without_reading_findings() -> None:
    class UnreadableFindings(Sequence[Finding]):
        def __len__(self) -> int:
            return 50_001

        def __getitem__(self, index):  # type: ignore[override]
            raise AssertionError(f"preview read finding {index}")

    with pytest.raises(ValueError, match="preview is unavailable"):
        preview_policy(
            UnreadableFindings(),
            {},
            None,
            CounterfactualPolicy(),
        )


def test_preview_rejects_unknown_finding_id() -> None:
    with pytest.raises(ValueError, match="no finding with id F99"):
        preview_policy(
            [_value_finding("F1")],
            {"F99": _basis()},
            None,
            CounterfactualPolicy(),
        )


def test_preview_rejects_basis_locator_mismatch() -> None:
    with pytest.raises(ValueError, match="counterfactual basis mismatch"):
        preview_policy(
            [_value_finding("F1", "B2")],
            {"F1": _basis("C5")},
            None,
            CounterfactualPolicy(),
        )


def test_counterfactual_policy_rejects_negative_tolerance() -> None:
    with pytest.raises(ValueError, match="cannot be negative"):
        CounterfactualPolicy(
            acceptance=NumericTolerance(absolute=-1.0, relative=0.0)
        )
