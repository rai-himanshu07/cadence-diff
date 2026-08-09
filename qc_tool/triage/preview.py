"""Pure counterfactual policy views over private typed numeric evidence."""

from __future__ import annotations

import copy
import hashlib
import json
from collections.abc import Mapping, Sequence
from enum import StrEnum

from pydantic import BaseModel, Field, model_validator

from qc_tool.config.profile import DeliverableProfile, NumericTolerance
from qc_tool.excel.materiality import classify_numeric_pair
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingSubtype,
    Materiality,
    NumericCounterfactualBasis,
    Severity,
)
from qc_tool.review import (
    ReviewGroup,
    build_pattern_groups,
    count_pattern_groups,
    format_group_ranges,
)
from qc_tool.triage.rules import assign_severity

POLICY_PREVIEW_MAX_FINDINGS = 50_000


class PolicyPreviewUnavailable(ValueError):
    """The run is too large to rebuild two complete policy views safely."""


class PreviewReviewFloor(StrEnum):
    NOISE = "noise"
    WITHIN_TOLERANCE = "within_tolerance"
    RECENT_RESTATEMENT = "recent_restatement"
    MATERIAL = "material"


_MATERIALITY_ORDER = (
    Materiality.NOISE,
    Materiality.WITHIN_TOLERANCE,
    Materiality.RECENT_RESTATEMENT,
    Materiality.MATERIAL,
)


class CounterfactualPolicy(BaseModel):
    acceptance: NumericTolerance = Field(default_factory=NumericTolerance)
    review_floor: PreviewReviewFloor = PreviewReviewFloor.NOISE

    @model_validator(mode="after")
    def validate_nonnegative(self) -> CounterfactualPolicy:
        if self.acceptance.absolute < 0 or self.acceptance.relative < 0:
            raise ValueError("counterfactual tolerances cannot be negative")
        return self


class PreviewDecision(BaseModel):
    before_group_id: str
    after_group_ids: tuple[str, ...]
    finding_class: str
    where: str
    location: str
    affected_members: int
    before_severity: str
    after_severities: tuple[str, ...]
    reasons: tuple[str, ...]

    model_config = {"frozen": True}


class PolicyPreview(BaseModel):
    before_atomic: dict[str, int]
    after_atomic: dict[str, int]
    before_decision: dict[str, int]
    after_decision: dict[str, int]
    affected_atomics: int
    accepted_atomics: int
    review_floor_atomics: int
    affected_decisions: tuple[PreviewDecision, ...] = ()

    model_config = {"frozen": True}


def canonical_basis_digest(basis: NumericCounterfactualBasis) -> str:
    encoded = json.dumps(
        basis.model_dump(mode="json"),
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def canonical_aggregate_digest(
    bases: Mapping[str, NumericCounterfactualBasis],
) -> str:
    payload = {
        finding_id: bases[finding_id].model_dump(mode="json")
        for finding_id in sorted(bases)
    }
    encoded = json.dumps(
        payload,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _within_acceptance(
    baseline: float,
    current: float,
    tolerance: NumericTolerance,
) -> bool:
    delta = abs(current - baseline)
    within_absolute = tolerance.absolute > 0 and delta <= tolerance.absolute
    within_relative = (
        tolerance.relative > 0
        and baseline != 0
        and delta / abs(baseline) <= tolerance.relative
    )
    return within_absolute or within_relative


def _engine_findings(
    findings: Sequence[Finding],
    profile: DeliverableProfile | None,
) -> list[Finding]:
    cloned = [copy.deepcopy(finding) for finding in findings]
    for finding in cloned:
        finding.analyst_comment = ""
        finding.severity_overridden = False
        finding.counterfactual_basis = None
        finding.severity = assign_severity(finding, profile)
    return cloned


def _severity_counts(findings: list[Finding]) -> dict[str, int]:
    return {
        severity.value: sum(finding.severity is severity for finding in findings)
        for severity in Severity
    }


def _decision_counts(groups: list[ReviewGroup]) -> dict[str, int]:
    counts = count_pattern_groups(groups).review_items
    return {severity.value: counts.get(severity, 0) for severity in Severity}


def _decision_records(
    before_groups: list[ReviewGroup],
    after_groups: list[ReviewGroup],
    changed_ids: set[str],
    reasons_by_finding: Mapping[str, set[str]],
) -> tuple[PreviewDecision, ...]:
    after_by_finding = {
        member.finding_id: group
        for group in after_groups
        for member in group.members
    }
    records: list[PreviewDecision] = []
    for before_group in before_groups:
        affected_ids = sorted(
            member.finding_id
            for member in before_group.members
            if member.finding_id in changed_ids
        )
        if not affected_ids:
            continue
        after_for_members = [
            after_by_finding[finding_id]
            for finding_id in affected_ids
            if finding_id in after_by_finding
        ]
        records.append(
            PreviewDecision(
                before_group_id=before_group.group_id,
                after_group_ids=tuple(
                    sorted({group.group_id for group in after_for_members})
                ),
                finding_class=before_group.finding_class.value,
                where=before_group.sheet or before_group.slide or "",
                location=format_group_ranges(before_group),
                affected_members=len(affected_ids),
                before_severity=before_group.severity.value,
                after_severities=tuple(
                    severity.value
                    for severity in Severity
                    if any(group.severity is severity for group in after_for_members)
                ),
                reasons=tuple(
                    sorted(
                        {
                            reason
                            for finding_id in affected_ids
                            for reason in reasons_by_finding[finding_id]
                        }
                    )
                ),
            )
        )
    return tuple(sorted(records, key=lambda item: item.before_group_id))


def preview_policy(
    findings: Sequence[Finding],
    bases: Mapping[str, NumericCounterfactualBasis],
    profile: DeliverableProfile | None,
    policy: CounterfactualPolicy,
) -> PolicyPreview:
    """Reclassify typed numeric evidence without mutating stored findings."""
    if len(findings) > POLICY_PREVIEW_MAX_FINDINGS:
        raise PolicyPreviewUnavailable(
            "What-if preview is unavailable for runs over "
            f"{POLICY_PREVIEW_MAX_FINDINGS:,} findings"
        )
    before = _engine_findings(findings, profile)
    after = _engine_findings(findings, profile)
    before_by_id = {finding.finding_id: finding for finding in before}
    after_by_id = {finding.finding_id: finding for finding in after}
    if len(before_by_id) != len(before):
        raise ValueError("finding ids are not unique")

    floor_tier = Materiality(policy.review_floor.value)
    floor_index = _MATERIALITY_ORDER.index(floor_tier)
    changed_ids: set[str] = set()
    reasons_by_finding: dict[str, set[str]] = {}
    accepted_atomics = 0
    review_floor_atomics = 0

    for finding_id in sorted(bases):
        basis = bases[finding_id]
        finding = after_by_id.get(finding_id)
        before_finding = before_by_id.get(finding_id)
        if finding is None or before_finding is None:
            raise ValueError(f"no finding with id {finding_id} for policy preview")
        if (
            finding.finding_class is not FindingClass.VALUE_CHANGED
            or finding.subtype is not FindingSubtype.VALUE_REPLACEMENT
            or finding.sheet != basis.sheet
            or finding.location != basis.location
        ):
            raise ValueError(f"counterfactual basis mismatch for finding {finding_id}")
        # Existing cadence/profile expectations already sit below the action
        # queue; a counterfactual acceptance must not make them less deferred.
        if before_finding.expected_reason is not None or before_finding.expected_growth:
            continue

        accepted = _within_acceptance(
            basis.baseline,
            basis.current,
            policy.acceptance,
        )
        finding.materiality = classify_numeric_pair(
            basis.baseline,
            basis.current,
            basis.number_format,
            within_acceptance=accepted,
        )
        reasons: set[str] = set()
        if accepted:
            finding.expected_reason = None
            finding.expected_growth = False
            reasons.add("within_acceptance")
        finding.severity = assign_severity(finding, profile)

        below_floor = (
            finding.materiality is not None
            and _MATERIALITY_ORDER.index(finding.materiality) < floor_index
        )
        if below_floor:
            finding.expected_reason = None
            finding.expected_growth = False
            finding.severity = Severity.INFO
            reasons.add("below_review_floor")

        if (
            finding.severity is not before_finding.severity
            or finding.materiality is not before_finding.materiality
        ):
            changed_ids.add(finding_id)
            reasons_by_finding[finding_id] = reasons
            accepted_atomics += int(accepted)
            review_floor_atomics += int(below_floor and not accepted)

    before_groups = build_pattern_groups(before)
    after_groups = build_pattern_groups(after)
    return PolicyPreview(
        before_atomic=_severity_counts(before),
        after_atomic=_severity_counts(after),
        before_decision=_decision_counts(before_groups),
        after_decision=_decision_counts(after_groups),
        affected_atomics=len(changed_ids),
        accepted_atomics=accepted_atomics,
        review_floor_atomics=review_floor_atomics,
        affected_decisions=_decision_records(
            before_groups,
            after_groups,
            changed_ids,
            reasons_by_finding,
        ),
    )
