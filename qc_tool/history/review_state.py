"""Versioned review evidence identity and lifecycle models."""

from __future__ import annotations

import hashlib
import json
from enum import Enum, StrEnum

from pydantic import BaseModel, Field

from qc_tool.findings import Finding

FINDING_EVIDENCE_VERSION = 1

#: Schema version of the versioned many-to-one `annotation_lineage` shape
#: (A4); embedded in attestation v4 payloads as `lineage_version`. Distinct
#: from `FINDING_EVIDENCE_VERSION`, which versions the evidence-digest
#: algorithm, not the lineage table's own shape.
ANNOTATION_LINEAGE_VERSION = 2


class RunFinalizedError(RuntimeError):
    """Raised when a signed run's review state would be mutated."""


class AnnotationLineageRelation(StrEnum):
    """How a lineage row's source finding relates to the current finding."""

    #: The current and source findings share the same identity (an atomic's
    #: location-based identity, or a population's shape-digest identity).
    IDENTITY = "identity"
    #: The source finding is one atomic member absorbed into a population
    #: that did not exist, in this shape, in the source run.
    MEMBER = "member"


class AnnotationLineageOutcome(StrEnum):
    """Why a lineage row's decision was -- or was not -- safe to carry."""

    #: Every source contributing to this decision agreed; carried cleanly.
    INHERITED = "inherited"
    #: Contributing sources disagreed on severity; not auto-applied.
    CONFLICT_SEVERITY = "conflict_severity"
    #: Severities agreed but comments differed; not auto-applied.
    CONFLICT_COMMENT = "conflict_comment"
    #: Only some of the population's members had a source decision.
    PARTIAL = "partial"
    #: The source run was never finalized; its decisions are provisional.
    UNFINALIZED_SOURCE = "unfinalized_source"


class AnnotationLineage(BaseModel):
    finding_id: str
    source_run_id: int
    source_finding_id: str
    relation: AnnotationLineageRelation
    outcome: AnnotationLineageOutcome
    evidence_version: int = FINDING_EVIDENCE_VERSION
    source_digest: str
    applied_at: str


class RunSignoff(BaseModel):
    run_id: int
    finalized_at: str
    acknowledgements: tuple[str, ...] = ()
    review_state_digest: str
    profile_sha256: str
    attestation_path: str
    attestation_sha256: str
    report_paths: dict[str, str] = Field(default_factory=dict)


class CarryForwardCandidate(BaseModel):
    finding_id: str
    source_finding_id: str
    severity: str | None = None
    comment: str = ""
    evidence_version: int = FINDING_EVIDENCE_VERSION
    evidence_digest: str


class PopulationLineageSource(BaseModel):
    """One atomic member's contribution to a population's carry-forward."""

    source_finding_id: str
    source_digest: str
    severity: str | None = None
    comment: str = ""


class PopulationCarryForwardCandidate(BaseModel):
    """A population's carry-forward disposition across all its members.

    Only ``outcome == inherited`` is auto-appliable (Criterion 8: a
    population accepts exactly one decision); every other outcome is
    disclosed for the analyst to review and decide fresh.
    """

    finding_id: str
    outcome: AnnotationLineageOutcome
    severity: str | None = None
    comment: str = ""
    member_count: int
    matched_member_count: int
    sources: tuple[PopulationLineageSource, ...] = ()


class CarryForwardPreview(BaseModel):
    source_run_id: int
    source_finalized: bool = False
    exact: tuple[CarryForwardCandidate, ...] = ()
    changed_evidence: tuple[str, ...] = ()
    resolved: tuple[str, ...] = ()
    ambiguous: tuple[str, ...] = ()
    populations: tuple[PopulationCarryForwardCandidate, ...] = ()
    #: Excel roles shared by both runs whose resolved formula engine differs
    #: (e.g. the native kernel became available/unavailable between runs).
    #: A non-empty tuple means evidence for those roles may not be directly
    #: comparable even where a digest happens to match or differ -- an
    #: explicit disclosure, not a block, since `exact`'s own digest equality
    #: check already stays conservative regardless.
    engine_provenance_mismatch: tuple[str, ...] = ()


class EvidenceFieldRole(StrEnum):
    INCLUDED = "included"
    NORMALIZED = "normalized"
    DERIVED = "derived"
    EXCLUDED = "excluded"


FINDING_EVIDENCE_FIELDS: dict[str, EvidenceFieldRole] = {
    "finding_id": EvidenceFieldRole.EXCLUDED,
    "artifact": EvidenceFieldRole.INCLUDED,
    "finding_class": EvidenceFieldRole.INCLUDED,
    "severity": EvidenceFieldRole.INCLUDED,
    "expected_growth": EvidenceFieldRole.DERIVED,
    "expected_reason": EvidenceFieldRole.INCLUDED,
    "provenance": EvidenceFieldRole.INCLUDED,
    "subtype": EvidenceFieldRole.INCLUDED,
    "materiality": EvidenceFieldRole.INCLUDED,
    "temporal_context": EvidenceFieldRole.INCLUDED,
    "evidence_tags": EvidenceFieldRole.NORMALIZED,
    "event_key": EvidenceFieldRole.INCLUDED,
    "sheet": EvidenceFieldRole.INCLUDED,
    "artifact_member": EvidenceFieldRole.DERIVED,
    "location": EvidenceFieldRole.INCLUDED,
    "baseline_location": EvidenceFieldRole.INCLUDED,
    "element": EvidenceFieldRole.INCLUDED,
    "slide": EvidenceFieldRole.INCLUDED,
    "slide_index": EvidenceFieldRole.INCLUDED,
    "baseline_slide_index": EvidenceFieldRole.INCLUDED,
    "focus_shape_id": EvidenceFieldRole.EXCLUDED,
    "baseline_focus_shape_id": EvidenceFieldRole.EXCLUDED,
    "baseline_value": EvidenceFieldRole.INCLUDED,
    "current_value": EvidenceFieldRole.INCLUDED,
    "message": EvidenceFieldRole.EXCLUDED,
    "impacts": EvidenceFieldRole.NORMALIZED,
    "baseline_excerpt": EvidenceFieldRole.EXCLUDED,
    "current_excerpt": EvidenceFieldRole.EXCLUDED,
    "analyst_comment": EvidenceFieldRole.EXCLUDED,
    "severity_overridden": EvidenceFieldRole.EXCLUDED,
    "root_cause_key": EvidenceFieldRole.INCLUDED,
    "waiver_reason": EvidenceFieldRole.INCLUDED,
    "waiver_expires": EvidenceFieldRole.INCLUDED,
    "counterfactual_basis": EvidenceFieldRole.EXCLUDED,
    "series_anchor": EvidenceFieldRole.EXCLUDED,
    "population": EvidenceFieldRole.INCLUDED,
}


def _json_value(value: object) -> object:
    if isinstance(value, Enum):
        return value.value
    if isinstance(value, BaseModel):
        return value.model_dump(mode="json")
    return value


def finding_evidence_payload(finding: Finding) -> dict[str, object]:
    """Canonical semantic evidence, excluding prose and analyst decisions."""
    classified = set(FINDING_EVIDENCE_FIELDS)
    actual = set(Finding.model_fields)
    if classified != actual:
        missing = sorted(actual - classified)
        extra = sorted(classified - actual)
        raise RuntimeError(
            f"finding evidence registry mismatch; missing={missing}; extra={extra}"
        )

    payload: dict[str, object] = {}
    for field_name, role in FINDING_EVIDENCE_FIELDS.items():
        if role is EvidenceFieldRole.INCLUDED:
            payload[field_name] = _json_value(getattr(finding, field_name))
        elif role is EvidenceFieldRole.NORMALIZED:
            values = getattr(finding, field_name)
            payload[field_name] = sorted(str(_json_value(value)) for value in values)
    payload["effective_expected"] = (
        finding.expected_reason.value
        if finding.expected_reason is not None
        else bool(finding.expected_growth)
    )
    if finding.artifact_member != "primary":
        payload["artifact_member"] = finding.artifact_member
    return {
        "version": FINDING_EVIDENCE_VERSION,
        "evidence": payload,
    }


def finding_evidence_digest(finding: Finding) -> str:
    encoded = json.dumps(
        finding_evidence_payload(finding),
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()
