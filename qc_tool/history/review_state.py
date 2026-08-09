"""Versioned review evidence identity and lifecycle models."""

from __future__ import annotations

import hashlib
import json
from enum import Enum, StrEnum

from pydantic import BaseModel, Field

from qc_tool.findings import Finding

FINDING_EVIDENCE_VERSION = 1


class RunFinalizedError(RuntimeError):
    """Raised when a signed run's review state would be mutated."""


class AnnotationLineage(BaseModel):
    finding_id: str
    source_run_id: int
    source_finding_id: str
    evidence_version: int = FINDING_EVIDENCE_VERSION
    evidence_digest: str
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


class CarryForwardPreview(BaseModel):
    source_run_id: int
    source_finalized: bool = False
    exact: tuple[CarryForwardCandidate, ...] = ()
    changed_evidence: tuple[str, ...] = ()
    resolved: tuple[str, ...] = ()
    ambiguous: tuple[str, ...] = ()


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
}


def _json_value(value: object) -> object:
    if isinstance(value, Enum):
        return value.value
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
