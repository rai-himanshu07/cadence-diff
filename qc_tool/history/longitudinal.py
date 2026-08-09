"""Versioned identities and private views for longitudinal review evidence."""

from __future__ import annotations

import hashlib
import json
from enum import StrEnum
from typing import Literal

from pydantic import BaseModel, Field

from qc_tool.findings import Finding
from qc_tool.package import MEMBER_ID_PATTERN


class DecisionOrigin(StrEnum):
    MANUAL = "manual"
    CARRIED = "carried"
    LEGACY_MANUAL = "legacy_manual"


class DecisionAnchorV1(BaseModel):
    version: Literal[1] = 1
    artifact_member: str = Field(default="primary", pattern=MEMBER_ID_PATTERN)
    artifact: str
    finding_class: str
    sheet: str | None = None
    slide: str | None = None
    locator: str | None = None
    element: str | None = None
    event_key: str | None = None
    root_cause_key: str | None = None

    model_config = {"frozen": True}


class DecisionOccurrence(BaseModel):
    run_id: int
    finding_id: str
    contract_scope: str
    anchor_version: int
    anchor_digest: str
    evidence_version: int
    evidence_digest: str

    model_config = {"frozen": True}


class DossierStatus(StrEnum):
    EXACT = "exact"
    CHANGED = "changed"
    AMBIGUOUS = "ambiguous"
    NO_STORED_OBSERVATION = "no_stored_observation"


class DossierEntry(BaseModel):
    run_id: int
    started_at: str
    status: DossierStatus
    finalized: bool
    finding_id: str | None = None
    engine_severity: str | None = None
    analyst_severity: str | None = None
    analyst_comment: str = ""
    origin: DecisionOrigin | None = None
    carry_source_run_id: int | None = None
    carry_source_finding_id: str | None = None
    baseline_value: str | None = None
    current_value: str | None = None
    waiver_reason: str = ""
    waiver_expires: str = ""
    promotion_kind: str | None = None
    promotion_profile_sha256: str | None = None
    promoted_at: str | None = None

    model_config = {"frozen": True}


class DossierResult(BaseModel):
    contract_scope: str
    anchor_digest: str
    entries: tuple[DossierEntry, ...] = ()

    model_config = {"frozen": True}


class RecurrenceEligibility(BaseModel):
    run_ids: tuple[int, ...]
    manual_count: int
    analyst_severity: str

    model_config = {"frozen": True}


def decision_anchor(
    finding: Finding,
    *,
    artifact_member: str | None = None,
) -> DecisionAnchorV1:
    """Build the additive anchor without changing frozen finding identity."""
    return DecisionAnchorV1(
        artifact_member=artifact_member or finding.artifact_member,
        artifact=finding.artifact,
        finding_class=finding.finding_class.value,
        sheet=finding.sheet or None,
        slide=finding.slide or None,
        locator=finding.location or finding.baseline_location or None,
        element=finding.element or None,
        event_key=finding.event_key or None,
        root_cause_key=finding.root_cause_key or None,
    )


def canonical_anchor_digest(anchor: DecisionAnchorV1) -> str:
    payload = json.dumps(
        anchor.model_dump(mode="json"),
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def contract_scope_for_profile(
    profile_name: str | None,
    contract_id: str | None,
) -> str:
    """Scope by portable contract id, or explicitly by legacy profile name."""
    if contract_id:
        return f"contract:{contract_id}"
    normalized = (profile_name or "").strip().casefold()
    return f"legacy-profile:{normalized}"
