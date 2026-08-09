"""Deterministic disagreement checks for repeated readable deck claims."""

from __future__ import annotations

import hashlib
import json
from collections import Counter, defaultdict
from dataclasses import dataclass

from qc_tool.coverage import CoverageItem, CoverageState
from qc_tool.crosscheck.claims import (
    ClaimIdentityAssessment,
    SemanticClaimIdentityV1,
    claim_population,
    classify_claim_identities,
)
from qc_tool.crosscheck.trace import extract_deck_figures
from qc_tool.findings import Finding, FindingClass
from qc_tool.ppt.model import DeckSnapshot


@dataclass(frozen=True, slots=True)
class RepeatedClaimResult:
    findings: tuple[Finding, ...]
    coverage: CoverageItem


def _identity_key(identity: SemanticClaimIdentityV1) -> tuple[object, ...]:
    return (
        identity.semantic_skeleton,
        identity.period_kind,
        identity.period_key,
        identity.unit_key,
        identity.scale,
        identity.decimals,
        identity.surface_kind,
        identity.surface_anchor,
        identity.occurrence_ordinal,
    )


def _display_tolerance(identity: SemanticClaimIdentityV1) -> float:
    half_unit = 0.5 * 10.0 ** (-identity.decimals)
    if identity.unit_key == "percent":
        return half_unit / 100.0 + 1e-12
    return half_unit * identity.scale + 1e-12


def _bounded_slides(members: list[ClaimIdentityAssessment]) -> str:
    names = sorted({item.occurrence.slide for item in members})
    rendered = ", ".join(names[:3])
    return rendered + (f"; +{len(names) - 3} more" if len(names) > 3 else "")


def check_repeated_claims(deck: DeckSnapshot) -> RepeatedClaimResult:
    occurrences = extract_deck_figures(deck)
    assessments = classify_claim_identities(deck, occurrences)
    population = claim_population(deck, occurrences, assessments)
    groups: dict[tuple[object, ...], list[ClaimIdentityAssessment]] = defaultdict(
        list
    )
    unavailable_reasons: Counter[str] = Counter()
    for assessment in assessments:
        if assessment.identity is None:
            unavailable_reasons.update(assessment.unavailable_reasons)
            continue
        groups[_identity_key(assessment.identity)].append(assessment)

    eligible = 0
    consistent = 0
    findings: list[Finding] = []
    for _key, group in sorted(groups.items()):
        members = sorted(
            group,
            key=lambda item: (
                item.occurrence.slide_index,
                item.occurrence.figure_index,
                item.occurrence.figure.raw,
            ),
        )
        slide_indices = sorted({item.occurrence.slide_index for item in members})
        if len(slide_indices) < 2:
            continue
        eligible += 1
        identity = members[0].identity
        if identity is None:  # narrowed by grouping above
            raise RuntimeError("repetition group lost its semantic identity")
        values = [item.occurrence.figure.value for item in members]
        if max(values) - min(values) <= _display_tolerance(identity):
            consistent += 1
            continue

        canonical = json.dumps(
            {
                "identity": identity.model_dump(mode="json"),
                "slides": slide_indices,
            },
            sort_keys=True,
            separators=(",", ":"),
        )
        digest = hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:12]
        raw_evidence = [
            f"slide {item.occurrence.slide_index}: {item.occurrence.figure.raw}"
            for item in members[:5]
        ]
        omitted = len(members) - len(raw_evidence)
        if omitted:
            raw_evidence.append(f"+{omitted} more occurrences")
        locations = "; ".join(f"slide {index}" for index in slide_indices)
        event_key = f"ppt-repetition:{digest}"
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_REPEATED_CLAIM_MISMATCH,
                slide=_bounded_slides(members),
                location=locations,
                element=identity.semantic_skeleton,
                baseline_value=raw_evidence[0],
                current_value="; ".join(raw_evidence[1:]),
                event_key=event_key,
                root_cause_key=event_key,
                message=(
                    "Repeated presentation claim disagrees beyond its common "
                    f"display precision on {locations}"
                ),
            )
        )

    limitation = ", ".join(
        f"{reason}={count}" for reason, count in sorted(unavailable_reasons.items())
    )
    details = [
        f"readable={population.readable}",
        f"identity_available={population.identity_available}",
        f"identity_unavailable={population.identity_unavailable}",
        f"eligible_groups={eligible}",
        f"consistent_groups={consistent}",
        f"mismatch_groups={len(findings)}",
        f"opaque_surfaces={population.unreadable_shapes}",
    ]
    if limitation:
        details.append(limitation)
    if not deck.charts_available:
        details.append("visible native chart labels unavailable")
    complete = population.identity_complete and deck.charts_available
    coverage = CoverageItem(
        check_id="ppt-internal-repetition",
        label="Repeated presentation claims",
        artifact="ppt",
        state=CoverageState.CHECKED if complete else CoverageState.DEGRADED,
        findings=len(findings),
        detail="; ".join(details),
    )
    return RepeatedClaimResult(findings=tuple(findings), coverage=coverage)
