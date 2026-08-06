"""Explicit evidence-gated analyst decision carry-forward."""

from __future__ import annotations

from collections import defaultdict

from qc_tool.history.review_state import (
    CarryForwardCandidate,
    CarryForwardPreview,
    finding_evidence_digest,
)
from qc_tool.history.store import RunHistory
from qc_tool.review import finding_identity_key


def preview_carry_forward(history: RunHistory, run_id: int) -> CarryForwardPreview:
    current = history.get_raw_run(run_id)
    if current.rerun_of is None:
        raise ValueError("run is not a Re-QC run")
    source = history.get_raw_run(current.rerun_of)
    annotations = history.get_annotations(source.run_id)
    source_by_id = {finding.finding_id: finding for finding in source.findings}
    current_by_identity: dict[tuple[str, ...], list] = defaultdict(list)
    source_by_identity: dict[tuple[str, ...], list] = defaultdict(list)
    for finding in current.findings:
        current_by_identity[finding_identity_key(finding)].append(finding)
    for finding in source.findings:
        source_by_identity[finding_identity_key(finding)].append(finding)

    exact: list[CarryForwardCandidate] = []
    changed: list[str] = []
    resolved: list[str] = []
    ambiguous: list[str] = []
    for source_finding_id, (severity, comment) in sorted(annotations.items()):
        source_finding = source_by_id.get(source_finding_id)
        if source_finding is None:
            ambiguous.append(source_finding_id)
            continue
        identity = finding_identity_key(source_finding)
        source_matches = source_by_identity[identity]
        current_matches = current_by_identity.get(identity, [])
        if len(source_matches) != 1 or len(current_matches) > 1:
            ambiguous.append(source_finding_id)
            continue
        if not current_matches:
            resolved.append(source_finding_id)
            continue
        current_finding = current_matches[0]
        source_digest = finding_evidence_digest(source_finding)
        if finding_evidence_digest(current_finding) != source_digest:
            changed.append(current_finding.finding_id)
            continue
        exact.append(
            CarryForwardCandidate(
                finding_id=current_finding.finding_id,
                source_finding_id=source_finding_id,
                severity=severity,
                comment=comment,
                evidence_digest=source_digest,
            )
        )
    return CarryForwardPreview(
        source_run_id=source.run_id,
        source_finalized=source.signoff is not None,
        exact=tuple(sorted(exact, key=lambda item: item.finding_id)),
        changed_evidence=tuple(sorted(changed)),
        resolved=tuple(sorted(resolved)),
        ambiguous=tuple(sorted(ambiguous)),
    )


def apply_carry_forward(
    history: RunHistory,
    run_id: int,
    selected_finding_ids: set[str],
) -> int:
    """Recompute eligibility, then atomically apply only the selected exact set."""
    preview = preview_carry_forward(history, run_id)
    by_id = {candidate.finding_id: candidate for candidate in preview.exact}
    unknown = selected_finding_ids - by_id.keys()
    if unknown:
        raise ValueError("one or more selected findings are not exact candidates")
    selected = [by_id[finding_id] for finding_id in sorted(selected_finding_ids)]
    return history.apply_carried_annotations(run_id, preview.source_run_id, selected)
