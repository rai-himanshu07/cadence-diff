"""Explicit evidence-gated analyst decision carry-forward."""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Sequence

from qc_tool.findings import Finding
from qc_tool.history.config_compatibility import (
    ConfigurationCompatibility,
    configuration_compatible,
)
from qc_tool.history.review_state import (
    AnnotationLineageOutcome,
    CarryForwardCandidate,
    CarryForwardPreview,
    PopulationCarryForwardCandidate,
    PopulationLineageSource,
    finding_evidence_digest,
)
from qc_tool.history.store import RunHistory
from qc_tool.review import (
    population_members_iter,
    requeue_identity_key,
)


def _source_atomics_by_location(
    source_findings: Sequence[Finding], annotated_ids: set[str]
) -> dict[tuple[str, str, str], dict[str, Finding]]:
    """Index annotated source atomics for population member look-up.

    Keyed on `(artifact, finding_class, sheet)` -> `{location: Finding}` --
    unannotated atomics carry nothing to inherit, so they are excluded up
    front rather than filtered per population.
    """
    index: dict[tuple[str, str, str], dict[str, Finding]] = defaultdict(dict)
    for finding in source_findings:
        if finding.population is not None or finding.finding_id not in annotated_ids:
            continue
        if finding.location is None:
            continue
        key = (finding.artifact, finding.finding_class.value, finding.sheet or "")
        index[key][finding.location] = finding
    return index


def _population_candidate(
    finding: Finding,
    *,
    atomics_by_location: dict[tuple[str, str, str], dict[str, Finding]],
    annotations: dict[str, tuple[str | None, str]],
    source_finalized: bool,
) -> PopulationCarryForwardCandidate | None:
    """Classify one current population's carry-forward disposition, or
    ``None`` when none of its members were previously annotated.

    Streams members via `population_members_iter` (O(1) extra memory
    regardless of population size) instead of materializing every member
    up front; `member_count` comes directly from the typed evidence field,
    never from `len()` of a decoded list (Criterion 9).
    """
    if finding.population is None:
        return None
    member_count = finding.population.member_count
    if member_count == 0:
        return None
    key = (finding.artifact, finding.finding_class.value, finding.sheet or "")
    by_location = atomics_by_location.get(key, {})
    sources: list[PopulationLineageSource] = []
    for current_location, _baseline_location in population_members_iter(finding):
        source_finding = by_location.get(current_location)
        if source_finding is None:
            continue
        severity, comment = annotations[source_finding.finding_id]
        sources.append(
            PopulationLineageSource(
                source_finding_id=source_finding.finding_id,
                source_digest=finding_evidence_digest(source_finding),
                severity=severity,
                comment=comment,
            )
        )
    if not sources:
        return None
    if not source_finalized:
        outcome = AnnotationLineageOutcome.UNFINALIZED_SOURCE
    elif len(sources) < member_count:
        outcome = AnnotationLineageOutcome.PARTIAL
    elif len({source.severity for source in sources}) > 1:
        outcome = AnnotationLineageOutcome.CONFLICT_SEVERITY
    elif len({source.comment for source in sources}) > 1:
        outcome = AnnotationLineageOutcome.CONFLICT_COMMENT
    else:
        outcome = AnnotationLineageOutcome.INHERITED
    inherited = outcome is AnnotationLineageOutcome.INHERITED
    return PopulationCarryForwardCandidate(
        finding_id=finding.finding_id,
        outcome=outcome,
        severity=sources[0].severity if inherited else None,
        comment=sources[0].comment if inherited else "",
        member_count=member_count,
        matched_member_count=len(sources),
        sources=tuple(sources),
    )


def preview_carry_forward(history: RunHistory, run_id: int) -> CarryForwardPreview:
    """Classify every annotated source finding's carry-forward disposition.

    plan-20260913 Step 4: a source (or, for populations, current) finding
    whose logical scope's comparison policy differs between the two runs
    is routed to ``ambiguous`` (atomics) or dropped (populations) rather
    than ``exact``/``resolved`` -- a policy change can make a finding
    disappear or look identical by pure identity/evidence-digest matching
    without anything in the underlying workbook actually changing.
    """
    current = history.get_raw_run(run_id)
    if current.rerun_of is None:
        raise ValueError("run is not a Re-QC run")
    source = history.get_raw_run(current.rerun_of)
    annotations = history.get_annotations(source.run_id)
    source_by_id = {finding.finding_id: finding for finding in source.findings}
    current_by_identity: dict[tuple[str, ...], list] = defaultdict(list)
    source_by_identity: dict[tuple[str, ...], list] = defaultdict(list)
    for finding in current.findings:
        current_by_identity[requeue_identity_key(finding)].append(finding)
    for finding in source.findings:
        source_by_identity[requeue_identity_key(finding)].append(finding)

    #: plan-20260913 Step 4: a scope whose comparison policy changed between
    #: the source and current run cannot safely resolve/carry an
    #: annotation -- a suppressed (not fixed) finding must never look
    #: "resolved", and an identical-looking match cannot be trusted "exact"
    #: either. `None` when either run predates `profile_snapshot`
    #: persistence -- legacy behavior, unchanged.
    compatibility: ConfigurationCompatibility | None = None
    if source.profile_snapshot is not None and current.profile_snapshot is not None:
        compatibility = configuration_compatible(
            source.profile_snapshot, current.profile_snapshot
        )

    exact: list[CarryForwardCandidate] = []
    changed: list[str] = []
    resolved: list[str] = []
    ambiguous: list[str] = []
    for source_finding_id, (severity, comment) in sorted(annotations.items()):
        source_finding = source_by_id.get(source_finding_id)
        if source_finding is None:
            ambiguous.append(source_finding_id)
            continue
        if compatibility is not None and not compatibility.comparable(source_finding):
            ambiguous.append(source_finding_id)
            continue
        identity = requeue_identity_key(source_finding)
        source_matches = source_by_identity[identity]
        current_matches = current_by_identity.get(identity, [])
        if len(source_matches) != 1 or len(current_matches) > 1:
            ambiguous.append(source_finding_id)
            continue
        if not current_matches:
            if source_finding.population is not None:
                # A population's decision cannot map onto plain atomics if
                # the group-first policy was later disabled; disclose this
                # as unresolved rather than silently treating it as fixed.
                ambiguous.append(source_finding_id)
            else:
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

    identity_handled = {candidate.finding_id for candidate in exact}
    identity_handled.update(changed)
    atomics_by_location = _source_atomics_by_location(
        source.findings, set(annotations)
    )
    source_finalized = source.signoff is not None
    populations: list[PopulationCarryForwardCandidate] = []
    for finding in current.findings:
        if finding.population is None or finding.finding_id in identity_handled:
            continue
        if compatibility is not None and not compatibility.comparable(finding):
            continue  # scope's comparison policy changed -- nothing safe to carry
        if source_by_identity.get(requeue_identity_key(finding)):
            continue  # an unannotated source population; nothing to carry
        candidate = _population_candidate(
            finding,
            atomics_by_location=atomics_by_location,
            annotations=annotations,
            source_finalized=source_finalized,
        )
        if candidate is not None:
            populations.append(candidate)

    return CarryForwardPreview(
        source_run_id=source.run_id,
        source_finalized=source_finalized,
        exact=tuple(sorted(exact, key=lambda item: item.finding_id)),
        changed_evidence=tuple(sorted(changed)),
        resolved=tuple(sorted(resolved)),
        ambiguous=tuple(sorted(ambiguous)),
        populations=tuple(
            sorted(populations, key=lambda item: item.finding_id)
        ),
        engine_provenance_mismatch=tuple(
            sorted(
                {
                    role
                    for source_engines, current_engines in (
                        (source.formula_engines, current.formula_engines),
                        (source.values_engines, current.values_engines),
                    )
                    for role in set(source_engines) & set(current_engines)
                    if source_engines[role] != current_engines[role]
                }
            )
        ),
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


def apply_population_carry_forward(
    history: RunHistory,
    run_id: int,
    finding_id: str,
) -> int:
    """Recompute eligibility, then apply one population's inherited decision.

    Only the single named population is applied per call -- unlike atomic
    carry-forward's batch selection -- since each population is its own,
    independent, single decision (Criterion 8).
    """
    preview = preview_carry_forward(history, run_id)
    by_id = {
        candidate.finding_id: candidate
        for candidate in preview.populations
        if candidate.outcome is AnnotationLineageOutcome.INHERITED
    }
    candidate = by_id.get(finding_id)
    if candidate is None:
        raise ValueError(
            f"population {finding_id!r} is not an inherited carry-forward candidate"
        )
    return history.apply_population_carry_forward(run_id, preview.source_run_id, candidate)

