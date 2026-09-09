"""Group-first population candidates: spill, grouping, and emission.

Producers (``_paired_cell_findings`` in ``qc_tool.excel.formulas``,
``iter_region_value_findings`` in ``qc_tool.excel.diff_values``) build an
ordinary ``Finding`` for eligible classes exactly as they always have, but
when the run's ``review_policy.populations`` is enabled they call
``CandidateSpill.add`` instead of appending/yielding it -- the default
(``sink=None``) keeps every existing call site byte-identical. A validated
``Finding(...)`` is not the expensive part of the atomic pipeline (measured:
constructing one costs a few microseconds); the cost this plan removes is
the per-cell impact/chart-impact/excerpt enrichment (`_enrich_retained_
findings`) that a population no longer needs for ~99% of its members.

``CandidateSpill`` reuses the same block-compressed spill codec
(``qc_tool.findings_store``) the findings store already uses, keyed on the
population identity/grouping key (``qc_tool.review.population_key``) instead
of triage order, so members of one population land contiguous once
``merge_spill`` sorts them. Reading candidates back uses ``Finding.
from_trusted_payload`` (the codec's own established fast path for
self-written JSON payloads), never ``model_validate``. ``finalize_
populations`` groups the sorted spill workbook-globally and, per group,
either builds one population ``Finding`` (within the policy's
rectangle/pair caps, at or above its threshold) or returns the group's
members for atomic replay through the unchanged pipeline -- the mechanism
that makes exact fallback-digest parity (Criterion 3b) provable by
construction rather than by a separate code path.
"""

from __future__ import annotations

import contextlib
import datetime as dt
import shutil
import tempfile
from collections import defaultdict
from collections.abc import Iterator
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from qc_tool.config.profile import DeliverableProfile, PopulationPolicy
from qc_tool.findings import (
    Finding,
    FindingClass,
    MembershipCodec,
    PopulationEvidence,
    PopulationSample,
)
from qc_tool.findings_store import (
    FindingsStoreError,
    SpillWriter,
    finding_payload,
    merge_spill,
)
from qc_tool.review import (
    Coordinate,
    _bounding_range,
    _coordinate,
    _range,
    _rectangles,
    population_identity_digest,
    population_key,
)
from qc_tool.scope import ComparisonScope
from qc_tool.triage.rules import assign_severities

_MAX_SAMPLES = 5

_Member = tuple[Finding, Coordinate, Coordinate | None]


class CandidateSpill:
    """Workbook-global spill of group-first population candidates.

    A candidate is an ordinary, fully validated ``Finding`` -- constructing
    one is cheap, so there is no reason to skip validation for it. Its
    severity and waiver state are assigned here, once, by the same pure
    projection (``assign_severities``) the atomic pipeline always used, so
    the population key computed from it -- and any later atomic replay --
    match atomic-mode behavior exactly.
    """

    def __init__(self, profile: DeliverableProfile, today: dt.date) -> None:
        self._profile = profile
        self._today = today
        self._dir = Path(tempfile.mkdtemp(prefix="qc-population-"))
        self._spill = SpillWriter(self._dir / "candidates.qcfb")
        self._spill.__enter__()
        self._closed = False
        self.count = 0

    def add(self, candidate: Finding, *, shape_before: str, shape_after: str) -> None:
        assign_severities([candidate], self._profile, today=self._today)
        key = population_key(candidate, (shape_before, shape_after))
        coordinate = _coordinate(candidate.location) or (0, 0)
        sort_key = (*key, coordinate[0], coordinate[1])
        payload: dict[str, Any] = {
            "finding": finding_payload(candidate),
            "shape_before": shape_before,
            "shape_after": shape_after,
        }
        self._spill.append(sort_key, payload)
        self.count += 1

    def groups(self) -> Iterator[list[dict[str, Any]]]:
        """Close the spill and yield contiguous same-key candidate groups.

        Consumes (and removes) the spill directory; call at most once.
        """
        if not self._closed:
            self._spill.__exit__(None, None, None)
            self._closed = True
        try:
            current_key: tuple[object, ...] | None = None
            bucket: list[dict[str, Any]] = []
            for row in merge_spill(self._spill.path):
                if not isinstance(row, dict):
                    raise FindingsStoreError("population candidate payload is not a mapping")
                payload: dict[str, Any] = row
                finding = Finding.from_trusted_payload(payload["finding"])
                key = population_key(
                    finding, (payload["shape_before"], payload["shape_after"])
                )
                if current_key is not None and key != current_key:
                    yield bucket
                    bucket = []
                current_key = key
                bucket.append(payload)
            if bucket:
                yield bucket
        finally:
            shutil.rmtree(self._dir, ignore_errors=True)

    def abort(self) -> None:
        if self._closed:
            return
        with contextlib.suppress(Exception):
            self._spill.__exit__(None, None, None)
        shutil.rmtree(self._dir, ignore_errors=True)
        self._closed = True

    def __del__(self) -> None:
        with contextlib.suppress(Exception):
            self.abort()


@dataclass(slots=True)
class ClassPopulationStats:
    """Per-class disclosure counters for the population coverage row."""

    candidates: int = 0
    populations: int = 0
    replayed: int = 0
    below_threshold: int = 0
    over_cap: int = 0
    #: Candidates that met every geometry/threshold gate but disagreed on
    #: producer-authored evidence (`event_key`/`evidence_tags`) and so
    #: replayed atomically rather than silently losing or misattributing
    #: that evidence to one representative member.
    heterogeneous_evidence: int = 0


@dataclass(slots=True)
class PopulationOutcome:
    population_findings: list[Finding] = field(default_factory=list)
    replay_findings: list[Finding] = field(default_factory=list)
    stats: dict[FindingClass, ClassPopulationStats] = field(default_factory=dict)


def _homogeneous_evidence(findings: list[Finding]) -> bool:
    """Whether every candidate shares identical producer-authored evidence.

    ``population_key`` deliberately excludes ``evidence_tags`` (and only
    conditionally includes ``event_key``) from its grouping key, so members
    sharing a key can still disagree on producer-authored evidence. Such a
    group must replay atomically rather than have that evidence silently
    dropped or misattributed to one representative member.
    """
    first = findings[0]
    return all(
        finding.event_key == first.event_key
        and finding.evidence_tags == first.evidence_tags
        for finding in findings[1:]
    )


def finalize_populations(
    spill: CandidateSpill,
    policy: PopulationPolicy,
    scope: ComparisonScope,
) -> PopulationOutcome:
    """Group spilled candidates and decide population vs. atomic replay."""
    population_findings: list[Finding] = []
    replay_findings: list[Finding] = []
    stats: dict[FindingClass, ClassPopulationStats] = defaultdict(ClassPopulationStats)
    for group in spill.groups():
        findings = [Finding.from_trusted_payload(row["finding"]) for row in group]
        findings = scope.filter_findings(findings)
        if not findings:
            continue
        representative = findings[0]
        class_stats = stats[representative.finding_class]
        class_stats.candidates += len(findings)
        members: list[_Member] = [
            (finding, current, _coordinate(finding.baseline_location))
            for finding in findings
            if (current := _coordinate(finding.location)) is not None
        ]
        if len(members) != len(findings) or len(members) < policy.threshold:
            class_stats.below_threshold += len(findings)
            class_stats.replayed += len(findings)
            replay_findings.extend(findings)
            continue
        if not _homogeneous_evidence(findings):
            class_stats.heterogeneous_evidence += len(findings)
            class_stats.replayed += len(findings)
            replay_findings.extend(findings)
            continue
        membership = _build_membership(members, policy)
        if membership is None:
            class_stats.over_cap += len(findings)
            class_stats.replayed += len(findings)
            replay_findings.extend(findings)
            continue
        shape_before = group[0]["shape_before"]
        shape_after = group[0]["shape_after"]
        population_findings.append(
            _build_population(representative, members, membership, shape_before, shape_after)
        )
        class_stats.populations += 1
    return PopulationOutcome(population_findings, replay_findings, dict(stats))


def _build_membership(
    members: list[_Member], policy: PopulationPolicy
) -> MembershipCodec | None:
    """Build the dual-sided membership codec, or None if it exceeds caps."""
    current_coords = {current for _, current, _ in members}
    rectangles = _rectangles(current_coords)
    if len(rectangles) > policy.max_rectangles:
        return None
    current_rectangles = tuple(_range(rectangle) for rectangle in rectangles)
    offsets = {
        (baseline[0] - current[0], baseline[1] - current[1])
        for _, current, baseline in members
        if baseline is not None
    }
    if len(offsets) == 1:
        return MembershipCodec(
            current_rectangles=current_rectangles,
            baseline_mode="shift",
            shift=next(iter(offsets)),
            member_count=len(members),
        )
    pairs: list[tuple[str, str]] = []
    for finding, _current, baseline in members:
        if (
            baseline is not None
            and finding.location is not None
            and finding.baseline_location is not None
        ):
            pairs.append((finding.location, finding.baseline_location))
    if len(pairs) != len(members) or len(pairs) > policy.max_explicit_pairs:
        return None
    return MembershipCodec(
        current_rectangles=current_rectangles,
        baseline_mode="pairs",
        pairs=tuple(pairs),
        member_count=len(members),
    )


def _build_population(
    representative: Finding,
    members: list[_Member],
    membership: MembershipCodec,
    shape_before: str,
    shape_after: str,
) -> Finding:
    ordered = sorted(members, key=lambda member: member[1])
    current_coords = {current for _, current, _ in members}
    baseline_coords = {
        baseline for _, _current, baseline in members if baseline is not None
    }
    samples = tuple(
        PopulationSample(
            current_location=finding.location,  # type: ignore[arg-type]
            baseline_location=finding.baseline_location,
        )
        for finding, _current, _baseline in ordered[:_MAX_SAMPLES]
    )
    population = PopulationEvidence(
        member_count=len(members),
        membership=membership,
        first=ordered[0][0].location,  # type: ignore[arg-type]
        last=ordered[-1][0].location,  # type: ignore[arg-type]
        samples=samples,
        shape_before_digest=shape_before,
        shape_after_digest=shape_after,
    )
    finding = Finding(
        artifact=representative.artifact,
        artifact_member=representative.artifact_member,
        finding_class=representative.finding_class,
        severity=representative.severity,
        expected_reason=representative.expected_reason,
        provenance=representative.provenance,
        subtype=representative.subtype,
        materiality=representative.materiality,
        temporal_context=representative.temporal_context,
        evidence_tags=representative.evidence_tags,
        event_key=representative.event_key,
        sheet=representative.sheet,
        location=_bounding_range(current_coords),
        baseline_location=_bounding_range(baseline_coords),
        element="population",
        message=(
            f"{representative.sheet}: {len(members)} cells share one "
            f"{representative.finding_class.value} population"
        ),
        waiver_reason=representative.waiver_reason,
        waiver_expires=representative.waiver_expires,
        population=population,
    )
    finding.root_cause_key = "population:" + population_identity_digest(finding)
    return finding
