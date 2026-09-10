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
import time
from collections import defaultdict
from collections.abc import Iterator
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, NamedTuple

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


class _MemberFacts(NamedTuple):
    """Per-candidate facts needed to build membership/samples for an
    eligible population, without constructing a full ``Finding`` (plan-
    20260910, Step 6). ``location``/``baseline_location`` are kept as their
    original strings (not re-derived from ``current``/``baseline``) so a
    round-trip can never lose formatting a coordinate tuple would drop
    (e.g. an absolute ``$`` reference, if one were ever present).
    """

    location: str
    baseline_location: str | None
    current: Coordinate
    baseline: Coordinate | None


@dataclass(slots=True)
class PopulationTelemetry:
    """Aggregate-only timers for the group-first candidate pipeline
    (plan-20260910): construction (``CandidateSpill.add``), spill/merge
    (``CandidateSpill.groups``), and finalize (``finalize_populations``'s own
    grouping/decision loop). No candidate content -- formula, sheet, value,
    or coordinate -- is ever recorded, only elapsed seconds.
    """

    construction_seconds: float = 0.0
    spill_seconds: float = 0.0
    finalize_seconds: float = 0.0


#: Internal spill row schema version (plan-20260910, Step 6) -- never
#: persisted past one run (the spill directory is a disposable tempdir), so
#: this exists for code clarity/future changes, not cross-run compatibility.
_CANDIDATE_ROW_VERSION = 1

_MISSING = object()


def _diff_against_template(
    template: dict[str, Any], payload: dict[str, Any]
) -> dict[str, Any] | None:
    """Fields of ``payload`` that differ from ``template``, generically --
    never a hand-picked field list, so no candidate producer's field can be
    silently dropped. ``population_key`` already guarantees roughly a dozen
    fields (artifact/sheet/finding_class/severity/provenance/subtype/
    materiality/temporal_context/waiver state/...) are identical across
    every member of one group; this is what makes the resulting delta small
    for the common case, without this function needing to know which fields
    those are.

    Returns ``None`` (meaning: store the full payload instead) if
    ``template`` carries a key ``payload`` altogether lacks -- no known
    producer creates this (a template's keys are always a subset of a later
    same-key candidate's own keys), but falling back keeps reconstruction
    correct regardless of any future producer's shape.
    """
    if not template.keys() <= payload.keys():
        return None
    return {
        field: value
        for field, value in payload.items()
        if template.get(field, _MISSING) != value
    }


def _reconstruct_group_findings(group: list[dict[str, Any]]) -> list[Finding]:
    """Rebuild every row's full ``Finding`` from its (possibly delta-encoded)
    spill row, using the group's own template row (the first ``add()`` call
    for this population key) to fill in whatever a delta row omitted.

    A group always contains at least one ``"full"`` row: the very first
    ``add()`` call for a given key has no template to diff against yet, and
    ``_diff_against_template`` itself falls back to ``"full"`` if it ever
    can't safely diff. Finding it once per group (not once per row) keeps
    this linear in group size, not quadratic.
    """
    template: dict[str, Any] | None = None
    for candidate in group:
        if candidate["row"]["kind"] == "full":
            template = candidate["row"]["finding"]
            break
    findings: list[Finding] = []
    for row in group:
        inner = row["row"]
        if inner["kind"] == "full":
            findings.append(Finding.from_trusted_payload(inner["finding"]))
        else:
            assert template is not None, "a delta row implies its group has a full row"
            findings.append(Finding.from_trusted_payload({**template, **inner["finding"]}))
    return findings


class CandidateSpill:
    """Workbook-global spill of group-first population candidates.

    A candidate is an ordinary, fully validated ``Finding`` -- constructing
    one is cheap, so there is no reason to skip validation for it. Its
    severity and waiver state are assigned here, once, by the same pure
    projection (``assign_severities``) the atomic pipeline always used, so
    the population key computed from it -- and any later atomic replay --
    match atomic-mode behavior exactly.

    Spill rows are delta-encoded against their group's own first-seen
    candidate (plan-20260910, Step 6): every field ``population_key``
    already guarantees homogeneous across a group -- artifact, sheet,
    finding_class, severity, provenance, subtype, materiality,
    temporal_context, waiver state, and more -- is written once per group,
    not once per member, shrinking the dominant per-candidate spill/merge
    cost for the common (large, homogeneous) population case. This is an
    internal storage optimization only: ``finalize_populations`` still
    reconstructs one complete, ordinary ``Finding`` per candidate before any
    grouping/threshold/homogeneity/membership decision runs, so every
    downstream behavior -- including a heterogeneous or below-threshold
    group's exact atomic replay -- is unaffected by this encoding.
    """

    def __init__(
        self,
        profile: DeliverableProfile,
        today: dt.date,
        *,
        telemetry: PopulationTelemetry | None = None,
    ) -> None:
        self._profile = profile
        self._today = today
        self._dir = Path(tempfile.mkdtemp(prefix="qc-population-"))
        self._spill = SpillWriter(self._dir / "candidates.qcfb")
        self._spill.__enter__()
        self._closed = False
        self.count = 0
        self._telemetry = telemetry
        self._templates: dict[tuple[object, ...], dict[str, Any]] = {}

    def add(self, candidate: Finding, *, shape_before: str, shape_after: str) -> None:
        start = time.perf_counter() if self._telemetry is not None else 0.0
        assign_severities([candidate], self._profile, today=self._today)
        key = population_key(candidate, (shape_before, shape_after))
        coordinate = _coordinate(candidate.location) or (0, 0)
        sort_key = (*key, coordinate[0], coordinate[1])
        full_payload = finding_payload(candidate)
        template = self._templates.get(key)
        if template is None:
            self._templates[key] = full_payload
            row: dict[str, Any] = {"kind": "full", "finding": full_payload}
        else:
            delta = _diff_against_template(template, full_payload)
            row = (
                {"kind": "delta", "finding": delta}
                if delta is not None
                else {"kind": "full", "finding": full_payload}
            )
        payload: dict[str, Any] = {
            "v": _CANDIDATE_ROW_VERSION,
            "group_key": list(key),
            "row": row,
            "shape_before": shape_before,
            "shape_after": shape_after,
        }
        self._spill.append(sort_key, payload)
        self.count += 1
        if self._telemetry is not None:
            self._telemetry.construction_seconds += time.perf_counter() - start

    def groups(self) -> Iterator[list[dict[str, Any]]]:
        """Close the spill and yield contiguous same-key candidate groups.

        Consumes (and removes) the spill directory; call at most once.
        Boundaries come directly from each row's own stored ``group_key``
        (computed once at ``add()`` time) rather than recomputing
        ``population_key`` from a reconstructed ``Finding`` -- a delta row's
        own payload deliberately omits most of what ``population_key`` reads,
        so recomputing it here would require reconstructing every row before
        grouping could even begin.
        """
        if not self._closed:
            self._spill.__exit__(None, None, None)
            self._closed = True
        try:
            current_key: list[object] | None = None
            bucket: list[dict[str, Any]] = []
            for row in merge_spill(self._spill.path):
                if not isinstance(row, dict):
                    raise FindingsStoreError("population candidate payload is not a mapping")
                payload: dict[str, Any] = row
                key = payload.get("group_key")
                if not isinstance(key, list):
                    raise FindingsStoreError(
                        "population candidate payload is missing its group key"
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


def _row_facts(row: dict[str, Any], template: dict[str, Any]) -> _MemberFacts | None:
    """(location, baseline_location, current coordinate, baseline
    coordinate) for one spill row, without constructing a ``Finding``.
    ``None`` when the row's location does not parse to a single-cell
    coordinate (``_coordinate``'s own exact rule) -- not eligible as a
    population member, matching the original per-``Finding`` filter exactly.
    """
    inner = row["row"]["finding"]
    location = inner["location"] if "location" in inner else template.get("location")
    current = _coordinate(location)
    if current is None or location is None:
        return None
    baseline_location = (
        inner["baseline_location"]
        if "baseline_location" in inner
        else template.get("baseline_location")
    )
    return _MemberFacts(location, baseline_location, current, _coordinate(baseline_location))


def _row_evidence(row: dict[str, Any], template: dict[str, Any]) -> tuple[str, frozenset[str]]:
    """(event_key, evidence_tags) for one spill row, without constructing a
    ``Finding``. ``population_key`` deliberately excludes ``evidence_tags``
    (and only conditionally includes ``event_key``) from its grouping key,
    so members sharing a key can still disagree on producer-authored
    evidence -- this is exactly what the homogeneity check compares.
    """
    inner = row["row"]["finding"]
    event_key = inner["event_key"] if "event_key" in inner else template.get("event_key", "")
    tags = inner["evidence_tags"] if "evidence_tags" in inner else template.get("evidence_tags", [])
    return event_key, frozenset(tags)


def finalize_populations(
    spill: CandidateSpill,
    policy: PopulationPolicy,
    scope: ComparisonScope,
    *,
    telemetry: PopulationTelemetry | None = None,
) -> PopulationOutcome:
    """Group spilled candidates and decide population vs. atomic replay.

    A full ``Finding`` is reconstructed for every member ONLY when a group
    turns out to need atomic replay (below threshold, heterogeneous
    evidence, or over the rectangle/pair cap) -- exactly where every
    member's complete evidence is genuinely needed. For a group that
    becomes one population, only the representative (the group's template
    row -- population_key already guarantees every member shares its
    artifact/sheet/finding_class/severity/provenance/subtype/materiality/
    temporal_context/waiver state) and up to 5 samples are ever fully
    reconstructed, however large the group (plan-20260910, Step 6). Scope
    filtering is likewise decided once per group from the representative
    alone: every field ``ComparisonScope.filter_findings`` reads (artifact,
    artifact_member, sheet, slide_index) is population_key-guaranteed
    identical across a group, so it can never partially trim one group's
    members -- an already-established invariant this reuses rather than
    introduces (see the scope test docstrings in ``tests/test_population_
    finalize.py``).

    When ``telemetry`` is given, ``finalize_seconds`` measures only this
    function's own per-group grouping/decision loop body; ``spill_seconds``
    is the residual (this function's total wall time minus that measured
    body), attributing ``spill.groups()``'s own generator work -- closing
    the spill writer and the ``merge_spill`` k-way merge -- without double
    counting time already spent inside the loop body.
    """
    total_start = time.perf_counter() if telemetry is not None else 0.0
    finalize_elapsed = 0.0
    population_findings: list[Finding] = []
    replay_findings: list[Finding] = []
    stats: dict[FindingClass, ClassPopulationStats] = defaultdict(ClassPopulationStats)
    for group in spill.groups():
        body_start = time.perf_counter() if telemetry is not None else 0.0
        template = next(
            (row["row"]["finding"] for row in group if row["row"]["kind"] == "full"), None
        )
        assert template is not None, "every group has at least one full row"
        representative = Finding.from_trusted_payload(template)
        if not scope.filter_findings([representative]):
            if telemetry is not None:
                finalize_elapsed += time.perf_counter() - body_start
            continue
        class_stats = stats[representative.finding_class]
        class_stats.candidates += len(group)

        facts = [_row_facts(row, template) for row in group]
        members = [fact for fact in facts if fact is not None]
        if len(members) != len(group) or len(members) < policy.threshold:
            class_stats.below_threshold += len(group)
            class_stats.replayed += len(group)
            replay_findings.extend(_reconstruct_group_findings(group))
            if telemetry is not None:
                finalize_elapsed += time.perf_counter() - body_start
            continue
        evidence = [_row_evidence(row, template) for row in group]
        first_event_key, first_tags = evidence[0]
        homogeneous = all(
            event_key == first_event_key and tags == first_tags
            for event_key, tags in evidence[1:]
        )
        if not homogeneous:
            class_stats.heterogeneous_evidence += len(group)
            class_stats.replayed += len(group)
            replay_findings.extend(_reconstruct_group_findings(group))
            if telemetry is not None:
                finalize_elapsed += time.perf_counter() - body_start
            continue
        membership = _build_membership(members, policy)
        if membership is None:
            class_stats.over_cap += len(group)
            class_stats.replayed += len(group)
            replay_findings.extend(_reconstruct_group_findings(group))
            if telemetry is not None:
                finalize_elapsed += time.perf_counter() - body_start
            continue
        shape_before = group[0]["shape_before"]
        shape_after = group[0]["shape_after"]
        population_findings.append(
            _build_population(representative, members, membership, shape_before, shape_after)
        )
        class_stats.populations += 1
        if telemetry is not None:
            finalize_elapsed += time.perf_counter() - body_start
    if telemetry is not None:
        telemetry.finalize_seconds += finalize_elapsed
        total_elapsed = time.perf_counter() - total_start
        telemetry.spill_seconds += max(0.0, total_elapsed - finalize_elapsed)
    return PopulationOutcome(population_findings, replay_findings, dict(stats))


def _build_membership(
    members: list[_MemberFacts], policy: PopulationPolicy
) -> MembershipCodec | None:
    """Build the dual-sided membership codec, or None if it exceeds caps."""
    current_coords = {member.current for member in members}
    rectangles = _rectangles(current_coords)
    if len(rectangles) > policy.max_rectangles:
        return None
    current_rectangles = tuple(_range(rectangle) for rectangle in rectangles)
    offsets = {
        (member.baseline[0] - member.current[0], member.baseline[1] - member.current[1])
        for member in members
        if member.baseline is not None
    }
    if len(offsets) == 1:
        return MembershipCodec(
            current_rectangles=current_rectangles,
            baseline_mode="shift",
            shift=next(iter(offsets)),
            member_count=len(members),
        )
    pairs: list[tuple[str, str]] = []
    for member in members:
        if member.baseline is not None and member.baseline_location is not None:
            pairs.append((member.location, member.baseline_location))
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
    members: list[_MemberFacts],
    membership: MembershipCodec,
    shape_before: str,
    shape_after: str,
) -> Finding:
    ordered = sorted(members, key=lambda member: member.current)
    current_coords = {member.current for member in members}
    baseline_coords = {
        member.baseline for member in members if member.baseline is not None
    }
    samples = tuple(
        PopulationSample(
            current_location=member.location,
            baseline_location=member.baseline_location,
        )
        for member in ordered[:_MAX_SAMPLES]
    )
    population = PopulationEvidence(
        member_count=len(members),
        membership=membership,
        first=ordered[0].location,
        last=ordered[-1].location,
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
