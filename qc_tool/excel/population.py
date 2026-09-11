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
import itertools
import shutil
import tempfile
import time
from collections import defaultdict
from collections.abc import Iterator, Sequence
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
    BLOCK_FINDINGS,
    BlockFile,
    FindingSequence,
    FindingsStoreError,
    SpillWriter,
    decode_block,
    finding_payload,
    merge_spill,
)
from qc_tool.review import (
    Coordinate,
    Rectangle,
    _bounding_range,
    _coordinate,
    _horizontal_runs,
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

# Delta templates are an optimization, never required for correctness. Once
# this cap is reached, unseen keys store full rows while already-admitted keys
# keep using their template. This bounds whole-run retained payloads under
# high-cardinality workloads without changing any reconstructed finding.
_CANDIDATE_TEMPLATE_CAP = 4_096

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
        self._disposed = False
        self.count = 0
        self._telemetry = telemetry
        self._templates: dict[tuple[object, ...], dict[str, Any]] = {}

    def add(self, candidate: Finding, *, shape_before: str, shape_after: str) -> None:
        start = time.perf_counter() if self._telemetry is not None else 0.0
        assign_severities([candidate], self._profile, today=self._today)
        key = population_key(candidate, (shape_before, shape_after))
        coordinate = _coordinate(candidate.location) or (0, 0)
        full_payload = finding_payload(candidate)
        template = self._templates.get(key)
        if template is None:
            if len(self._templates) < _CANDIDATE_TEMPLATE_CAP:
                self._templates[key] = full_payload
                self._spill.append(
                    (*key, -1, -1, -1),
                    {
                        "v": _CANDIDATE_ROW_VERSION,
                        "group_key": list(key),
                        "kind": "template",
                        "template": full_payload,
                    },
                )
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
        self._spill.append((*key, 0, coordinate[0], coordinate[1]), payload)
        self.count += 1
        if self._telemetry is not None:
            self._telemetry.construction_seconds += time.perf_counter() - start

    def iter_groups(
        self,
    ) -> Iterator[tuple[dict[str, Any], Iterator[dict[str, Any]]]]:
        """Yield each template and its coordinate-sorted candidate-row stream.

        The optional template marker sorts before every real candidate in its
        group. Keys rejected by the bounded template cache store only full
        rows, so their first candidate supplies the reconstruction template.
        The caller must exhaust each row iterator before requesting the next
        group, matching ``itertools.groupby``'s streaming contract.
        """
        if not self._closed:
            self._spill.__exit__(None, None, None)
            self._closed = True

        def group_key(raw: object) -> tuple[object, ...]:
            if not isinstance(raw, dict):
                raise FindingsStoreError("population candidate payload is not a mapping")
            key = raw.get("group_key")
            if not isinstance(key, list):
                raise FindingsStoreError(
                    "population candidate payload is missing its group key"
                )
            return tuple(key)

        for _key, raw_group in itertools.groupby(
            merge_spill(self._spill.path), key=group_key
        ):
            group = iter(raw_group)
            first = next(group)
            if not isinstance(first, dict):
                raise FindingsStoreError("population candidate payload is not a mapping")
            if first.get("kind") == "template":
                template = first.get("template")
                if not isinstance(template, dict):
                    raise FindingsStoreError(
                        "population candidate template is not a mapping"
                    )
                rows = group
            else:
                inner = first.get("row")
                template = (
                    inner.get("finding")
                    if isinstance(inner, dict) and inner.get("kind") == "full"
                    else None
                )
                if not isinstance(template, dict):
                    raise FindingsStoreError(
                        "population group has no reconstruction template"
                    )
                rows = itertools.chain((first,), group)

            def candidate_rows(
                source: Iterator[object] = rows,
            ) -> Iterator[dict[str, Any]]:
                for raw in source:
                    if not isinstance(raw, dict) or not isinstance(raw.get("row"), dict):
                        raise FindingsStoreError(
                            "population candidate row is not a mapping"
                        )
                    yield raw

            yield template, candidate_rows()

    def groups(self) -> Iterator[list[dict[str, Any]]]:
        """Compatibility wrapper that materializes one group at a time.

        Consumes (and removes) the spill directory; call at most once.
        Boundaries come directly from each row's own stored ``group_key``
        (computed once at ``add()`` time) rather than recomputing
        ``population_key`` from a reconstructed ``Finding`` -- a delta row's
        own payload deliberately omits most of what ``population_key`` reads,
        so recomputing it here would require reconstructing every row before
        grouping could even begin.
        """
        try:
            for _template, rows in self.iter_groups():
                yield list(rows)
        finally:
            self._dispose()

    def _dispose(self) -> None:
        if self._disposed:
            return
        shutil.rmtree(self._dir, ignore_errors=True)
        self._templates.clear()
        self._disposed = True

    def abort(self) -> None:
        if not self._closed:
            with contextlib.suppress(Exception):
                self._spill.__exit__(None, None, None)
            self._closed = True
        self._dispose()

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
    population_findings: Sequence[Finding] = field(default_factory=list)
    replay_findings: Sequence[Finding] = field(default_factory=list)
    stats: dict[FindingClass, ClassPopulationStats] = field(default_factory=dict)
    _storage: object | None = field(default=None, repr=False, compare=False)


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


@dataclass(slots=True)
class _GroupAccumulator:
    """Bounded facts needed to decide and encode one sorted candidate group."""

    policy: PopulationPolicy
    count: int = 0
    malformed: bool = False
    homogeneous: bool = True
    shape_before: str = ""
    shape_after: str = ""
    first_location: str = ""
    last_location: str = ""
    samples: list[PopulationSample] = field(default_factory=list)
    _first_evidence: tuple[str, frozenset[str]] | None = None
    _current_row: int | None = None
    _current_columns: list[int] = field(default_factory=list)
    _previous_row: int | None = None
    _active: dict[tuple[int, int], int] = field(default_factory=dict)
    _rectangles: list[Rectangle] = field(default_factory=list)
    _rectangle_overflow: bool = False
    _offsets: set[tuple[int, int]] = field(default_factory=set)
    _pairs: list[tuple[str, str]] = field(default_factory=list)
    _pair_count: int = 0
    _current_bounds: Rectangle | None = None
    _baseline_bounds: Rectangle | None = None
    _finished: bool = False

    def observe(self, row: dict[str, Any], template: dict[str, Any]) -> None:
        self.count += 1
        if self.count == 1:
            self.shape_before = str(row["shape_before"])
            self.shape_after = str(row["shape_after"])
        evidence = _row_evidence(row, template)
        if self._first_evidence is None:
            self._first_evidence = evidence
        elif evidence != self._first_evidence:
            self.homogeneous = False
        facts = _row_facts(row, template)
        if facts is None:
            self.malformed = True
            return
        if not self.first_location:
            self.first_location = facts.location
        self.last_location = facts.location
        if len(self.samples) < _MAX_SAMPLES:
            self.samples.append(
                PopulationSample(
                    current_location=facts.location,
                    baseline_location=facts.baseline_location,
                )
            )
        self._current_bounds = _extend_bounds(self._current_bounds, facts.current)
        if facts.baseline is not None:
            self._baseline_bounds = _extend_bounds(self._baseline_bounds, facts.baseline)
            offset = (
                facts.baseline[0] - facts.current[0],
                facts.baseline[1] - facts.current[1],
            )
            if len(self._offsets) < 2 or offset in self._offsets:
                self._offsets.add(offset)
        if facts.baseline_location is not None and facts.baseline is not None:
            self._pair_count += 1
            if len(self._pairs) <= self.policy.max_explicit_pairs:
                self._pairs.append((facts.location, facts.baseline_location))
        if self._current_row is None:
            self._current_row = facts.current[0]
        elif facts.current[0] != self._current_row:
            self._flush_row()
            self._current_row = facts.current[0]
        self._current_columns.append(facts.current[1])

    def _append_rectangle(self, rectangle: Rectangle) -> None:
        if len(self._rectangles) < self.policy.max_rectangles:
            self._rectangles.append(rectangle)
        else:
            self._rectangle_overflow = True

    def _flush_active(self, end_row: int) -> None:
        for (start_col, end_col), start_row in self._active.items():
            self._append_rectangle((start_row, start_col, end_row, end_col))
        self._active.clear()

    def _flush_row(self) -> None:
        if self._current_row is None:
            return
        row = self._current_row
        if self._previous_row is not None and row != self._previous_row + 1:
            self._flush_active(self._previous_row)
        runs = set(_horizontal_runs(sorted(set(self._current_columns))))
        for run in set(self._active).difference(runs):
            start_col, end_col = run
            self._append_rectangle((self._active.pop(run), start_col, row - 1, end_col))
        for run in runs.difference(self._active):
            self._active[run] = row
        self._previous_row = row
        self._current_columns.clear()

    def membership(self) -> MembershipCodec | None:
        if not self._finished:
            self._flush_row()
            if self._previous_row is not None:
                self._flush_active(self._previous_row)
            self._finished = True
        if self._rectangle_overflow:
            return None
        current_rectangles = tuple(_range(item) for item in sorted(self._rectangles))
        if len(self._offsets) == 1:
            return MembershipCodec(
                current_rectangles=current_rectangles,
                baseline_mode="shift",
                shift=next(iter(self._offsets)),
                member_count=self.count,
            )
        if self._pair_count != self.count or len(self._pairs) > self.policy.max_explicit_pairs:
            return None
        return MembershipCodec(
            current_rectangles=current_rectangles,
            baseline_mode="pairs",
            pairs=tuple(self._pairs),
            member_count=self.count,
        )

    @property
    def current_bounding_range(self) -> str:
        return _range(self._current_bounds) if self._current_bounds is not None else ""

    @property
    def baseline_bounding_range(self) -> str:
        return _range(self._baseline_bounds) if self._baseline_bounds is not None else ""

    @property
    def evidence(self) -> tuple[str, frozenset[str]]:
        return self._first_evidence or ("", frozenset())


def _extend_bounds(bounds: Rectangle | None, coordinate: Coordinate) -> Rectangle:
    row, column = coordinate
    if bounds is None:
        return row, column, row, column
    min_row, min_col, max_row, max_col = bounds
    return (
        min(min_row, row),
        min(min_col, column),
        max(max_row, row),
        max(max_col, column),
    )


def _finding_from_candidate_row(
    row: dict[str, Any], template: dict[str, Any]
) -> Finding:
    inner = row["row"]
    payload = (
        inner["finding"]
        if inner["kind"] == "full"
        else {**template, **inner["finding"]}
    )
    return Finding.from_trusted_payload(payload)


def _iter_block_payloads(path: Path) -> Iterator[object]:
    source = BlockFile.open(path)
    for index in range(len(source.block_infos())):
        yield from decode_block(source.read_block(index))


def _append_block_payload(
    container: BlockFile,
    buffer: list[object],
    payload: object,
) -> None:
    buffer.append(payload)
    if len(buffer) >= BLOCK_FINDINGS:
        container.append_rows(buffer)
        buffer.clear()


def _flush_block_payloads(container: BlockFile, buffer: list[object]) -> None:
    if buffer:
        container.append_rows(buffer)
        buffer.clear()


def _build_streamed_population(
    representative: Finding,
    group: _GroupAccumulator,
    membership: MembershipCodec,
) -> Finding:
    population = PopulationEvidence(
        member_count=group.count,
        membership=membership,
        first=group.first_location,
        last=group.last_location,
        samples=tuple(group.samples),
        shape_before_digest=group.shape_before,
        shape_after_digest=group.shape_after,
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
        location=group.current_bounding_range,
        baseline_location=group.baseline_bounding_range,
        element="population",
        message=(
            f"{representative.sheet}: {group.count} cells share one "
            f"{representative.finding_class.value} population"
        ),
        waiver_reason=representative.waiver_reason,
        waiver_expires=representative.waiver_expires,
        population=population,
    )
    finding.root_cause_key = "population:" + population_identity_digest(finding)
    return finding


def finalize_populations(
    spill: CandidateSpill,
    policy: PopulationPolicy,
    scope: ComparisonScope,
    *,
    telemetry: PopulationTelemetry | None = None,
) -> PopulationOutcome:
    """Stream candidate groups into bounded membership or exact replay output.

    One first pass aggregates only capped rectangles/pairs, five samples,
    bounds, counts, and evidence homogeneity. Compact rows are copied in sorted
    order to a block file. A second sequential pass reconstructs full findings
    only for groups marked for atomic replay. Both result collections are lazy
    block-backed sequences owned by the returned outcome.
    """
    total_start = time.perf_counter() if telemetry is not None else 0.0
    finalize_elapsed = 0.0
    stats: dict[FindingClass, ClassPopulationStats] = defaultdict(ClassPopulationStats)
    storage = tempfile.TemporaryDirectory(prefix="qc-population-outcome-")
    storage_path = Path(storage.name)
    ordered_path = storage_path / "ordered.qcfb"
    decisions_path = storage_path / "decisions.qcfb"
    populations_path = storage_path / "populations.qcfb"
    replay_path = storage_path / "replay.qcfb"
    try:
        ordered_buffer: list[object] = []
        decision_buffer: list[object] = []
        population_buffer: list[object] = []
        with (
            BlockFile(ordered_path) as ordered_file,
            BlockFile(decisions_path) as decision_file,
            BlockFile(populations_path) as population_file,
        ):
            for template, rows in spill.iter_groups():
                body_start = time.perf_counter() if telemetry is not None else 0.0
                representative = Finding.from_trusted_payload(template)
                group = _GroupAccumulator(policy)
                _append_block_payload(
                    ordered_file,
                    ordered_buffer,
                    {"kind": "template", "template": template},
                )
                for row in rows:
                    _append_block_payload(ordered_file, ordered_buffer, row)
                    group.observe(row, template)

                action = "drop"
                if scope.filter_findings([representative]):
                    class_stats = stats[representative.finding_class]
                    class_stats.candidates += group.count
                    if group.malformed or group.count < policy.threshold:
                        class_stats.below_threshold += group.count
                        class_stats.replayed += group.count
                        action = "replay"
                    elif not group.homogeneous:
                        class_stats.heterogeneous_evidence += group.count
                        class_stats.replayed += group.count
                        action = "replay"
                    else:
                        membership = group.membership()
                        if membership is None:
                            class_stats.over_cap += group.count
                            class_stats.replayed += group.count
                            action = "replay"
                        else:
                            population = _build_streamed_population(
                                representative,
                                group,
                                membership,
                            )
                            _append_block_payload(
                                population_file,
                                population_buffer,
                                finding_payload(population),
                            )
                            class_stats.populations += 1
                            action = "population"
                _append_block_payload(
                    decision_file,
                    decision_buffer,
                    {"count": group.count, "action": action},
                )
                if telemetry is not None:
                    finalize_elapsed += time.perf_counter() - body_start
            _flush_block_payloads(ordered_file, ordered_buffer)
            _flush_block_payloads(decision_file, decision_buffer)
            _flush_block_payloads(population_file, population_buffer)

        spill.abort()
        ordered = iter(_iter_block_payloads(ordered_path))
        replay_buffer: list[object] = []
        with BlockFile(replay_path) as replay_file:
            for raw_decision in _iter_block_payloads(decisions_path):
                if not isinstance(raw_decision, dict):
                    raise FindingsStoreError("population decision is not a mapping")
                raw_marker = next(ordered)
                if not isinstance(raw_marker, dict) or raw_marker.get("kind") != "template":
                    raise FindingsStoreError("population replay template is missing")
                template = raw_marker.get("template")
                if not isinstance(template, dict):
                    raise FindingsStoreError("population replay template is not a mapping")
                count = raw_decision.get("count")
                action = raw_decision.get("action")
                if not isinstance(count, int) or action not in {
                    "drop",
                    "population",
                    "replay",
                }:
                    raise FindingsStoreError("population decision is invalid")
                for _ in range(count):
                    raw_row = next(ordered)
                    if not isinstance(raw_row, dict):
                        raise FindingsStoreError("population replay row is not a mapping")
                    if action == "replay":
                        _append_block_payload(
                            replay_file,
                            replay_buffer,
                            finding_payload(_finding_from_candidate_row(raw_row, template)),
                        )
            try:
                next(ordered)
            except StopIteration:
                pass
            else:
                raise FindingsStoreError("population replay contains unconsumed rows")
            _flush_block_payloads(replay_file, replay_buffer)

        population_findings = FindingSequence(BlockFile.open(populations_path))
        replay_findings = FindingSequence(BlockFile.open(replay_path))
        if telemetry is not None:
            telemetry.finalize_seconds += finalize_elapsed
            total_elapsed = time.perf_counter() - total_start
            telemetry.spill_seconds += max(0.0, total_elapsed - finalize_elapsed)
        return PopulationOutcome(
            population_findings=population_findings,
            replay_findings=replay_findings,
            stats=dict(stats),
            _storage=storage,
        )
    except BaseException:
        spill.abort()
        storage.cleanup()
        raise


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
