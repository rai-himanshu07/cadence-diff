"""Streaming builders for review groups, pattern groups, and stories.

These produce byte-identical group IDs, ordering, counts, and stories to the
list-based builders in ``qc_tool.review`` and ``qc_tool.story`` while holding
compact per-member records instead of ``Finding`` objects, so arbitrarily
large runs can be summarized in bounded memory. Equality with the list-based
builders is pinned by tests on every fixture.
"""

from __future__ import annotations

import json
import zlib
from collections import Counter, defaultdict, deque
from collections.abc import Iterable, Iterator
from dataclasses import dataclass, replace
from typing import NamedTuple

from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingEvidenceTag,
    FindingProvenance,
    FindingSubtype,
    FindingTemporalContext,
    Materiality,
    Severity,
)
from qc_tool.review import (
    _DEFERRAL_TIERS,
    _PRIORITY_WEIGHTS,
    _SEVERITY_RANK,
    _WIDE_POPULATION,
    Coordinate,
    ReviewCounts,
    ReviewGroup,
    _bounding_range,
    _coordinate,
    _eligible,
    _finding_order,
    _group_id_from_identities,
    _partition_key,
    _pattern_key,
    _range,
    _rationale,
    _rectangles,
    finding_identity_key,
    represented_change_count,
)
from qc_tool.story import (
    _DRIVER_CLASSES,
    _MAX_EVIDENCE_LINES,
    ChangeStory,
    StoryKind,
    _added_reference_tokens,
    _clip,
    _driver_tokens,
    _missing_evidence,
    _severity_names,
    _stable_id,
    _story_sort_key,
    _UnionFind,
)

__all__ = [
    "VIEW_SUMMARY_VERSION",
    "GroupPriorityAggregate",
    "GroupSummary",
    "RecordSummaryAccumulators",
    "StoryStreamer",
    "SummaryReviewPriority",
    "counts_from_summaries",
    "decode_view_summaries",
    "encode_view_summaries",
    "prioritize_review_summaries",
    "stream_stories",
    "summarize_pattern_groups",
    "summarize_pattern_groups_with_priority",
    "summarize_review_groups",
    "summary_of",
]


@dataclass(frozen=True, slots=True)
class GroupSummary:
    """Everything a review surface needs from a group except live members."""

    group_id: str
    finding_class: FindingClass
    severity: Severity
    artifact: str
    sheet: str | None
    slide: str | None
    element: str
    expected_growth: bool
    ranges: tuple[str, ...]
    bounding_range: str
    baseline_ranges: tuple[str, ...]
    baseline_bounding_range: str
    baseline_mixed: bool
    spatial: bool
    artifact_member: str
    member_finding_ids: tuple[str, ...]
    represented_change_count: int | None = None

    @property
    def member_count(self) -> int:
        return len(self.member_finding_ids)


def summary_of(group: ReviewGroup) -> GroupSummary:
    """Project a list-built group onto the summary shape (test oracle)."""
    return GroupSummary(
        group_id=group.group_id,
        finding_class=group.finding_class,
        severity=group.severity,
        artifact=group.artifact,
        sheet=group.sheet,
        slide=group.slide,
        element=group.element,
        expected_growth=group.expected_growth,
        ranges=group.ranges,
        bounding_range=group.bounding_range,
        baseline_ranges=group.baseline_ranges,
        baseline_bounding_range=group.baseline_bounding_range,
        baseline_mixed=group.baseline_mixed,
        spatial=group.spatial,
        artifact_member=group.artifact_member,
        member_finding_ids=tuple(member.finding_id for member in group.members),
        represented_change_count=sum(
            represented_change_count(member) for member in group.members
        ),
    )


def counts_from_summaries(summaries: Iterable[GroupSummary]) -> ReviewCounts:
    review_items = dict.fromkeys(Severity, 0)
    atomic_findings = dict.fromkeys(Severity, 0)
    represented_changes = dict.fromkeys(Severity, 0)
    represented_complete = True
    for summary in summaries:
        review_items[summary.severity] += 1
        atomic_findings[summary.severity] += summary.member_count
        if summary.represented_change_count is None:
            represented_complete = False
        else:
            represented_changes[summary.severity] += summary.represented_change_count
    return ReviewCounts(
        review_items=review_items,
        atomic_findings=atomic_findings,
        represented_changes=represented_changes if represented_complete else None,
    )


VIEW_SUMMARY_VERSION = 2


@dataclass(frozen=True, slots=True)
class SummaryReviewPriority:
    """``ReviewPriority`` twin backed by a summary instead of live members."""

    summary: GroupSummary
    rank: int
    score: int
    signals: tuple[str, ...]
    rationale: str


def _summary_sort_key(summary: GroupSummary) -> tuple[object, ...]:
    """``_review_sort_key`` twin without the member tail.

    The materialized key ends in per-member payload tuples AFTER the unique
    ``group_id``, so distinct groups never compare past it — dropping the
    tail cannot change the order.
    """
    coordinate = _coordinate(summary.bounding_range.partition(":")[0])
    return (
        _SEVERITY_RANK[summary.severity],
        summary.artifact,
        summary.sheet or summary.slide or "",
        coordinate or (10**9, 10**9),
        summary.finding_class.value,
        summary.element,
        summary.group_id,
    )


def prioritize_review_summaries(
    summaries: list[GroupSummary],
    aggregates: dict[str, GroupPriorityAggregate],
    stories: list[ChangeStory] | None = None,
    *,
    reviewed_ids: frozenset[str] | set[str] = frozenset(),
) -> list[SummaryReviewPriority]:
    """``prioritize_review`` twin over stored summaries.

    ``reviewed_ids`` carries live analyst annotations (the materialized path
    reads them off decorated members); waivers need no live source because
    they only enter findings through profile promotion plus re-QC.
    """
    story_by_finding: dict[str, str] = {}
    for story in stories or ():
        if story.kind is StoryKind.RESIDUAL:
            continue
        for finding_id in story.finding_ids:
            story_by_finding.setdefault(finding_id, story.story_id)

    scored: list[
        tuple[int, int, tuple[object, ...], GroupSummary, list[str], dict[str, object]]
    ] = []
    for summary in summaries:
        aggregate = aggregates.get(summary.group_id, GroupPriorityAggregate())
        counts: dict[str, object] = {
            "members": summary.member_count,
            "material": aggregate.material,
            "historical": aggregate.historical,
            "new": aggregate.new,
            "impacts": aggregate.impacts,
        }
        signals: list[str] = []
        if summary.severity is Severity.CRITICAL:
            signals.append("severity_critical")
        elif summary.severity is Severity.WARNING:
            signals.append("severity_warning")
        elif summary.severity is Severity.INFO:
            signals.append("severity_info")
        if aggregate.material:
            signals.append("material_delta")
        if aggregate.historical:
            signals.append("historical_change")
        if aggregate.new:
            signals.append("new_defect")
        if aggregate.impacts:
            signals.append("downstream_impacts")
        if summary.member_count >= _WIDE_POPULATION:
            signals.append("wide_population")
        if aggregate.tiers == (Materiality.NOISE.value,):
            signals.append("noise_only")
        elif aggregate.tiers == (Materiality.WITHIN_TOLERANCE.value,):
            signals.append("within_tolerance_only")
        if summary.member_count and aggregate.all_inherited:
            signals.append("inherited_only")
        story_id = next(
            (
                story_by_finding[finding_id]
                for finding_id in summary.member_finding_ids
                if finding_id in story_by_finding
            ),
            None,
        )
        if story_id is None:
            signals.append("unexplained_residual")
        else:
            signals.append("story_explained")
            counts["story"] = story_id
        if aggregate.any_waiver:
            signals.append("waived")
        if aggregate.any_reviewed or any(
            finding_id in reviewed_ids
            for finding_id in summary.member_finding_ids
        ):
            signals.append("already_reviewed")
        if summary.expected_growth or summary.severity is Severity.EXPECTED:
            signals.append("expected_growth")
        score = sum(_PRIORITY_WEIGHTS[signal] for signal in signals)
        tier = max((_DEFERRAL_TIERS.get(signal, 0) for signal in signals), default=0)
        scored.append(
            (tier, score, _summary_sort_key(summary), summary, signals, counts)
        )

    scored.sort(key=lambda entry: (entry[0], -entry[1], entry[2]))
    return [
        SummaryReviewPriority(
            summary=summary,
            rank=index,
            score=score,
            signals=tuple(signals),
            rationale=_rationale(signals, counts),
        )
        for index, (_tier, score, _order, summary, signals, counts) in enumerate(
            scored, start=1
        )
    ]


def encode_view_summaries(
    summaries: list[GroupSummary],
    aggregates: dict[str, GroupPriorityAggregate],
    stories: list[ChangeStory],
) -> bytes:
    """Compressed view-layer payload stored beside a recorded run."""
    payload = {
        "version": VIEW_SUMMARY_VERSION,
        "groups": [
            {
                "group_id": summary.group_id,
                "finding_class": summary.finding_class.value,
                "severity": summary.severity.value,
                "artifact": summary.artifact,
                "sheet": summary.sheet,
                "slide": summary.slide,
                "element": summary.element,
                "expected_growth": summary.expected_growth,
                "ranges": list(summary.ranges),
                "bounding_range": summary.bounding_range,
                "baseline_ranges": list(summary.baseline_ranges),
                "baseline_bounding_range": summary.baseline_bounding_range,
                "baseline_mixed": summary.baseline_mixed,
                "spatial": summary.spatial,
                "artifact_member": summary.artifact_member,
                "member_finding_ids": list(summary.member_finding_ids),
                "represented_change_count": summary.represented_change_count,
                "priority": {
                    "material": aggregates[summary.group_id].material,
                    "historical": aggregates[summary.group_id].historical,
                    "new": aggregates[summary.group_id].new,
                    "impacts": aggregates[summary.group_id].impacts,
                    "tiers": list(aggregates[summary.group_id].tiers),
                    "all_inherited": aggregates[summary.group_id].all_inherited,
                    "any_waiver": aggregates[summary.group_id].any_waiver,
                    "any_reviewed": aggregates[summary.group_id].any_reviewed,
                },
            }
            for summary in summaries
        ],
        "stories": [
            {
                "story_id": story.story_id,
                "kind": story.kind.value,
                "title": story.title,
                "description": story.description,
                "evidence": list(story.evidence),
                "finding_ids": list(story.finding_ids),
                "member_count": story.member_count,
                "severity_counts": story.severity_counts,
            }
            for story in stories
        ],
    }
    return zlib.compress(
        json.dumps(payload, ensure_ascii=False).encode("utf-8"), 6
    )


def decode_view_summaries(
    blob: bytes,
) -> tuple[
    list[GroupSummary], dict[str, GroupPriorityAggregate], list[ChangeStory]
] | None:
    """Decode a stored view payload; anything unreadable yields ``None``.

    ``None`` means the caller falls back to the bounded view — never a crash
    and never a silently wrong queue.
    """
    try:
        payload = json.loads(zlib.decompress(blob).decode("utf-8"))
        version = payload.get("version")
        if version not in {1, VIEW_SUMMARY_VERSION}:
            return None
        summaries: list[GroupSummary] = []
        aggregates: dict[str, GroupPriorityAggregate] = {}
        for item in payload["groups"]:
            summaries.append(
                GroupSummary(
                    group_id=item["group_id"],
                    finding_class=FindingClass(item["finding_class"]),
                    severity=Severity(item["severity"]),
                    artifact=item["artifact"],
                    sheet=item["sheet"],
                    slide=item["slide"],
                    element=item["element"],
                    expected_growth=item["expected_growth"],
                    ranges=tuple(item["ranges"]),
                    bounding_range=item["bounding_range"],
                    baseline_ranges=tuple(item["baseline_ranges"]),
                    baseline_bounding_range=item["baseline_bounding_range"],
                    baseline_mixed=item["baseline_mixed"],
                    spatial=item["spatial"],
                    artifact_member=item["artifact_member"],
                    member_finding_ids=tuple(item["member_finding_ids"]),
                    represented_change_count=(
                        item.get("represented_change_count")
                        if version >= 2
                        else None
                    ),
                )
            )
            priority = item["priority"]
            aggregates[item["group_id"]] = GroupPriorityAggregate(
                material=priority["material"],
                historical=priority["historical"],
                new=priority["new"],
                impacts=priority["impacts"],
                tiers=tuple(priority["tiers"]),
                all_inherited=priority["all_inherited"],
                any_waiver=priority["any_waiver"],
                any_reviewed=priority["any_reviewed"],
            )
        stories = [
            ChangeStory(
                story_id=item["story_id"],
                kind=StoryKind(item["kind"]),
                title=item["title"],
                description=item["description"],
                evidence=tuple(item["evidence"]),
                finding_ids=tuple(item["finding_ids"]),
                member_count=item["member_count"],
                severity_counts=dict(item["severity_counts"]),
            )
            for item in payload["stories"]
        ]
        return summaries, aggregates, stories
    except Exception:
        return None


class _Member(NamedTuple):
    order: tuple[object, ...]
    finding_id: str
    identity: tuple[str, ...]
    coordinate: Coordinate | None
    baseline: Coordinate | None
    element_raw: str
    location: str
    represented_change_count: int
    sort_payload: tuple[object, ...]


def _member_record(finding: Finding, *, with_payload: bool) -> _Member:
    return _Member(
        order=_finding_order(finding),
        finding_id=finding.finding_id,
        identity=finding_identity_key(finding),
        coordinate=_coordinate(finding.location),
        baseline=_coordinate(finding.baseline_location),
        element_raw=finding.element or "",
        location=finding.location or finding.baseline_location or "",
        represented_change_count=represented_change_count(finding),
        # Mirrors _review_sort_key's member tuple. Only groups that can share
        # a group_id (identical key + identities, i.e. duplicate singletons)
        # ever reach this tiebreak, so it is retained only when requested.
        sort_payload=(
            (
                finding.message,
                finding.baseline_value or "",
                finding.current_value or "",
                tuple(finding.impacts),
                finding.finding_id,
            )
            if with_payload
            else ()
        ),
    )


@dataclass(frozen=True, slots=True)
class _SortableSummary:
    summary: GroupSummary
    member_payloads: tuple[tuple[object, ...], ...]

    def sort_key(self) -> tuple[object, ...]:
        summary = self.summary
        coordinate = _coordinate(summary.bounding_range.partition(":")[0])
        return (
            _SEVERITY_RANK[summary.severity],
            summary.artifact,
            summary.sheet or summary.slide or "",
            coordinate or (10**9, 10**9),
            summary.finding_class.value,
            summary.element,
            summary.group_id,
            self.member_payloads,
        )


def _finish(entries: list[_SortableSummary]) -> list[GroupSummary]:
    entries.sort(key=_SortableSummary.sort_key)
    seen: dict[str, int] = defaultdict(int)
    finished: list[GroupSummary] = []
    for entry in entries:
        seen[entry.summary.group_id] += 1
        occurrence = seen[entry.summary.group_id]
        finished.append(
            entry.summary
            if occurrence == 1
            else replace(
                entry.summary,
                group_id=f"{entry.summary.group_id}-{occurrence}",
            )
        )
    return finished


def _singleton_summary(finding: Finding) -> _SortableSummary:
    severity = finding.severity or Severity.WARNING
    current = _coordinate(finding.location)
    baseline = _coordinate(finding.baseline_location)
    key = (
        "singleton",
        finding_identity_key(finding),
        severity.value,
        finding.message,
        finding.baseline_value or "",
        finding.current_value or "",
        tuple(finding.impacts),
    )
    record = _member_record(finding, with_payload=True)
    return _SortableSummary(
        summary=GroupSummary(
            group_id=_group_id_from_identities(key, (record.identity,)),
            finding_class=finding.finding_class,
            severity=severity,
            artifact=finding.artifact,
            sheet=finding.sheet,
            slide=finding.slide,
            element=finding.element or "",
            expected_growth=finding.expected_growth,
            ranges=(_range((*current, *current)),) if current is not None else (),
            bounding_range=(
                _range((*current, *current))
                if current is not None
                else finding.location or finding.baseline_location or ""
            ),
            baseline_ranges=(
                (_range((*baseline, *baseline)),) if baseline is not None else ()
            ),
            baseline_bounding_range=(
                _range((*baseline, *baseline)) if baseline is not None else ""
            ),
            baseline_mixed=False,
            spatial=False,
            artifact_member=finding.artifact_member,
            member_finding_ids=(finding.finding_id,),
            represented_change_count=record.represented_change_count,
        ),
        member_payloads=(record.sort_payload,),
    )


class _SpatialPartition:
    __slots__ = (
        "artifact",
        "artifact_member",
        "by_coordinate",
        "expected_growth",
        "finding_class",
        "severity",
        "sheet",
    )

    def __init__(self, finding: Finding) -> None:
        self.by_coordinate: dict[Coordinate, list[_Member]] = defaultdict(list)
        self.finding_class = finding.finding_class
        self.severity = finding.severity or Severity.WARNING
        self.artifact = finding.artifact
        self.sheet = finding.sheet
        self.expected_growth = finding.expected_growth
        self.artifact_member = finding.artifact_member


def _spatial_summaries(
    key: tuple[object, ...],
    partition: _SpatialPartition,
) -> Iterator[_SortableSummary]:
    by_coordinate = partition.by_coordinate
    pending = set(by_coordinate)
    # Repeatedly calling min() on a shrinking set is O(n) per call, O(n^2)
    # overall when a partition has many small disconnected components (a
    # sheet with mixed finding classes at scattered coordinates). A sorted
    # scan that skips already-visited coordinates picks the exact same
    # "smallest remaining" seed every time (it's the first not-yet-visited
    # entry in ascending order) in O(n log n) total instead.
    for first in sorted(by_coordinate):
        if first not in pending:
            continue
        pending.remove(first)
        component = {first}
        queue = deque([first])
        while queue:
            row, column = queue.popleft()
            for neighbor in (
                (row - 1, column),
                (row + 1, column),
                (row, column - 1),
                (row, column + 1),
            ):
                if neighbor in pending:
                    pending.remove(neighbor)
                    component.add(neighbor)
                    queue.append(neighbor)
        members = sorted(
            (
                member
                for coordinate in component
                for member in by_coordinate[coordinate]
            ),
            key=lambda member: member.order,
        )
        representative = members[0]
        baseline_coordinates = {
            member.baseline for member in members if member.baseline is not None
        }
        baseline_ranges = (
            tuple(_range(item) for item in _rectangles(baseline_coordinates))
            if len(baseline_coordinates) == len(members)
            else ()
        )
        yield _SortableSummary(
            summary=GroupSummary(
                group_id=_group_id_from_identities(
                    key, tuple(member.identity for member in members)
                ),
                finding_class=partition.finding_class,
                severity=partition.severity,
                artifact=partition.artifact,
                sheet=partition.sheet,
                slide=None,
                element=representative.element_raw,
                expected_growth=partition.expected_growth,
                ranges=tuple(_range(item) for item in _rectangles(component)),
                bounding_range=_bounding_range(component),
                baseline_ranges=baseline_ranges,
                baseline_bounding_range=(
                    _bounding_range(baseline_coordinates) if baseline_ranges else ""
                ),
                baseline_mixed=0 < len(baseline_coordinates) < len(members),
                spatial=True,
                artifact_member=partition.artifact_member,
                member_finding_ids=tuple(member.finding_id for member in members),
                represented_change_count=sum(
                    member.represented_change_count for member in members
                ),
            ),
            member_payloads=(),
        )


class _ReviewGroupAccumulator:
    """Per-finding observer form of ``summarize_review_groups``."""

    __slots__ = ("entries", "partitions")

    def __init__(self) -> None:
        self.partitions: dict[tuple[object, ...], _SpatialPartition] = {}
        self.entries: list[_SortableSummary] = []

    def observe(self, finding: Finding) -> None:
        coordinate = _eligible(finding)
        if coordinate is None:
            self.entries.append(_singleton_summary(finding))
            return
        key = _partition_key(finding, coordinate)
        partition = self.partitions.get(key)
        if partition is None:
            partition = self.partitions[key] = _SpatialPartition(finding)
        partition.by_coordinate[coordinate].append(
            _member_record(finding, with_payload=False)
        )

    def finish(self) -> list[GroupSummary]:
        for key, partition in self.partitions.items():
            self.entries.extend(_spatial_summaries(key, partition))
        return _finish(self.entries)


def summarize_review_groups(findings: Iterable[Finding]) -> list[GroupSummary]:
    """Streaming twin of ``build_review_groups`` returning summaries."""
    accumulator = _ReviewGroupAccumulator()
    for finding in findings:
        accumulator.observe(finding)
    return accumulator.finish()


class _PatternPartition:
    __slots__ = (
        "artifact",
        "artifact_member",
        "expected_growth",
        "finding_class",
        "members",
        "severity",
        "sheet",
        "slide",
    )

    def __init__(self, finding: Finding) -> None:
        self.members: list[_Member] = []
        self.finding_class = finding.finding_class
        self.severity = finding.severity or Severity.WARNING
        self.artifact = finding.artifact
        self.sheet = finding.sheet
        self.slide = finding.slide
        self.expected_growth = finding.expected_growth
        self.artifact_member = finding.artifact_member


@dataclass(frozen=True, slots=True)
class GroupPriorityAggregate:
    """Record-time member rollup that prioritization needs per group.

    Everything ``_priority_signals`` reads from live members, accumulated
    while the summarize pass already has each finding in hand, so a stored
    run can be prioritized without hydrating a single member.
    """

    material: int = 0
    historical: int = 0
    new: int = 0
    impacts: int = 0
    tiers: tuple[str, ...] = ()
    all_inherited: bool = False
    any_waiver: bool = False
    any_reviewed: bool = False


class _PriorityAccumulator:
    __slots__ = (
        "all_inherited",
        "any_reviewed",
        "any_waiver",
        "historical",
        "impacts",
        "material",
        "new",
        "tiers",
    )

    def __init__(self) -> None:
        self.material = 0
        self.historical = 0
        self.new = 0
        self.impacts = 0
        self.tiers: set[str] = set()
        self.all_inherited = True
        self.any_waiver = False
        self.any_reviewed = False

    def observe(self, finding: Finding) -> None:
        if finding.materiality is Materiality.MATERIAL:
            self.material += 1
        if finding.temporal_context is FindingTemporalContext.HISTORICAL:
            self.historical += 1
        if finding.provenance is FindingProvenance.NEW:
            self.new += 1
        self.impacts += len(finding.impacts)
        if finding.materiality is not None:
            self.tiers.add(finding.materiality.value)
        if finding.provenance is not FindingProvenance.INHERITED:
            self.all_inherited = False
        if finding.waiver_reason:
            self.any_waiver = True
        if finding.severity_overridden or finding.analyst_comment:
            self.any_reviewed = True

    def frozen(self) -> GroupPriorityAggregate:
        return GroupPriorityAggregate(
            material=self.material,
            historical=self.historical,
            new=self.new,
            impacts=self.impacts,
            tiers=tuple(sorted(self.tiers)),
            all_inherited=self.all_inherited,
            any_waiver=self.any_waiver,
            any_reviewed=self.any_reviewed,
        )


class _PatternAccumulator:
    """Per-finding observer form of ``summarize_pattern_groups_with_priority``."""

    __slots__ = ("accumulators", "partitions")

    def __init__(self) -> None:
        self.partitions: dict[tuple[object, ...], _PatternPartition] = {}
        self.accumulators: dict[tuple[object, ...], _PriorityAccumulator] = {}

    def observe(self, finding: Finding) -> None:
        key = _pattern_key(finding)
        partition = self.partitions.get(key)
        if partition is None:
            partition = self.partitions[key] = _PatternPartition(finding)
            self.accumulators[key] = _PriorityAccumulator()
        partition.members.append(_member_record(finding, with_payload=False))
        self.accumulators[key].observe(finding)

    def finish(
        self,
    ) -> tuple[list[GroupSummary], dict[str, GroupPriorityAggregate]]:
        entries: list[_SortableSummary] = []
        aggregates_by_key: dict[tuple[object, ...], GroupPriorityAggregate] = {
            key: accumulator.frozen()
            for key, accumulator in self.accumulators.items()
        }
        # Keyed by first member id: unique per group and immune to the
        # duplicate ``-N`` group-id suffixes applied in ``_finish``.
        aggregates_by_first: dict[str, GroupPriorityAggregate] = {}
        for key, partition in self.partitions.items():
            members = sorted(partition.members, key=lambda member: member.order)
            representative = members[0]
            aggregates_by_first[representative.finding_id] = aggregates_by_key[key]
            coordinates = {
                member.coordinate
                for member in members
                if member.coordinate is not None
            }
            baseline_coordinates = {
                member.baseline for member in members if member.baseline is not None
            }
            baseline_ranges = (
                tuple(_range(item) for item in _rectangles(baseline_coordinates))
                if len(baseline_coordinates) == len(members)
                else ()
            )
            group_id = _group_id_from_identities(
                key, tuple(member.identity for member in members)
            )
            entries.append(
                _SortableSummary(
                    summary=GroupSummary(
                        group_id=group_id,
                        finding_class=partition.finding_class,
                        severity=partition.severity,
                        artifact=partition.artifact,
                        sheet=partition.sheet,
                        slide=partition.slide,
                        element=representative.element_raw,
                        expected_growth=partition.expected_growth,
                        ranges=tuple(
                            _range(item) for item in _rectangles(coordinates)
                        ),
                        bounding_range=(
                            _bounding_range(coordinates)
                            if coordinates
                            else representative.location
                        ),
                        baseline_ranges=baseline_ranges,
                        baseline_bounding_range=(
                            _bounding_range(baseline_coordinates)
                            if baseline_ranges
                            else ""
                        ),
                        baseline_mixed=0 < len(baseline_coordinates) < len(members),
                        spatial=False,
                        artifact_member=partition.artifact_member,
                        member_finding_ids=tuple(
                            member.finding_id for member in members
                        ),
                        represented_change_count=sum(
                            member.represented_change_count for member in members
                        ),
                    ),
                    member_payloads=(),
                )
            )
        finished = _finish(entries)
        return finished, {
            summary.group_id: aggregates_by_first[summary.member_finding_ids[0]]
            for summary in finished
        }


def summarize_pattern_groups(findings: Iterable[Finding]) -> list[GroupSummary]:
    """Streaming twin of ``build_pattern_groups`` returning summaries."""
    return summarize_pattern_groups_with_priority(findings)[0]


def summarize_pattern_groups_with_priority(
    findings: Iterable[Finding],
) -> tuple[list[GroupSummary], dict[str, GroupPriorityAggregate]]:
    """Summaries plus the per-group rollup, in the same single pass."""
    accumulator = _PatternAccumulator()
    for finding in findings:
        accumulator.observe(finding)
    return accumulator.finish()


class RecordSummaryAccumulators:
    """Record-time passes fused into one findings iteration.

    ``record_run`` observes each finding once while it already streams them
    for blocks, seeds, and occurrences — the separate review, pattern, and
    story-pass-1 iterations disappear.
    """

    __slots__ = ("pattern", "review", "stories")

    def __init__(self) -> None:
        self.review = _ReviewGroupAccumulator()
        self.pattern = _PatternAccumulator()
        self.stories = StoryStreamer()

    def observe(self, finding: Finding) -> None:
        self.review.observe(finding)
        self.pattern.observe(finding)
        self.stories.observe(finding)


# --- streaming stories -------------------------------------------------------


class _StoryState:
    """Accumulated materialization inputs for one (kind, component)."""

    __slots__ = (
        "component",
        "driver_anchor_tokens",
        "driver_elements",
        "driver_messages",
        "finding_ids",
        "first_event_key",
        "formula_event_keys",
        "formula_member_count",
        "kind",
        "materiality_tiers",
        "member_count",
        "member_messages",
        "missing_evidence",
        "population_tags",
        "provenances",
        "residual_classes",
        "severity_counts",
    )

    def __init__(self, kind: StoryKind, component: int) -> None:
        self.kind = kind
        self.component = component
        self.finding_ids: list[str] = []
        self.severity_counts: Counter[str] = Counter()
        self.driver_messages: list[str] = []
        self.driver_elements: set[str] = set()
        self.driver_anchor_tokens: set[str] = set()
        self.formula_member_count = 0
        self.formula_event_keys: set[str] = set()
        self.member_messages: list[str] = []
        self.population_tags: set[FindingEvidenceTag] = set()
        self.first_event_key: str | None = None
        self.materiality_tiers: Counter[str] = Counter()
        self.provenances: Counter[str] = Counter()
        self.residual_classes: Counter[str] = Counter()
        self.missing_evidence: set[str] = set()
        self.member_count = 0

    def add(self, finding: Finding) -> None:
        self.member_count += 1
        if finding.finding_id:
            self.finding_ids.append(finding.finding_id)
        self.severity_counts[(finding.severity or Severity.WARNING).value] += 1
        kind = self.kind
        if kind is StoryKind.STRUCTURE_DRIVER:
            if finding.finding_class in _DRIVER_CLASSES:
                if len(self.driver_messages) < _MAX_EVIDENCE_LINES:
                    self.driver_messages.append(_clip(finding.message))
                if finding.element:
                    self.driver_elements.add(finding.element)
                self.driver_anchor_tokens.update(_driver_tokens(finding))
            if finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED:
                self.formula_member_count += 1
                if finding.event_key:
                    self.formula_event_keys.add(finding.event_key)
        elif kind is StoryKind.ERROR_POPULATION:
            self.population_tags.update(
                tag
                for tag in finding.evidence_tags
                if tag
                in {
                    FindingEvidenceTag.CONCENTRATED_POPULATION,
                    FindingEvidenceTag.CONTIGUOUS_POPULATION,
                    FindingEvidenceTag.SPARSE_MASS_POPULATION,
                }
            )
            if len(self.member_messages) < _MAX_EVIDENCE_LINES:
                self.member_messages.append(_clip(finding.message))
            if self.first_event_key is None and finding.event_key:
                self.first_event_key = finding.event_key
        elif kind is StoryKind.DERIVED_LABELS:
            if len(self.member_messages) < _MAX_EVIDENCE_LINES:
                self.member_messages.append(_clip(finding.message))
        elif kind is StoryKind.DATA_REFRESH:
            if finding.materiality:
                self.materiality_tiers[finding.materiality.value] += 1
        elif kind is StoryKind.INHERITED:
            if finding.provenance:
                self.provenances[finding.provenance.value] += 1
        elif kind is StoryKind.RESIDUAL:
            self.residual_classes[finding.finding_class.value] += 1
            self.missing_evidence.update(_missing_evidence(finding))

    def materialize(self) -> ChangeStory:
        kind = self.kind
        evidence: list[str] = []
        if kind is StoryKind.STRUCTURE_DRIVER:
            evidence.extend(f"driver: {message}" for message in self.driver_messages)
            if self.formula_member_count:
                evidence.append(
                    f"{self.formula_member_count} formula change(s) reference the new "
                    f"structure or share {len(self.formula_event_keys)} wrapper skeleton(s)"
                )
            elements = sorted(self.driver_elements)
            subject = (
                ", ".join(elements[:3]) if elements else "systematic formula rollout"
            )
            title = f"Coordinated structure/formula change: {subject}"
            description = (
                "Structural edits and the formula changes that reference them, "
                "reviewed as one deliberate change."
                if self.driver_messages
                else "A systematic formula rewrite sharing one wrapper skeleton; "
                "no structural driver was identified."
            )
            anchor = ",".join(sorted(self.driver_anchor_tokens)) or (
                sorted(self.formula_event_keys) or [str(self.component)]
            )[0]
        elif kind is StoryKind.ERROR_POPULATION:
            title = f"Error population: {self.member_count} finding records"
            description = (
                "A proven same-artifact, sheet, column, and literal population. "
                "Grouping identifies one incident; it does not imply reduced risk."
            )
            if self.population_tags:
                evidence.append(
                    "evidence: "
                    + ", ".join(
                        tag.value for tag in sorted(self.population_tags, key=str)
                    )
                )
            evidence.extend(self.member_messages)
            anchor = self.first_event_key or str(self.component)
        elif kind is StoryKind.DERIVED_LABELS:
            title = "Derived labels re-resolved"
            description = (
                "Row/column keys are formula-derived; their upstream driver "
                "changed, not the historical rows themselves."
            )
            evidence.extend(self.member_messages)
            anchor = "derived-labels"
        elif kind is StoryKind.DATA_REFRESH:
            title = "New-cycle data arrival and in-window restatements"
            description = (
                "Expected cadence growth plus constant changes confined to the "
                "trailing restatement window or declared acceptance bands."
            )
            if self.materiality_tiers:
                evidence.append(
                    "tiers: "
                    + ", ".join(
                        f"{name}={count}"
                        for name, count in sorted(self.materiality_tiers.items())
                    )
                )
            anchor = "data-refresh"
        elif kind is StoryKind.ACCEPTED_DIFFERENCES:
            title = "Accepted differences within declared tolerance"
            description = (
                "Changes inside explicit analyst acceptance bounds. They remain "
                "visible for audit and are not inferred data refreshes."
            )
            evidence.append(f"within_tolerance={self.member_count}")
            anchor = "accepted-differences"
        elif kind is StoryKind.REPRESENTATION_NOISE:
            title = "Representation noise (display-identical)"
            description = (
                "Constants whose stored floating-point representation moved a few "
                "ULPs while every rendered value stays identical — a source "
                "re-export artifact, not a data change."
            )
            anchor = "noise"
        elif kind is StoryKind.INHERITED:
            title = "Pre-existing conditions carried from the baseline"
            description = (
                "Errors and consistency deviations proven identical in the "
                "baseline file (or reusing established historical patterns); "
                "not regressions introduced this cycle."
            )
            evidence.append(
                "provenance: "
                + ", ".join(
                    f"{name}={count}"
                    for name, count in sorted(self.provenances.items())
                )
            )
            anchor = "inherited"
        else:
            title = "Unexplained changes — review required"
            description = (
                "No systematic driver, recency, provenance, or noise evidence "
                "explains these findings. This is the primary review queue."
            )
            evidence.append(
                "unexplained by class: "
                + ", ".join(
                    f"{name}={count}"
                    for name, count in sorted(
                        self.residual_classes.items(),
                        key=lambda item: (-item[1], item[0]),
                    )[:_MAX_EVIDENCE_LINES]
                )
            )
            missing = sorted(self.missing_evidence)
            if missing:
                evidence.append("no evidence recorded for: " + ", ".join(missing))
            anchor = "residual"

        return ChangeStory(
            story_id=_stable_id(kind, anchor),
            kind=kind,
            title=title,
            description=description,
            evidence=tuple(evidence),
            finding_ids=tuple(self.finding_ids),
            member_count=self.member_count,
            severity_counts=_severity_names(self.severity_counts),
        )


class _StoryMachine:
    """Per-artifact-member replica of ``build_stories``' index machine."""

    _KINDS = tuple(StoryKind)
    _UNASSIGNED = -1

    def __init__(self) -> None:
        self.total = 0
        self.union = _UnionFind()
        self.token_owner: dict[str, int] = {}
        self.event_owner: dict[str, int] = {}
        self.driver_indices: set[int] = set()
        self.driver_tokens_by_index: dict[int, set[str]] = {}
        self.all_driver_tokens: dict[str, int] = {}
        self.wrapper_components: dict[str, int] = {}
        self.population_components: dict[str, int] = {}
        self.next_component = 0
        self.linked_formula_cells: dict[tuple[str, str], tuple[StoryKind, int]] = {}
        self.index = 0
        self.assigned_kind: list[int] = []
        self.assigned_component: list[int] = []
        self.states: dict[tuple[StoryKind, int], _StoryState] = {}

    # pass 1 -----------------------------------------------------------------
    def observe(self, finding: Finding) -> None:
        index = self.total
        self.total += 1
        if finding.finding_class in _DRIVER_CLASSES and not finding.expected_growth:
            self.union.add(index)
            self.driver_indices.add(index)
            tokens = _driver_tokens(finding)
            self.driver_tokens_by_index[index] = tokens
            for token in tokens:
                if token in self.token_owner:
                    self.union.union(self.token_owner[token], index)
                else:
                    self.token_owner[token] = index
            if finding.event_key:
                if finding.event_key in self.event_owner:
                    self.union.union(self.event_owner[finding.event_key], index)
                else:
                    self.event_owner[finding.event_key] = index

    def freeze(self) -> None:
        for index, tokens in self.driver_tokens_by_index.items():
            root = self.union.find(index)
            for token in tokens:
                self.all_driver_tokens[token] = root
        self.next_component = self.total + 1
        self.index = 0

    # pass 2 -----------------------------------------------------------------
    def classify(self, finding: Finding) -> None:
        index = self.index
        self.index += 1
        kind_component = self._classify(finding, index)
        if kind_component is None:
            self.assigned_kind.append(self._UNASSIGNED)
            self.assigned_component.append(0)
        else:
            kind, component = kind_component
            self.assigned_kind.append(self._KINDS.index(kind))
            self.assigned_component.append(component)

    def _classify(
        self, finding: Finding, index: int
    ) -> tuple[StoryKind, int] | None:
        if finding.materiality is Materiality.WITHIN_TOLERANCE:
            return (StoryKind.ACCEPTED_DIFFERENCES, 0)
        if finding.expected_growth:
            return (StoryKind.DATA_REFRESH, 0)
        if finding.subtype is FindingSubtype.AXIS_ROLLING_TURNOVER:
            return (StoryKind.DATA_REFRESH, 0)
        if finding.materiality is Materiality.NOISE:
            return (StoryKind.REPRESENTATION_NOISE, 0)
        if (
            finding.subtype is FindingSubtype.COLUMNAR_ERROR_POPULATION
            and finding.event_key
        ):
            component = self.population_components.setdefault(
                finding.event_key,
                self.next_component
                + len(self.wrapper_components)
                + len(self.population_components),
            )
            return (StoryKind.ERROR_POPULATION, component)
        if finding.provenance in (
            FindingProvenance.INHERITED,
            FindingProvenance.HISTORICAL_PATTERN,
        ):
            return (StoryKind.INHERITED, 0)
        if index in self.driver_indices:
            return (StoryKind.STRUCTURE_DRIVER, self.union.find(index))
        if finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED:
            added = _added_reference_tokens(finding)
            components = {
                component
                for token in added
                if (component := self.all_driver_tokens.get(token)) is not None
            }
            if components:
                assigned = (StoryKind.STRUCTURE_DRIVER, min(components))
                if finding.sheet and finding.location:
                    self.linked_formula_cells[
                        (finding.sheet, finding.location)
                    ] = assigned
                return assigned
            if (
                finding.subtype
                in (FindingSubtype.FORMULA_WRAPPED, FindingSubtype.FORMULA_UNWRAPPED)
                and finding.event_key
            ):
                component = self.wrapper_components.setdefault(
                    finding.event_key,
                    self.next_component + len(self.wrapper_components),
                )
                assigned = (StoryKind.STRUCTURE_DRIVER, component)
                if finding.sheet and finding.location:
                    self.linked_formula_cells[
                        (finding.sheet, finding.location)
                    ] = assigned
                return assigned
        if finding.finding_class in (
            FindingClass.ROW_KEY_CHANGED,
            FindingClass.COLUMN_KEY_CHANGED,
        ) and finding.subtype is FindingSubtype.AXIS_KEY_DERIVED_LABEL:
            return (StoryKind.DERIVED_LABELS, 0)
        if finding.temporal_context in (
            FindingTemporalContext.CURRENT_PERIOD,
            FindingTemporalContext.RECENT_WINDOW,
        ) or finding.materiality is Materiality.RECENT_RESTATEMENT:
            return (StoryKind.DATA_REFRESH, 0)
        return None

    # pass 3 -----------------------------------------------------------------
    def start_accumulation(self) -> None:
        self.index = 0

    def accumulate(self, finding: Finding) -> None:
        index = self.index
        self.index += 1
        kind_index = self.assigned_kind[index]
        if kind_index == self._UNASSIGNED:
            assigned = None
            if finding.finding_class in (
                FindingClass.NUMBER_FORMAT_CHANGED,
                FindingClass.STYLE_CHANGED,
            ):
                assigned = self.linked_formula_cells.get(
                    (finding.sheet or "", finding.location or "")
                )
            key = assigned if assigned is not None else (StoryKind.RESIDUAL, 0)
        else:
            key = (self._KINDS[kind_index], self.assigned_component[index])
        state = self.states.get(key)
        if state is None:
            state = self.states[key] = _StoryState(*key)
        state.add(finding)

    def stories(self) -> list[ChangeStory]:
        return [state.materialize() for state in self.states.values()]


class StoryStreamer:
    """``stream_stories`` with its first pass exposed as per-finding observes.

    ``record_run`` folds pass 1 into its existing main iteration; the two
    remaining machine passes stay sequential because classification and
    accumulation both replay the frozen component structure in order.
    """

    __slots__ = ("_machines",)

    def __init__(self) -> None:
        self._machines: dict[str, _StoryMachine] = {}

    def observe(self, finding: Finding) -> None:
        machine = self._machines.get(finding.artifact_member)
        if machine is None:
            machine = self._machines[finding.artifact_member] = _StoryMachine()
        machine.observe(finding)

    def finish(
        self,
        classify_pass: Iterable[Finding],
        accumulate_pass: Iterable[Finding],
    ) -> list[ChangeStory]:
        machines = self._machines
        for machine in machines.values():
            machine.freeze()
        for finding in classify_pass:
            machines[finding.artifact_member].classify(finding)
        for machine in machines.values():
            machine.start_accumulation()
        for finding in accumulate_pass:
            machines[finding.artifact_member].accumulate(finding)

        multi_member = len(machines) > 1
        stories: list[ChangeStory] = []
        for member_id in sorted(machines):
            member_stories = sorted(
                machines[member_id].stories(), key=_story_sort_key
            )
            if multi_member and member_id != "primary":
                member_stories = [
                    replace(
                        story,
                        story_id=_stable_id(
                            story.kind, f"{member_id}|{story.story_id}"
                        ),
                        title=f"{member_id} · {story.title}",
                    )
                    for story in member_stories
                ]
            stories.extend(member_stories)
        return sorted(stories, key=_story_sort_key)


def stream_stories(passes: Iterable[Iterable[Finding]]) -> list[ChangeStory]:
    """Streaming twin of ``build_stories`` over three ordered passes.

    ``passes`` must yield three iterations of the same triaged, already
    story-annotated findings in identical order.
    """
    iterator = iter(passes)
    streamer = StoryStreamer()
    for finding in next(iterator):
        streamer.observe(finding)
    return streamer.finish(next(iterator), next(iterator))
