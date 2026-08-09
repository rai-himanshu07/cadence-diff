"""Logical-series review lens over canonical decision groups.

The lens is a navigation view only. `build_pattern_groups` remains the
authoritative decision builder; nothing here changes finding identity, evidence
digests, canonical group IDs, review counts, or any public schema.
"""

from __future__ import annotations

import hashlib
import json
import re
from collections import defaultdict
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import TypeAlias

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple

from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingSubtype,
    SeriesAnchor,
)
from qc_tool.review import ReviewAnnotation, ReviewGroup

SeriesClusterKey: TypeAlias = tuple[str, str, str, str, int]

_A1_RE = re.compile(r"^\$?([A-Z]{1,3})\$?([1-9]\d*)$", re.IGNORECASE)


def _canonical_json(payload: object) -> bytes:
    return json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")


def anchor_binding_payload(
    finding_id: str,
    artifact_member: str,
    location: str,
    anchor: SeriesAnchor,
) -> dict[str, object]:
    """Exact payload bound by a sidecar entry digest."""
    return {
        "finding_id": finding_id,
        "artifact_member": artifact_member,
        "location": location,
        "anchor": anchor.model_dump(mode="json"),
    }


def canonical_series_anchor_digest(
    finding_id: str,
    artifact_member: str,
    location: str,
    anchor: SeriesAnchor,
) -> str:
    """Bind one anchor to its finding ID, package member and current locator."""
    payload = anchor_binding_payload(finding_id, artifact_member, location, anchor)
    return hashlib.sha256(_canonical_json(payload)).hexdigest()


def canonical_series_aggregate_digest(
    bindings: Mapping[str, tuple[str, str, SeriesAnchor]],
) -> str:
    """Bind the complete sorted anchor population of one run.

    ``bindings`` maps finding ID to ``(artifact_member, location, anchor)``.
    """
    payload = {
        finding_id: anchor_binding_payload(finding_id, member, location, anchor)
        for finding_id, (member, location, anchor) in sorted(bindings.items())
    }
    return hashlib.sha256(_canonical_json(payload)).hexdigest()


def anchor_segment(anchor: SeriesAnchor) -> str:
    """Which segment of its series an anchor claims; V1 predates the field."""
    return getattr(anchor, "segment", "restatement")


def series_anchor_segment(finding: Finding) -> str | None:
    """The one segment a finding may claim, or ``None`` when it may claim none."""
    if finding.finding_class is not FindingClass.VALUE_CHANGED:
        return None
    if finding.subtype is FindingSubtype.VALUE_REPLACEMENT:
        if finding.materiality is None or finding.temporal_context is None:
            return None
        return "restatement"
    if finding.subtype is FindingSubtype.VALUE_ADDED_POPULATION:
        return "new_period"
    if finding.subtype is FindingSubtype.VALUE_CLEARED_POPULATION:
        return "cleared_period"
    return None


def series_anchor_eligible(finding: Finding) -> bool:
    """Proved numeric restatements, first-time populations and single clears."""
    return series_anchor_segment(finding) is not None


def anchor_matches_finding(finding: Finding, anchor: SeriesAnchor) -> bool:
    """Require segment agreement, sheet identity and exact A1 index agreement."""
    segment = series_anchor_segment(finding)
    if segment is None or segment != anchor_segment(anchor):
        return False
    if (finding.sheet or "") != anchor.sheet:
        return False
    match = _A1_RE.fullmatch((finding.location or "").replace("$", ""))
    if match is None:
        return False
    try:
        row, col = coordinate_to_tuple(match.group(0).upper())
    except ValueError:
        return False
    if anchor.period_axis == "rows":
        return anchor.series_index == col and anchor.period_index == row
    return anchor.series_index == row and anchor.period_index == col


def series_cluster_key(finding: Finding, anchor: SeriesAnchor) -> SeriesClusterKey:
    """Exact runtime cluster identity; the member always comes from the finding."""
    return (
        finding.artifact_member,
        anchor.sheet,
        anchor.current_region_id,
        anchor.period_axis,
        anchor.series_index,
    )


def series_key_token(key: SeriesClusterKey) -> str:
    """Short stable token for derived row IDs."""
    return hashlib.sha256(_canonical_json(list(key))).hexdigest()[:10]


def series_position(key: SeriesClusterKey) -> str:
    """Axis position of a series, e.g. ``column D`` or ``row 7``."""
    _member, _sheet, _region, period_axis, series_index = key
    return (
        f"column {get_column_letter(series_index)}"
        if period_axis == "rows"
        else f"row {series_index}"
    )


def series_cluster_label(
    key: SeriesClusterKey,
    changed_periods: int,
    *,
    mixed_segments: bool = False,
) -> str:
    """Structural-only parent label: sheet, axis position and period count.

    No metric name, header text, cell value or excerpt content ever appears
    here; the anchor payload deliberately cannot supply one.
    """
    _member, sheet, _region, _period_axis, _series_index = key
    position = series_position(key)
    # "changed" would misdescribe a period that appeared or disappeared.
    noun = "period" if changed_periods == 1 else "periods"
    unit = f"{noun} affected" if mixed_segments else f"changed {noun}"
    return f"{sheet} · {position} · {changed_periods} {unit}"


@dataclass(frozen=True, slots=True)
class ReviewSlice:
    """One canonical group intersected with at most one logical series."""

    slice_id: str
    group_id: str
    members: tuple[Finding, ...]
    series_key: SeriesClusterKey | None = None
    #: True when the slice still holds every member of its canonical group.
    whole_group: bool = True

    @property
    def member_count(self) -> int:
        return len(self.members)


@dataclass(frozen=True, slots=True)
class SeriesCluster:
    """A promoted logical series holding child slices of several groups."""

    cluster_id: str
    key: SeriesClusterKey
    label: str
    slices: tuple[ReviewSlice, ...]
    members: tuple[Finding, ...]

    @property
    def member_count(self) -> int:
        return len(self.members)

    @property
    def sheet(self) -> str:
        return self.key[1]

    @property
    def artifact_member(self) -> str:
        return self.key[0]


@dataclass(frozen=True, slots=True)
class SeriesReviewLens:
    """Ordered lens rows plus lookup maps; always a lossless view."""

    rows: tuple[SeriesCluster | ReviewSlice, ...]
    clusters: tuple[SeriesCluster, ...]
    fallbacks: tuple[ReviewSlice, ...]
    cluster_by_id: Mapping[str, SeriesCluster]
    slice_by_id: Mapping[str, ReviewSlice]

    @property
    def promoted(self) -> bool:
        return bool(self.clusters)


class SeriesLensError(RuntimeError):
    """Raised when a lens would not be an exact permutation of its input."""


def build_series_review_lens(
    prioritized_groups: Sequence[ReviewGroup],
    anchors: Mapping[str, SeriesAnchor],
) -> SeriesReviewLens:
    """Split canonical groups into a lossless parent/child navigation view.

    A series is promoted only when its anchored findings span at least two
    canonical groups, so a one-group series never gains a redundant parent.
    """
    keyed: list[dict[str, SeriesClusterKey]] = []
    groups_per_key: dict[SeriesClusterKey, set[int]] = defaultdict(set)
    for priority, group in enumerate(prioritized_groups):
        per_group: dict[str, SeriesClusterKey] = {}
        for finding in group.members:
            anchor = anchors.get(finding.finding_id)
            if anchor is None or not anchor_matches_finding(finding, anchor):
                continue
            key = series_cluster_key(finding, anchor)
            per_group[finding.finding_id] = key
            groups_per_key[key].add(priority)
        keyed.append(per_group)

    promoted = {key for key, priorities in groups_per_key.items() if len(priorities) > 1}

    slices_by_key: dict[SeriesClusterKey, list[tuple[int, ReviewSlice]]] = defaultdict(
        list
    )
    fallbacks: list[tuple[int, ReviewSlice]] = []
    for priority, group in enumerate(prioritized_groups):
        per_group = keyed[priority]
        buckets: dict[SeriesClusterKey, list[Finding]] = defaultdict(list)
        remainder: list[Finding] = []
        for finding in group.members:
            key = per_group.get(finding.finding_id)
            if key is not None and key in promoted:
                buckets[key].append(finding)
            else:
                remainder.append(finding)
        for key in sorted(buckets, key=series_key_token):
            slices_by_key[key].append(
                (
                    priority,
                    ReviewSlice(
                        slice_id=f"{group.group_id}~{series_key_token(key)}",
                        group_id=group.group_id,
                        members=tuple(buckets[key]),
                        series_key=key,
                        whole_group=len(buckets[key]) == group.member_count,
                    ),
                )
            )
        if not remainder:
            continue
        whole_group = len(remainder) == group.member_count
        fallbacks.append(
            (
                priority,
                ReviewSlice(
                    slice_id=(
                        group.group_id if whole_group else f"{group.group_id}~rest"
                    ),
                    group_id=group.group_id,
                    members=tuple(remainder),
                    series_key=None,
                    whole_group=whole_group,
                ),
            )
        )

    clusters: list[tuple[int, SeriesCluster]] = []
    for key, entries in slices_by_key.items():
        # A parent reads chronologically, so a Critical new or cleared period
        # follows the earlier restatements instead of jumping above them.
        ordered_children = sorted(
            entries,
            key=lambda entry: (
                _earliest_period(entry[1], anchors),
                entry[0],
                entry[1].slice_id,
            ),
        )
        child_slices = tuple(item for _priority, item in ordered_children)
        members = tuple(
            member for child in child_slices for member in child.members
        )
        periods = {
            anchors[member.finding_id].period_index
            for member in members
            if member.finding_id in anchors
        }
        mixed_segments = any(
            anchor_segment(anchors[member.finding_id]) != "restatement"
            for member in members
            if member.finding_id in anchors
        )
        clusters.append(
            (
                min(priority for priority, _item in entries),
                SeriesCluster(
                    cluster_id=f"S{series_key_token(key)}",
                    key=key,
                    label=series_cluster_label(
                        key, len(periods), mixed_segments=mixed_segments
                    ),
                    slices=child_slices,
                    members=members,
                ),
            )
        )

    rows: list[tuple[int, str, SeriesCluster | ReviewSlice]] = [
        (priority, cluster.cluster_id, cluster) for priority, cluster in clusters
    ]
    rows.extend(
        (priority, item.slice_id, item) for priority, item in fallbacks
    )
    rows.sort(key=lambda entry: (entry[0], entry[1]))
    ordered = tuple(item for _priority, _identifier, item in rows)

    ordered_clusters = tuple(
        item for item in ordered if isinstance(item, SeriesCluster)
    )
    ordered_fallbacks = tuple(
        item for item in ordered if isinstance(item, ReviewSlice)
    )
    lens = SeriesReviewLens(
        rows=ordered,
        clusters=ordered_clusters,
        fallbacks=ordered_fallbacks,
        cluster_by_id={cluster.cluster_id: cluster for cluster in ordered_clusters},
        slice_by_id={
            child.slice_id: child
            for cluster in ordered_clusters
            for child in cluster.slices
        }
        | {item.slice_id: item for item in ordered_fallbacks},
    )
    _assert_lossless(prioritized_groups, lens)
    return lens


def _earliest_period(
    item: ReviewSlice, anchors: Mapping[str, SeriesAnchor]
) -> int:
    positions = [
        anchors[member.finding_id].period_index
        for member in item.members
        if member.finding_id in anchors
    ]
    return min(positions) if positions else 0


def _assert_lossless(
    prioritized_groups: Sequence[ReviewGroup], lens: SeriesReviewLens
) -> None:
    expected = [
        finding.finding_id for group in prioritized_groups for finding in group.members
    ]
    seen = [
        member.finding_id
        for cluster in lens.clusters
        for child in cluster.slices
        for member in child.members
    ] + [
        member.finding_id for item in lens.fallbacks for member in item.members
    ]
    if sorted(seen) != sorted(expected) or len(seen) != len(set(seen)):
        raise SeriesLensError(
            "series lens is not an exact permutation of its canonical findings"
        )


def cluster_confirmation_updates(
    visible_slices: Sequence[ReviewSlice],
    comment: str,
) -> list[ReviewAnnotation]:
    """Pure same-severity confirmations for visible, unreviewed findings.

    No finding is mutated here, and the severity is always written explicitly so
    a comment-free confirmation reloads as reviewed.
    """
    normalized = comment.strip()
    updates: list[ReviewAnnotation] = []
    seen: set[str] = set()
    for item in visible_slices:
        for finding in item.members:
            if finding.finding_id in seen:
                continue
            if finding.finding_class is FindingClass.FINDINGS_CAPPED:
                continue
            if finding.severity_overridden or finding.analyst_comment:
                continue
            if finding.severity is None:
                continue
            seen.add(finding.finding_id)
            updates.append(
                ReviewAnnotation(
                    finding_id=finding.finding_id,
                    severity=finding.severity.value,
                    comment=normalized,
                )
            )
    return updates
