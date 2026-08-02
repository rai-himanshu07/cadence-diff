"""Deterministic analyst review groups derived from atomic findings."""

from __future__ import annotations

import hashlib
import json
import re
from collections import defaultdict, deque
from dataclasses import dataclass, replace
from typing import TypeAlias

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple

from qc_tool.excel.formula_tokens import tokenize_formula
from qc_tool.findings import Finding, FindingClass, FindingSubtype, Severity

Coordinate: TypeAlias = tuple[int, int]
Rectangle: TypeAlias = tuple[int, int, int, int]
FindingIdentity: TypeAlias = tuple[str, ...]

_A1_RE = re.compile(r"^\$?[A-Z]{1,3}\$?[1-9]\d*$", re.IGNORECASE)
_FINDING_ID_RE = re.compile(r"^F(\d+)$")
_PRESENTATION_CLASSES = frozenset(
    {
        FindingClass.NUMBER_FORMAT_CHANGED,
        FindingClass.STYLE_CHANGED,
    }
)
_GROUPABLE_CLASSES = frozenset(
    {
        FindingClass.VALUE_CHANGED,
        FindingClass.FORMULA_ERROR,
        FindingClass.FORMULA_HARDCODED,
        FindingClass.FORMULA_REMOVED,
        FindingClass.FORMULA_MISSING,
        FindingClass.FORMULA_CACHE_MISSING,
        FindingClass.FORMULA_NOT_EXTENDED,
        FindingClass.FORMULA_LOGIC_CHANGED,
        FindingClass.FORMULA_INCONSISTENT,
        FindingClass.NUMBER_FORMAT_CHANGED,
        FindingClass.STYLE_CHANGED,
        FindingClass.REQUIRED_VALUE_MISSING,
        FindingClass.NUMERIC_BOUND_VIOLATION,
    }
)
_SEVERITY_RANK = {
    Severity.CRITICAL: 0,
    Severity.WARNING: 1,
    Severity.INFO: 2,
    Severity.EXPECTED: 3,
}


@dataclass(frozen=True, slots=True)
class ReviewGroup:
    """One analyst decision backed by one or more atomic findings."""

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
    members: tuple[Finding, ...]
    spatial: bool

    @property
    def member_count(self) -> int:
        return len(self.members)


@dataclass(frozen=True, slots=True)
class ReviewCounts:
    """Review decisions and their underlying atomic evidence by severity."""

    review_items: dict[Severity, int]
    atomic_findings: dict[Severity, int]


@dataclass(frozen=True, slots=True)
class ReviewAnnotation:
    """One atomic annotation update produced by a group review action."""

    finding_id: str
    severity: str | None
    comment: str


def finding_identity_key(finding: Finding) -> FindingIdentity:
    """Stable identity shared with Re-QC and review-group fingerprints."""
    return (
        finding.artifact,
        finding.finding_class.value,
        finding.sheet or "",
        finding.slide or "",
        finding.location or finding.baseline_location or "",
        finding.element or "",
    )


def _coordinate(location: str | None) -> Coordinate | None:
    if location is None or ":" in location or _A1_RE.fullmatch(location) is None:
        return None
    try:
        return coordinate_to_tuple(location.replace("$", ""))
    except ValueError:
        return None


def _baseline_mode(finding: Finding, current: Coordinate) -> tuple[object, ...]:
    if finding.baseline_location is None:
        return ("no-baseline",)
    baseline = _coordinate(finding.baseline_location)
    if baseline is None:
        return ("marker", finding.baseline_location)
    return ("translation", current[0] - baseline[0], current[1] - baseline[1])


def _producer_subtype(finding: Finding) -> str:
    if finding.finding_class is FindingClass.FORMULA_ERROR:
        value = finding.current_value or ""
        return "formula-text" if value.startswith("=") else "saved-value"
    return ""


def _partition_key(finding: Finding, current: Coordinate) -> tuple[object, ...]:
    severity = finding.severity or Severity.WARNING
    return (
        finding.artifact,
        finding.sheet or "",
        finding.finding_class.value,
        severity.value,
        finding.expected_growth,
        (finding.element or "").strip(),
        finding.waiver_reason,
        finding.waiver_expires,
        _producer_subtype(finding),
        _baseline_mode(finding, current),
    )


def _eligible(finding: Finding) -> Coordinate | None:
    if (
        finding.artifact != "excel"
        or finding.sheet is None
        or finding.finding_class not in _GROUPABLE_CLASSES
    ):
        return None
    return _coordinate(finding.location)


def _finding_order(finding: Finding) -> tuple[object, ...]:
    match = _FINDING_ID_RE.fullmatch(finding.finding_id)
    if match is not None:
        return (0, int(match.group(1)))
    return (1, finding_identity_key(finding), finding.finding_id)


def _horizontal_runs(columns: list[int]) -> list[tuple[int, int]]:
    if not columns:
        return []
    runs: list[tuple[int, int]] = []
    start = previous = columns[0]
    for column in columns[1:]:
        if column == previous + 1:
            previous = column
            continue
        runs.append((start, previous))
        start = previous = column
    runs.append((start, previous))
    return runs


def _rectangles(coordinates: set[Coordinate]) -> list[Rectangle]:
    by_row: dict[int, list[int]] = defaultdict(list)
    for row, column in coordinates:
        by_row[row].append(column)

    rectangles: list[Rectangle] = []
    active: dict[tuple[int, int], int] = {}
    previous_row: int | None = None
    for row in sorted(by_row):
        if previous_row is not None and row != previous_row + 1:
            rectangles.extend(
                (start_row, start_col, previous_row, end_col)
                for (start_col, end_col), start_row in active.items()
            )
            active.clear()
        runs = set(_horizontal_runs(sorted(set(by_row[row]))))
        for run in set(active).difference(runs):
            start_col, end_col = run
            rectangles.append((active.pop(run), start_col, row - 1, end_col))
        for run in runs.difference(active):
            active[run] = row
        previous_row = row

    if previous_row is not None:
        rectangles.extend(
            (start_row, start_col, previous_row, end_col)
            for (start_col, end_col), start_row in active.items()
        )
    return sorted(rectangles)


def _range(rectangle: Rectangle) -> str:
    min_row, min_col, max_row, max_col = rectangle
    start = f"{get_column_letter(min_col)}{min_row}"
    end = f"{get_column_letter(max_col)}{max_row}"
    return start if start == end else f"{start}:{end}"


def _bounding_range(coordinates: set[Coordinate]) -> str:
    if not coordinates:
        return ""
    rows = [row for row, _ in coordinates]
    columns = [column for _, column in coordinates]
    return _range((min(rows), min(columns), max(rows), max(columns)))


def _group_id(key: tuple[object, ...], members: tuple[Finding, ...]) -> str:
    identities = sorted(finding_identity_key(member) for member in members)
    payload = json.dumps(
        {"key": key, "members": identities},
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")
    return "G" + hashlib.sha256(payload).hexdigest()[:12].upper()


def _component_groups(
    key: tuple[object, ...],
    entries: list[tuple[Coordinate, Finding]],
) -> list[ReviewGroup]:
    by_coordinate: dict[Coordinate, list[Finding]] = defaultdict(list)
    for coordinate, finding in entries:
        by_coordinate[coordinate].append(finding)
    pending = set(by_coordinate)
    groups: list[ReviewGroup] = []
    while pending:
        first = min(pending)
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
                if neighbor not in pending:
                    continue
                pending.remove(neighbor)
                component.add(neighbor)
                queue.append(neighbor)

        members = tuple(
            sorted(
                (
                    finding
                    for coordinate in component
                    for finding in by_coordinate[coordinate]
                ),
                key=_finding_order,
            )
        )
        current_ranges = tuple(_range(item) for item in _rectangles(component))
        baseline_coordinates = {
            coordinate
            for member in members
            if (coordinate := _coordinate(member.baseline_location)) is not None
        }
        baseline_mixed = 0 < len(baseline_coordinates) < len(members)
        baseline_ranges = (
            tuple(_range(item) for item in _rectangles(baseline_coordinates))
            if len(baseline_coordinates) == len(members)
            else ()
        )
        representative = members[0]
        severity = representative.severity or Severity.WARNING
        groups.append(
            ReviewGroup(
                group_id=_group_id(key, members),
                finding_class=representative.finding_class,
                severity=severity,
                artifact=representative.artifact,
                sheet=representative.sheet,
                slide=representative.slide,
                element=representative.element or "",
                expected_growth=representative.expected_growth,
                ranges=current_ranges,
                bounding_range=_bounding_range(component),
                baseline_ranges=baseline_ranges,
                baseline_bounding_range=(
                    _bounding_range(baseline_coordinates)
                    if baseline_ranges
                    else ""
                ),
                baseline_mixed=baseline_mixed,
                members=members,
                spatial=True,
            )
        )
    return groups


def _singleton(finding: Finding) -> ReviewGroup:
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
    return ReviewGroup(
        group_id=_group_id(key, (finding,)),
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
        members=(finding,),
        spatial=False,
    )


def _review_sort_key(group: ReviewGroup) -> tuple[object, ...]:
    coordinate = _coordinate(group.bounding_range.partition(":")[0])
    return (
        _SEVERITY_RANK[group.severity],
        group.artifact,
        group.sheet or group.slide or "",
        coordinate or (10**9, 10**9),
        group.finding_class.value,
        group.element,
        group.group_id,
        tuple(
            (
                member.message,
                member.baseline_value or "",
                member.current_value or "",
                tuple(member.impacts),
                member.finding_id,
            )
            for member in group.members
        ),
    )


def build_review_groups(findings: list[Finding]) -> list[ReviewGroup]:
    """Build a lossless deterministic review layer over atomic findings."""
    partitions: dict[tuple[object, ...], list[tuple[Coordinate, Finding]]] = (
        defaultdict(list)
    )
    groups: list[ReviewGroup] = []
    for finding in findings:
        coordinate = _eligible(finding)
        if coordinate is None:
            groups.append(_singleton(finding))
            continue
        partitions[_partition_key(finding, coordinate)].append((coordinate, finding))
    for key, entries in partitions.items():
        groups.extend(_component_groups(key, entries))
    ordered = sorted(groups, key=_review_sort_key)
    seen: dict[str, int] = defaultdict(int)
    unique: list[ReviewGroup] = []
    for group in ordered:
        seen[group.group_id] += 1
        occurrence = seen[group.group_id]
        unique.append(
            group
            if occurrence == 1
            else replace(group, group_id=f"{group.group_id}-{occurrence}")
        )
    return unique


def review_counts(groups: list[ReviewGroup]) -> ReviewCounts:
    """Count analyst decisions and their atomic members by severity."""
    review_items = dict.fromkeys(Severity, 0)
    atomic_findings = dict.fromkeys(Severity, 0)
    for group in groups:
        review_items[group.severity] += 1
        atomic_findings[group.severity] += group.member_count
    return ReviewCounts(review_items=review_items, atomic_findings=atomic_findings)


# --- semantic pattern groups -------------------------------------------------


def _formula_signature(formula: str | None) -> tuple[str, ...]:
    """Token shape with concrete references and literals abstracted away."""
    if not formula:
        return ()
    try:
        tokens = tokenize_formula(formula)
    except Exception:  # malformed formulas keep their own stable shape
        return ("<unparsed>",)
    shape: list[str] = []
    for token in tokens:
        if token.type == "OPERAND" and token.subtype == "RANGE":
            shape.append("REF")
        elif token.type == "OPERAND" and token.subtype == "NUMBER":
            shape.append("NUM")
        elif token.type == "OPERAND" and token.subtype == "TEXT":
            shape.append("TEXT")
        else:
            shape.append(token.value.casefold())
    return tuple(shape)


def _transformation_shape(finding: Finding) -> str:
    if finding.finding_class is FindingClass.FORMULA_INCONSISTENT:
        # The dominant pattern stays exact so one group never mixes two patterns.
        payload: tuple[tuple[str, ...], tuple[str, ...]] = (
            (finding.baseline_value or "",),
            _formula_signature(finding.current_value),
        )
    elif finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED:
        if (
            finding.subtype
            in (FindingSubtype.FORMULA_WRAPPED, FindingSubtype.FORMULA_UNWRAPPED)
            and finding.event_key
        ):
            # One wrapper rollout = one skeleton, regardless of the varied
            # inner formulas it wrapped.
            payload = ((finding.event_key,), ())
        else:
            payload = (
                _formula_signature(finding.baseline_value),
                _formula_signature(finding.current_value),
            )
    elif finding.finding_class in _PRESENTATION_CLASSES:
        payload = ((finding.baseline_value or "",), (finding.current_value or "",))
    else:
        return ""
    encoded = json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()[:12]


def _baseline_translation_mode(finding: Finding) -> str:
    if finding.baseline_location is None:
        return "no-baseline"
    baseline = _coordinate(finding.baseline_location)
    current = _coordinate(finding.location)
    if baseline is None or current is None:
        return "marker"
    return "in-place" if baseline == current else "translated"


def _pattern_key(finding: Finding) -> tuple[object, ...]:
    severity = finding.severity or Severity.WARNING
    population_event = (
        finding.event_key
        if finding.subtype is FindingSubtype.COLUMNAR_ERROR_POPULATION
        else ""
    )
    return (
        finding.artifact,
        finding.sheet or "",
        finding.slide or "",
        finding.slide_index or 0,
        finding.baseline_slide_index or 0,
        finding.finding_class.value,
        severity.value,
        finding.expected_growth,
        finding.expected_reason.value if finding.expected_reason is not None else "",
        finding.provenance.value if finding.provenance is not None else "",
        finding.subtype.value if finding.subtype is not None else "",
        finding.materiality.value if finding.materiality is not None else "",
        (
            finding.temporal_context.value
            if finding.temporal_context is not None
            else ""
        ),
        tuple(sorted(tag.value for tag in finding.evidence_tags)),
        population_event,
        _transformation_shape(finding),
        _baseline_translation_mode(finding),
        finding.waiver_reason,
        finding.waiver_expires,
    )


def _pattern_group(key: tuple[object, ...], members: tuple[Finding, ...]) -> ReviewGroup:
    representative = members[0]
    coordinates = {
        coordinate
        for member in members
        if (coordinate := _coordinate(member.location)) is not None
    }
    baseline_coordinates = {
        coordinate
        for member in members
        if (coordinate := _coordinate(member.baseline_location)) is not None
    }
    ranges = tuple(_range(item) for item in _rectangles(coordinates))
    baseline_ranges = (
        tuple(_range(item) for item in _rectangles(baseline_coordinates))
        if len(baseline_coordinates) == len(members)
        else ()
    )
    return ReviewGroup(
        group_id=_group_id(key, members),
        finding_class=representative.finding_class,
        severity=representative.severity or Severity.WARNING,
        artifact=representative.artifact,
        sheet=representative.sheet,
        slide=representative.slide,
        element=representative.element or "",
        expected_growth=representative.expected_growth,
        ranges=ranges,
        bounding_range=(
            _bounding_range(coordinates)
            if coordinates
            else representative.location or representative.baseline_location or ""
        ),
        baseline_ranges=baseline_ranges,
        baseline_bounding_range=(
            _bounding_range(baseline_coordinates) if baseline_ranges else ""
        ),
        baseline_mixed=0 < len(baseline_coordinates) < len(members),
        members=members,
        spatial=False,
    )


def build_pattern_groups(findings: list[Finding]) -> list[ReviewGroup]:
    """Semantically pure review groups keyed on meaning, not adjacency.

    One group may span disconnected ranges. Every member shares its finding
    class, severity, expected state, provenance, subtype, transformation shape,
    baseline-translation mode, and waiver state.
    """
    partitions: dict[tuple[object, ...], list[Finding]] = defaultdict(list)
    for finding in findings:
        partitions[_pattern_key(finding)].append(finding)
    groups = [
        _pattern_group(key, tuple(sorted(members, key=_finding_order)))
        for key, members in partitions.items()
    ]
    ordered = sorted(groups, key=_review_sort_key)
    seen: dict[str, int] = defaultdict(int)
    unique: list[ReviewGroup] = []
    for group in ordered:
        seen[group.group_id] += 1
        occurrence = seen[group.group_id]
        unique.append(
            group
            if occurrence == 1
            else replace(group, group_id=f"{group.group_id}-{occurrence}")
        )
    return unique


def count_pattern_groups(groups: list[ReviewGroup]) -> ReviewCounts:
    """Count semantic pattern decisions and their atomic members by severity."""
    return review_counts(groups)


def format_group_ranges(group: ReviewGroup, *, max_spans: int = 3) -> str:
    """Compact exact group geometry for human-facing tables."""
    if not group.ranges:
        return group.bounding_range
    if len(group.ranges) <= max_spans:
        return "; ".join(group.ranges)
    return f"{'; '.join(group.ranges[:max_spans])}; +{len(group.ranges) - max_spans} spans"


def apply_group_review(
    group: ReviewGroup,
    *,
    severity: Severity | None,
    comment: str,
    replace_existing: bool,
) -> list[ReviewAnnotation]:
    """Apply one explicit group decision and return changed atomic annotations."""
    updates: list[ReviewAnnotation] = []
    normalized_comment = comment.strip()
    for finding in group.members:
        severity_changed = bool(
            severity is not None
            and (replace_existing or not finding.severity_overridden)
        )
        comment_changed = bool(
            normalized_comment
            and (replace_existing or not finding.analyst_comment)
        )
        if severity_changed:
            finding.severity = severity
            finding.severity_overridden = True
        if comment_changed:
            finding.analyst_comment = normalized_comment
        if not severity_changed and not comment_changed:
            continue
        updates.append(
            ReviewAnnotation(
                finding_id=finding.finding_id,
                severity=(
                    finding.severity.value
                    if finding.severity_overridden and finding.severity is not None
                    else None
                ),
                comment=finding.analyst_comment,
            )
        )
    return updates
