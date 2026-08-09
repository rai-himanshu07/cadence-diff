"""Change stories: cross-class narratives linking findings to shared drivers.

A story is a lens, never a truth change: severities, counts, findings, and
groups are untouched. Every linkage cites its evidence, and anything the
engine cannot explain lands in the residual story — the analyst's true
review queue. Membership is a partition: each finding belongs to exactly
one story.
"""

from __future__ import annotations

import hashlib
import re
from collections import Counter, defaultdict
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, replace
from enum import StrEnum

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

#: Structural events that can drive coordinated formula rollouts.
_DRIVER_CLASSES = frozenset(
    {
        FindingClass.TABLE_STRUCTURE_CHANGED,
        FindingClass.NAMED_RANGE_CHANGED,
        FindingClass.DATA_VALIDATION_CHANGED,
        FindingClass.COLUMN_INSERTED,
        FindingClass.COLUMN_DELETED,
    }
)

_MAX_EVIDENCE_LINES = 8

#: Structured reference (table[column]) or plain identifier of length >= 3.
_TOKEN_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*\[[^\]]+\]|[A-Za-z_][A-Za-z0-9_.]{2,}")

_TOKEN_STOPWORDS = frozenset(
    {
        "added",
        "and",
        "changed",
        "column",
        "columns",
        "count",
        "criteria",
        "data",
        "display",
        "extended",
        "list",
        "named",
        "new",
        "range",
        "removed",
        "renamed",
        "reordered",
        "repointed",
        "rule",
        "settings",
        "table",
        "text",
        "the",
        "validation",
        "with",
    }
)


class StoryKind(StrEnum):
    STRUCTURE_DRIVER = "structure_driver"
    ERROR_POPULATION = "error_population"
    DERIVED_LABELS = "derived_labels"
    DATA_REFRESH = "data_refresh"
    ACCEPTED_DIFFERENCES = "accepted_differences"
    REPRESENTATION_NOISE = "representation_noise"
    INHERITED = "inherited"
    RESIDUAL = "residual"


@dataclass(frozen=True, slots=True)
class ChangeStory:
    story_id: str
    kind: StoryKind
    title: str
    description: str
    evidence: tuple[str, ...]
    finding_ids: tuple[str, ...]
    member_count: int
    severity_counts: dict[str, int]


def _tokens(text: str | None) -> set[str]:
    if not text:
        return set()
    found = set()
    for match in _TOKEN_RE.findall(text):
        token = match.casefold().strip()
        if token in _TOKEN_STOPWORDS:
            continue
        found.add(token)
        if "[" in token:
            found.add(token.split("[", 1)[0])
    return found


def _driver_tokens(finding: Finding) -> set[str]:
    return _tokens(finding.element) | _tokens(finding.current_value)


def _added_reference_tokens(finding: Finding) -> set[str]:
    return _tokens(finding.current_value) - _tokens(finding.baseline_value)


class _UnionFind:
    def __init__(self) -> None:
        self._parent: dict[int, int] = {}

    def add(self, item: int) -> None:
        self._parent.setdefault(item, item)

    def find(self, item: int) -> int:
        parent = self._parent[item]
        if parent != item:
            self._parent[item] = parent = self.find(parent)
        return parent

    def union(self, a: int, b: int) -> None:
        self.add(a)
        self.add(b)
        root_a, root_b = self.find(a), self.find(b)
        if root_a != root_b:
            self._parent[max(root_a, root_b)] = min(root_a, root_b)


def _stable_id(kind: StoryKind, anchor: str) -> str:
    digest = hashlib.sha256(f"{kind.value}|{anchor}".encode()).hexdigest()[:12]
    return f"S{digest.upper()}"


def _severity_counts(members: list[Finding]) -> dict[str, int]:
    counts = Counter(
        (member.severity or Severity.WARNING).value for member in members
    )
    return _severity_names(counts)


def _severity_names(counts: Counter[str]) -> dict[str, int]:
    return {severity.value: counts.get(severity.value, 0) for severity in Severity}


def _story_sort_key(story: ChangeStory) -> tuple[int, int, str]:
    kind_rank = {
        StoryKind.STRUCTURE_DRIVER: 0,
        StoryKind.ERROR_POPULATION: 1,
        StoryKind.DERIVED_LABELS: 2,
        StoryKind.DATA_REFRESH: 3,
        StoryKind.ACCEPTED_DIFFERENCES: 4,
        StoryKind.REPRESENTATION_NOISE: 5,
        StoryKind.INHERITED: 6,
        StoryKind.RESIDUAL: 7,
    }[story.kind]
    return (kind_rank, -story.member_count, story.story_id)


def _clip(text: str, limit: int = 140) -> str:
    return text if len(text) <= limit else text[: limit - 1] + "\u2026"


@dataclass(slots=True)
class StoryEvidenceContext:
    """Cross-finding evidence collected in one pass, applied per finding."""

    driver_tokens: set[tuple[str, str]]
    colocated_cells: set[tuple[str, str, str]]

    @classmethod
    def collect(cls, findings: Iterable[Finding]) -> StoryEvidenceContext:
        collector = StoryEvidenceCollector()
        for finding in findings:
            collector.observe(finding)
        return collector.context()

    def apply(self, finding: Finding) -> None:
        """Idempotent per-finding twin of the batch annotation."""
        key = (
            finding.artifact_member,
            finding.sheet or "",
            finding.location or "",
        )
        if finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED:
            added = _added_reference_tokens(finding)
            if {(finding.artifact_member, token) for token in added} & (
                self.driver_tokens
            ):
                finding.evidence_tags.update(
                    {
                        FindingEvidenceTag.ADDED_REFERENCE,
                        FindingEvidenceTag.RESOLVED_DRIVER,
                    }
                )
            if key in self.colocated_cells:
                finding.evidence_tags.add(FindingEvidenceTag.EXACT_COLOCATION)
        elif finding.finding_class in (
            FindingClass.NUMBER_FORMAT_CHANGED,
            FindingClass.STYLE_CHANGED,
        ):
            if key in self.colocated_cells:
                finding.evidence_tags.add(FindingEvidenceTag.EXACT_COLOCATION)


class StoryEvidenceCollector:
    """Incremental twin of ``StoryEvidenceContext.collect`` for streamed parts."""

    __slots__ = ("_driver_tokens", "_formula_cells", "_style_cells")

    def __init__(self) -> None:
        self._driver_tokens: set[tuple[str, str]] = set()
        self._formula_cells: list[tuple[tuple[str, str, str], set[str], bool]] = []
        self._style_cells: set[tuple[str, str, str]] = set()

    def observe(self, finding: Finding) -> None:
        if (
            finding.finding_class in _DRIVER_CLASSES
            and not finding.expected_growth
        ):
            self._driver_tokens.update(
                (finding.artifact_member, token)
                for token in _driver_tokens(finding)
            )
        elif finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED:
            if finding.sheet and finding.location:
                self._formula_cells.append(
                    (
                        (
                            finding.artifact_member,
                            finding.sheet,
                            finding.location,
                        ),
                        _added_reference_tokens(finding),
                        FindingEvidenceTag.RESOLVED_DRIVER
                        in finding.evidence_tags
                        or (
                            finding.subtype
                            in (
                                FindingSubtype.FORMULA_WRAPPED,
                                FindingSubtype.FORMULA_UNWRAPPED,
                            )
                            and bool(finding.event_key)
                        ),
                    )
                )
        elif finding.finding_class in (
            FindingClass.NUMBER_FORMAT_CHANGED,
            FindingClass.STYLE_CHANGED,
        ):
            self._style_cells.add(
                (
                    finding.artifact_member,
                    finding.sheet or "",
                    finding.location or "",
                )
            )

    def context(self) -> StoryEvidenceContext:
        linked_cells = {
            key
            for key, added, already_linked in self._formula_cells
            if already_linked
            or {(key[0], token) for token in added} & self._driver_tokens
        }
        return StoryEvidenceContext(
            driver_tokens=self._driver_tokens,
            colocated_cells=linked_cells & self._style_cells,
        )


def annotate_story_evidence(findings: Sequence[Finding]) -> None:
    """Attach direct cross-finding evidence used by deterministic story edges."""
    context = StoryEvidenceContext.collect(findings)
    for finding in findings:
        context.apply(finding)


def build_stories(findings: Sequence[Finding]) -> list[ChangeStory]:
    """Partition triaged findings into deterministic, evidence-cited stories."""
    member_ids = sorted({finding.artifact_member for finding in findings})
    if len(member_ids) > 1:
        stories: list[ChangeStory] = []
        for member_id in member_ids:
            member_stories = build_stories(
                [
                    finding
                    for finding in findings
                    if finding.artifact_member == member_id
                ]
            )
            if member_id == "primary":
                stories.extend(member_stories)
                continue
            stories.extend(
                replace(
                    story,
                    story_id=_stable_id(
                        story.kind,
                        f"{member_id}|{story.story_id}",
                    ),
                    title=f"{member_id} · {story.title}",
                )
                for story in member_stories
            )
        return sorted(stories, key=_story_sort_key)
    annotate_story_evidence(findings)
    assigned: dict[int, tuple[StoryKind, int]] = {}  # index -> (kind, component)

    # --- structural drivers: connected components over explicit evidence -----
    driver_indices = [
        index
        for index, finding in enumerate(findings)
        if finding.finding_class in _DRIVER_CLASSES and not finding.expected_growth
    ]
    union = _UnionFind()
    token_owner: dict[str, int] = {}
    event_owner: dict[str, int] = {}
    driver_token_map: dict[int, set[str]] = {}
    for index in driver_indices:
        union.add(index)
        finding = findings[index]
        tokens = _driver_tokens(finding)
        driver_token_map[index] = tokens
        for token in tokens:
            if token in token_owner:
                union.union(token_owner[token], index)
            else:
                token_owner[token] = index
        if finding.event_key:
            if finding.event_key in event_owner:
                union.union(event_owner[finding.event_key], index)
            else:
                event_owner[finding.event_key] = index

    all_driver_tokens: dict[str, int] = {}
    for index, tokens in driver_token_map.items():
        for token in tokens:
            all_driver_tokens[token] = union.find(index)

    # --- pass 1: per-finding classification ---------------------------------
    wrapper_components: dict[str, int] = {}
    population_components: dict[str, int] = {}
    next_component = len(findings) + 1
    linked_formula_cells: dict[tuple[str, str], tuple[StoryKind, int]] = {}

    for index, finding in enumerate(findings):
        if finding.materiality is Materiality.WITHIN_TOLERANCE:
            assigned[index] = (StoryKind.ACCEPTED_DIFFERENCES, 0)
            continue
        if finding.expected_growth:
            assigned[index] = (StoryKind.DATA_REFRESH, 0)
            continue
        if finding.subtype is FindingSubtype.AXIS_ROLLING_TURNOVER:
            # The oldest position leaving a detected rolling window is that
            # window's cadence design.
            assigned[index] = (StoryKind.DATA_REFRESH, 0)
            continue
        if finding.materiality is Materiality.NOISE:
            assigned[index] = (StoryKind.REPRESENTATION_NOISE, 0)
            continue
        if (
            finding.subtype is FindingSubtype.COLUMNAR_ERROR_POPULATION
            and finding.event_key
        ):
            component = population_components.setdefault(
                finding.event_key,
                next_component + len(wrapper_components) + len(population_components),
            )
            assigned[index] = (StoryKind.ERROR_POPULATION, component)
            continue
        if finding.provenance in (
            FindingProvenance.INHERITED,
            FindingProvenance.HISTORICAL_PATTERN,
        ):
            assigned[index] = (StoryKind.INHERITED, 0)
            continue
        if index in driver_token_map:
            assigned[index] = (StoryKind.STRUCTURE_DRIVER, union.find(index))
            continue
        if finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED:
            added = _added_reference_tokens(finding)
            components = {
                component
                for token in added
                if (component := all_driver_tokens.get(token)) is not None
            }
            if components:
                component = min(components)
                assigned[index] = (StoryKind.STRUCTURE_DRIVER, component)
                if finding.sheet and finding.location:
                    linked_formula_cells[
                        (finding.sheet, finding.location)
                    ] = assigned[index]
                continue
            if (
                finding.subtype
                in (FindingSubtype.FORMULA_WRAPPED, FindingSubtype.FORMULA_UNWRAPPED)
                and finding.event_key
            ):
                component = wrapper_components.setdefault(
                    finding.event_key, next_component + len(wrapper_components)
                )
                assigned[index] = (StoryKind.STRUCTURE_DRIVER, component)
                if finding.sheet and finding.location:
                    linked_formula_cells[
                        (finding.sheet, finding.location)
                    ] = assigned[index]
                continue
        if finding.finding_class in (
            FindingClass.ROW_KEY_CHANGED,
            FindingClass.COLUMN_KEY_CHANGED,
        ) and finding.subtype is FindingSubtype.AXIS_KEY_DERIVED_LABEL:
            assigned[index] = (StoryKind.DERIVED_LABELS, 0)
            continue
        if finding.temporal_context in (
            FindingTemporalContext.CURRENT_PERIOD,
            FindingTemporalContext.RECENT_WINDOW,
        ) or finding.materiality is Materiality.RECENT_RESTATEMENT:
            assigned[index] = (StoryKind.DATA_REFRESH, 0)
            continue

    # --- pass 2: co-located presentation changes join their formula story ---
    for index, finding in enumerate(findings):
        if index in assigned:
            continue
        if finding.finding_class in (
            FindingClass.NUMBER_FORMAT_CHANGED,
            FindingClass.STYLE_CHANGED,
        ):
            key = (finding.sheet or "", finding.location or "")
            story = linked_formula_cells.get(key)
            if story is not None:
                assigned[index] = story
                continue
        assigned[index] = (StoryKind.RESIDUAL, 0)

    # --- materialize stories -------------------------------------------------
    buckets: dict[tuple[StoryKind, int], list[int]] = defaultdict(list)
    for index, key in assigned.items():
        buckets[key].append(index)

    stories: list[ChangeStory] = []
    for (kind, component), indices in buckets.items():
        members = [findings[i] for i in sorted(indices)]
        stories.append(_materialize(kind, component, members))

    return sorted(stories, key=_story_sort_key)


def _missing_evidence(finding: Finding) -> tuple[str, ...]:
    """Which explanation axes a residual finding carries no evidence on."""
    absent: list[str] = []
    if finding.provenance is None:
        absent.append("provenance")
    if finding.materiality is None:
        absent.append("materiality")
    if finding.temporal_context is None:
        absent.append("temporal context")
    if not finding.event_key:
        absent.append("shared event")
    return tuple(absent)


def _materialize(
    kind: StoryKind,
    component: int,
    members: list[Finding],
) -> ChangeStory:
    finding_ids = tuple(member.finding_id for member in members if member.finding_id)
    drivers = [
        member for member in members if member.finding_class in _DRIVER_CLASSES
    ]
    formula_members = [
        member
        for member in members
        if member.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
    ]
    evidence: list[str] = []
    if kind is StoryKind.STRUCTURE_DRIVER:
        for driver in drivers[:_MAX_EVIDENCE_LINES]:
            evidence.append(f"driver: {_clip(driver.message)}")
        skeletons = sorted(
            {m.event_key for m in formula_members if m.event_key}
        )
        if formula_members:
            evidence.append(
                f"{len(formula_members)} formula change(s) reference the new "
                f"structure or share {len(skeletons)} wrapper skeleton(s)"
            )
        elements = sorted(
            {d.element for d in drivers if d.element}
        )
        subject = ", ".join(elements[:3]) if elements else "systematic formula rollout"
        title = f"Coordinated structure/formula change: {subject}"
        description = (
            "Structural edits and the formula changes that reference them, "
            "reviewed as one deliberate change."
            if drivers
            else "A systematic formula rewrite sharing one wrapper skeleton; "
            "no structural driver was identified."
        )
        anchor = ",".join(
            sorted(
                token for driver in drivers for token in _driver_tokens(driver)
            )
        ) or (sorted({m.event_key for m in formula_members if m.event_key}) or [str(component)])[0]
    elif kind is StoryKind.ERROR_POPULATION:
        title = f"Error population: {len(members)} atomic findings"
        description = (
            "A proven same-artifact, sheet, column, and literal population. "
            "Grouping identifies one incident; it does not imply reduced risk."
        )
        population_tags = {
            tag
            for member in members
            for tag in member.evidence_tags
            if tag
            in {
                FindingEvidenceTag.CONCENTRATED_POPULATION,
                FindingEvidenceTag.CONTIGUOUS_POPULATION,
                FindingEvidenceTag.SPARSE_MASS_POPULATION,
            }
        }
        if population_tags:
            evidence.append(
                "evidence: "
                + ", ".join(tag.value for tag in sorted(population_tags, key=str))
            )
        evidence.extend(
            _clip(member.message) for member in members[:_MAX_EVIDENCE_LINES]
        )
        anchor = next(
            (member.event_key for member in members if member.event_key),
            str(component),
        )
    elif kind is StoryKind.DERIVED_LABELS:
        title = "Derived labels re-resolved"
        description = (
            "Row/column keys are formula-derived; their upstream driver "
            "changed, not the historical rows themselves."
        )
        for member in members[:_MAX_EVIDENCE_LINES]:
            evidence.append(_clip(member.message))
        anchor = "derived-labels"
    elif kind is StoryKind.DATA_REFRESH:
        title = "New-cycle data arrival and in-window restatements"
        description = (
            "Expected cadence growth plus constant changes confined to the "
            "trailing restatement window or declared acceptance bands."
        )
        tiers = Counter(
            member.materiality.value for member in members if member.materiality
        )
        if tiers:
            evidence.append(
                "tiers: "
                + ", ".join(f"{name}={count}" for name, count in sorted(tiers.items()))
            )
        anchor = "data-refresh"
    elif kind is StoryKind.ACCEPTED_DIFFERENCES:
        title = "Accepted differences within declared tolerance"
        description = (
            "Changes inside explicit analyst acceptance bounds. They remain "
            "visible for audit and are not inferred data refreshes."
        )
        evidence.append(f"within_tolerance={len(members)}")
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
        provenances = Counter(
            member.provenance.value for member in members if member.provenance
        )
        evidence.append(
            "provenance: "
            + ", ".join(f"{name}={count}" for name, count in sorted(provenances.items()))
        )
        anchor = "inherited"
    else:
        title = "Unexplained changes — review required"
        description = (
            "No systematic driver, recency, provenance, or noise evidence "
            "explains these findings. This is the primary review queue."
        )
        classes = Counter(member.finding_class.value for member in members)
        evidence.append(
            "unexplained by class: "
            + ", ".join(
                f"{name}={count}"
                for name, count in sorted(
                    classes.items(), key=lambda item: (-item[1], item[0])
                )[:_MAX_EVIDENCE_LINES]
            )
        )
        missing = sorted(
            {
                reason
                for member in members
                for reason in _missing_evidence(member)
            }
        )
        if missing:
            evidence.append("no evidence recorded for: " + ", ".join(missing))
        anchor = "residual"

    return ChangeStory(
        story_id=_stable_id(kind, anchor),
        kind=kind,
        title=title,
        description=description,
        evidence=tuple(evidence),
        finding_ids=finding_ids,
        member_count=len(members),
        severity_counts=_severity_counts(members),
    )
