"""Standalone HTML report rendered from the triaged `QCRunResult`.

Self-contained (inline CSS, vanilla JS, no external assets) and
autoescaped — finding text originates from client files and must never
inject markup. Decisions, stories, and coverage are always complete;
the embedded atomic-member detail is inlined only up to
``ATOMICS_INLINE_THRESHOLD`` rows, past which a disclosure points at the
Excel/JSON exports and the app (an HTML file near the threshold reaches
hundreds of MB and may not open in a browser).
"""

import datetime as dt
import logging
from collections.abc import Iterator, Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import TypedDict

from jinja2 import Environment, FileSystemLoader, select_autoescape

from qc_tool.coverage import capability_limited
from qc_tool.engine import QCRunResult
from qc_tool.excel.formulas import formula_token_diff
from qc_tool.findings import Finding, FindingClass, Severity
from qc_tool.findings_store import finding_by_id
from qc_tool.review import population_summary_text
from qc_tool.review_stream import (
    GroupSummary,
    counts_from_summaries,
    stream_stories,
    summarize_pattern_groups,
)
from qc_tool.security import private_directory, private_file

logger = logging.getLogger(__name__)

#: User decision (2026-08-08): inline atomics up to one million rows.
ATOMICS_INLINE_THRESHOLD = 1_000_000


class _HtmlMember(TypedDict):
    finding_id: str
    severity: str
    overridden: bool
    location: str
    baseline: str
    current: str
    message: str
    impacts: str
    comment: str
    element: str
    root: str
    provenance: str
    subtype: str
    materiality: str
    temporal_context: str
    expected_reason: str
    evidence_tags: str
    waiver: str
    formula_diff: list[dict]
    artifact_member: str
    population: str


@dataclass(slots=True)
class _SingletonMember:
    message: str


@dataclass(slots=True)
class _GroupView:
    """Exactly the group attributes the template touches, summary-backed."""

    group_id: str
    severity: Severity
    finding_class: FindingClass
    artifact_member: str
    sheet: str | None
    slide: str | None
    ranges: tuple[str, ...]
    bounding_range: str
    member_count: int
    members: tuple[_SingletonMember, ...]


def _group_views(
    result: QCRunResult, summaries: list[GroupSummary]
) -> list[_GroupView]:
    views: list[_GroupView] = []
    for summary in summaries:
        members: tuple[_SingletonMember, ...] = ()
        if summary.member_count == 1:
            first = finding_by_id(result.findings, summary.member_finding_ids[0])
            members = (_SingletonMember(first.message if first else ""),)
        views.append(
            _GroupView(
                group_id=summary.group_id,
                severity=summary.severity,
                finding_class=summary.finding_class,
                artifact_member=summary.artifact_member,
                sheet=summary.sheet,
                slide=summary.slide,
                ranges=summary.ranges,
                bounding_range=summary.bounding_range,
                member_count=summary.member_count,
                members=members,
            )
        )
    return views


def _html_member(finding: Finding) -> _HtmlMember:
    return {
        "finding_id": finding.finding_id,
        "artifact_member": finding.artifact_member,
        "severity": finding.severity.value if finding.severity else "warning",
        "overridden": finding.severity_overridden,
        "location": finding.location or finding.baseline_location or "",
        "baseline": finding.baseline_value or "",
        "current": finding.current_value or "",
        "message": finding.message,
        "impacts": "; ".join(finding.impacts),
        "comment": finding.analyst_comment,
        "element": finding.element or "",
        "root": finding.root_cause_key,
        "provenance": (
            finding.provenance.value if finding.provenance is not None else ""
        ),
        "subtype": finding.subtype.value if finding.subtype is not None else "",
        "materiality": (
            finding.materiality.value if finding.materiality is not None else ""
        ),
        "temporal_context": (
            finding.temporal_context.value
            if finding.temporal_context is not None
            else ""
        ),
        "expected_reason": (
            finding.expected_reason.value
            if finding.expected_reason is not None
            else ""
        ),
        "evidence_tags": "; ".join(
            sorted(tag.value for tag in finding.evidence_tags)
        ),
        "waiver": (
            f"{finding.waiver_reason} (expires {finding.waiver_expires})"
            if finding.waiver_reason
            else ""
        ),
        "formula_diff": (
            [
                {"text": seg.text, "kind": seg.kind.value}
                for seg in formula_token_diff(
                    finding.baseline_value,
                    finding.current_value,
                    finding.baseline_location or "",
                    finding.location or "",
                )
            ]
            if finding.finding_class.value == "formula_logic_changed"
            else []
        ),
        "population": population_summary_text(finding),
    }


class _LazyMemberPayload(Mapping[str, list[_HtmlMember]]):
    """Hydrates one group's members per template access, never the whole run.

    The template reads ``member_payload[group.group_id]`` exactly once per
    group; hydrating on access keeps peak memory at one group's members
    instead of every finding in the run.
    """

    def __init__(
        self, result: QCRunResult, summaries: list[GroupSummary]
    ) -> None:
        self._result = result
        self._summaries = {summary.group_id: summary for summary in summaries}

    def __getitem__(self, group_id: str) -> list[_HtmlMember]:
        summary = self._summaries[group_id]
        members: list[_HtmlMember] = []
        for finding_id in summary.member_finding_ids:
            finding = finding_by_id(self._result.findings, finding_id)
            if finding is not None:
                members.append(_html_member(finding))
        return members

    def __iter__(self) -> Iterator[str]:
        return iter(self._summaries)

    def __len__(self) -> int:
        return len(self._summaries)


def _member_payload(
    result: QCRunResult, summaries: list[GroupSummary]
) -> dict[str, list[_HtmlMember]]:
    lazy = _LazyMemberPayload(result, summaries)
    return {group_id: lazy[group_id] for group_id in lazy}

_ENV = Environment(
    loader=FileSystemLoader(Path(__file__).parent / "templates"),
    autoescape=select_autoescape(enabled_extensions=("j2", "html")),
)


def _render_context(result: QCRunResult) -> dict[str, object]:
    summaries = summarize_pattern_groups(result.findings)
    atomics_included = len(result.findings) <= ATOMICS_INLINE_THRESHOLD
    return {
        "result": result,
        "groups": _group_views(result, summaries),
        "stories": stream_stories(
            [iter(result.findings), iter(result.findings), iter(result.findings)]
        ),
        "member_payload": (
            _LazyMemberPayload(result, summaries) if atomics_included else {}
        ),
        "atomics_included": atomics_included,
        "atomics_threshold": ATOMICS_INLINE_THRESHOLD,
        "counts": counts_from_summaries(summaries),
        "capability_limited": capability_limited(result.coverage),
        "generated_at": dt.datetime.now(dt.UTC).isoformat(timespec="seconds"),
    }


def render_html_report(result: QCRunResult) -> str:
    template = _ENV.get_template("report.html.j2")
    return template.render(**_render_context(result))


def write_html_report(result: QCRunResult, path: Path) -> None:
    """Stream the rendered template to disk chunk by chunk.

    ``Template.render`` is ``concat(generate())``, so the streamed bytes are
    identical to the monolithic string; streaming avoids holding the whole
    document (plus its UTF-8 encode copy) in memory on million-finding runs.
    """
    private_directory(path.parent)
    template = _ENV.get_template("report.html.j2")
    with path.open("w", encoding="utf-8") as handle:
        for chunk in template.generate(**_render_context(result)):
            handle.write(chunk)
    private_file(path)
    logger.info("HTML report written to %s (%d findings)", path, len(result.findings))
