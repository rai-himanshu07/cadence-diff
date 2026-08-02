"""Standalone HTML report rendered from the triaged `QCRunResult`.

Self-contained (inline CSS, vanilla JS, no external assets) and
autoescaped — finding text originates from client files and must never
inject markup.
"""

import datetime as dt
import logging
from pathlib import Path
from typing import TypedDict

from jinja2 import Environment, FileSystemLoader, select_autoescape

from qc_tool.coverage import capability_limited
from qc_tool.engine import QCRunResult
from qc_tool.review import build_pattern_groups, count_pattern_groups
from qc_tool.security import private_directory, private_file
from qc_tool.story import build_stories

logger = logging.getLogger(__name__)


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


def _member_payload(result: QCRunResult) -> dict[str, list[_HtmlMember]]:
    payload: dict[str, list[_HtmlMember]] = {}
    for group in build_pattern_groups(result.findings):
        payload[group.group_id] = [
            {
                "finding_id": finding.finding_id,
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
            }
            for finding in group.members
        ]
    return payload

_ENV = Environment(
    loader=FileSystemLoader(Path(__file__).parent / "templates"),
    autoescape=select_autoescape(enabled_extensions=("j2", "html")),
)


def render_html_report(result: QCRunResult) -> str:
    template = _ENV.get_template("report.html.j2")
    groups = build_pattern_groups(result.findings)
    return template.render(
        result=result,
        groups=groups,
        stories=build_stories(result.findings),
        member_payload=_member_payload(result),
        counts=count_pattern_groups(groups),
        capability_limited=capability_limited(result.coverage),
        generated_at=dt.datetime.now(dt.UTC).isoformat(timespec="seconds"),
    )


def write_html_report(result: QCRunResult, path: Path) -> None:
    private_directory(path.parent)
    path.write_text(render_html_report(result), encoding="utf-8")
    private_file(path)
    logger.info("HTML report written to %s (%d findings)", path, len(result.findings))
