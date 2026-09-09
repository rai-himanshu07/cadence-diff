"""Machine-readable findings export (JSON) for headless runs and tooling."""

import datetime as dt
import json
import textwrap
from pathlib import Path

from qc_tool.coverage import capability_limited
from qc_tool.engine import QCRunResult
from qc_tool.findings_store import finding_by_id
from qc_tool.review_stream import (
    counts_from_summaries,
    stream_stories,
    summarize_pattern_groups,
)
from qc_tool.security import private_directory, private_file


def review_summary(
    result: QCRunResult,
    *,
    include_alignment_trust: bool = False,
) -> dict:
    """Versioned, additive semantic review summary; findings stay atomic."""
    summaries = summarize_pattern_groups(result.findings)
    counts = counts_from_summaries(summaries)
    stories = stream_stories(
        [iter(result.findings), iter(result.findings), iter(result.findings)]
    )
    groups_payload = []
    for summary in summaries:
        first = finding_by_id(result.findings, summary.member_finding_ids[0])
        groups_payload.append(
            {
                "group_id": summary.group_id,
                "finding_class": summary.finding_class.value,
                "severity": summary.severity.value,
                "artifact": summary.artifact,
                "sheet": summary.sheet,
                "slide": summary.slide,
                "element": summary.element,
                "expected_growth": summary.expected_growth,
                "expected_reason": (
                    first.expected_reason.value
                    if first is not None and first.expected_reason is not None
                    else None
                ),
                "provenance": (
                    first.provenance.value
                    if first is not None and first.provenance is not None
                    else None
                ),
                "subtype": (
                    first.subtype.value
                    if first is not None and first.subtype is not None
                    else None
                ),
                "materiality": (
                    first.materiality.value
                    if first is not None and first.materiality is not None
                    else None
                ),
                "temporal_context": (
                    first.temporal_context.value
                    if first is not None and first.temporal_context is not None
                    else None
                ),
                "evidence_tags": sorted(
                    tag.value
                    for tag in (first.evidence_tags if first is not None else set())
                ),
                "ranges": list(summary.ranges),
                "bounding_range": summary.bounding_range,
                "member_count": summary.member_count,
                "finding_ids": list(summary.member_finding_ids),
                "population": (
                    {
                        "member_count": first.population.member_count,
                        "shape_before_digest": first.population.shape_before_digest,
                        "shape_after_digest": first.population.shape_after_digest,
                        "baseline_mode": first.population.membership.baseline_mode,
                        "current_rectangles": list(
                            first.population.membership.current_rectangles
                        ),
                        "first": first.population.first,
                        "last": first.population.last,
                    }
                    if first is not None and first.population is not None
                    else None
                ),
            }
        )
    payload = {
        "summary_version": 3,
        "capability_limited": capability_limited(result.coverage),
        "pattern_review_counts": {
            severity.value: count for severity, count in counts.review_items.items()
        },
        "atomic_findings_by_severity": {
            severity.value: count for severity, count in counts.atomic_findings.items()
        },
        "stories": [
            {
                "story_id": story.story_id,
                "kind": story.kind.value,
                "title": story.title,
                "description": story.description,
                "evidence": list(story.evidence),
                "member_count": story.member_count,
                "severity_counts": story.severity_counts,
                "finding_ids": list(story.finding_ids),
            }
            for story in stories
        ],
        "groups": groups_payload,
    }
    if include_alignment_trust:
        payload["summary_version"] = 4
        payload["alignment_trust"] = (
            result.alignment_trust.model_dump(mode="json")
            if result.alignment_trust is not None
            else None
        )
    return payload


def _payload_scaffold(
    result: QCRunResult,
    *,
    include_context: bool = False,
    include_review_summary: bool = False,
) -> dict:
    """Everything but the findings array, with ``findings`` in position."""
    package_manifest = (
        result.package_manifest
        if result.package_manifest is not None
        and not result.package_manifest.is_legacy_projection
        else None
    )
    has_population = any(finding.population is not None for finding in result.findings)
    extended_schema = include_review_summary or package_manifest is not None
    schema_version = 3 if has_population else (2 if extended_schema else 1)
    payload = {
        "schema_version": schema_version,
        "schema": (
            f"https://cadence-diff.local/schema/findings-v{schema_version}.json"
        ),
        "generated_at": dt.datetime.now(dt.UTC).isoformat(timespec="seconds"),
        "context_included": include_context,
        "mode": result.mode.value,
        "profile": result.profile_name,
        "comparison_scope": result.comparison_scope.model_dump(
            mode="json",
            exclude_none=True,
        ),
        "files": result.files,
        "counts": {sev.value: count for sev, count in result.counts.items()},
        "disclosures": result.disclosures,
        "coverage": [item.model_dump(mode="json") for item in result.coverage],
        "mapping_coverage": (
            result.mapping_coverage.model_dump(mode="json")
            if result.mapping_coverage is not None
            else None
        ),
        "verified_crosschecks": result.verified_crosschecks,
        "mapping_suggestions": (
            [item.model_dump(mode="json") for item in result.mapping_suggestions]
            if include_context
            else []
        ),
        "findings": [],
    }
    if package_manifest is not None:
        payload["package_manifest"] = package_manifest.model_dump(mode="json")
    if include_review_summary:
        payload["review_summary"] = review_summary(
            result,
            include_alignment_trust=True,
        )
    return payload


def _finding_excludes(include_context: bool) -> set[str]:
    return set() if include_context else {"baseline_excerpt", "current_excerpt"}


def result_payload(
    result: QCRunResult,
    *,
    include_context: bool = False,
    include_review_summary: bool = False,
) -> dict:
    """A stable, versioned JSON payload for downstream tooling."""
    payload = _payload_scaffold(
        result,
        include_context=include_context,
        include_review_summary=include_review_summary,
    )
    excludes = _finding_excludes(include_context)
    payload["findings"] = [
        finding.model_dump(mode="json", exclude=excludes)
        for finding in result.findings
    ]
    return payload


def write_json_report(
    result: QCRunResult,
    path: Path,
    *,
    include_context: bool = False,
    include_review_summary: bool = False,
) -> None:
    """Stream the payload so the findings array never materializes at once."""
    scaffold = _payload_scaffold(
        result,
        include_context=include_context,
        include_review_summary=include_review_summary,
    )
    marker = "__qc_findings_stream_marker__"
    scaffold["findings"] = marker
    text = json.dumps(scaffold, indent=2)
    head, _, tail = text.partition(f'"{marker}"')
    excludes = _finding_excludes(include_context)
    private_directory(path.parent)
    with path.open("w", encoding="utf-8") as handle:
        handle.write(head)
        handle.write("[")
        first = True
        for finding in result.findings:
            handle.write("\n" if first else ",\n")
            first = False
            item = json.dumps(
                finding.model_dump(mode="json", exclude=excludes), indent=2
            )
            handle.write(textwrap.indent(item, "    "))
        if not first:
            handle.write("\n  ")
        handle.write("]")
        handle.write(tail)
    private_file(path)
