"""Machine-readable findings export (JSON) for headless runs and tooling."""

import datetime as dt
import json
from pathlib import Path

from qc_tool.coverage import capability_limited
from qc_tool.engine import QCRunResult
from qc_tool.review import build_pattern_groups, count_pattern_groups
from qc_tool.security import private_directory, private_file
from qc_tool.story import build_stories


def review_summary(result: QCRunResult) -> dict:
    """Versioned, additive semantic review summary; findings stay atomic."""
    groups = build_pattern_groups(result.findings)
    counts = count_pattern_groups(groups)
    stories = build_stories(result.findings)
    return {
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
        "groups": [
            {
                "group_id": group.group_id,
                "finding_class": group.finding_class.value,
                "severity": group.severity.value,
                "artifact": group.artifact,
                "sheet": group.sheet,
                "slide": group.slide,
                "element": group.element,
                "expected_growth": group.expected_growth,
                "expected_reason": (
                    group.members[0].expected_reason.value
                    if group.members[0].expected_reason is not None
                    else None
                ),
                "provenance": (
                    group.members[0].provenance.value
                    if group.members[0].provenance is not None
                    else None
                ),
                "subtype": (
                    group.members[0].subtype.value
                    if group.members[0].subtype is not None
                    else None
                ),
                "materiality": (
                    group.members[0].materiality.value
                    if group.members[0].materiality is not None
                    else None
                ),
                "temporal_context": (
                    group.members[0].temporal_context.value
                    if group.members[0].temporal_context is not None
                    else None
                ),
                "evidence_tags": sorted(
                    tag.value for tag in group.members[0].evidence_tags
                ),
                "ranges": list(group.ranges),
                "bounding_range": group.bounding_range,
                "member_count": group.member_count,
                "finding_ids": [member.finding_id for member in group.members],
            }
            for group in groups
        ],
    }


def result_payload(
    result: QCRunResult,
    *,
    include_context: bool = False,
    include_review_summary: bool = False,
) -> dict:
    """A stable, versioned JSON payload for downstream tooling."""
    finding_excludes = (
        set()
        if include_context
        else {"baseline_excerpt", "current_excerpt"}
    )
    payload = {
        "schema_version": 1,
        "schema": "https://cadence-diff.local/schema/findings-v1.json",
        "generated_at": dt.datetime.now(dt.UTC).isoformat(timespec="seconds"),
        "context_included": include_context,
        "mode": result.mode.value,
        "profile": result.profile_name,
        "comparison_scope": result.comparison_scope.model_dump(mode="json"),
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
        "findings": [
            finding.model_dump(mode="json", exclude=finding_excludes)
            for finding in result.findings
        ],
    }
    if include_review_summary:
        payload["review_summary"] = review_summary(result)
    return payload


def write_json_report(
    result: QCRunResult,
    path: Path,
    *,
    include_context: bool = False,
    include_review_summary: bool = False,
) -> None:
    private_directory(path.parent)
    path.write_text(
        json.dumps(
            result_payload(
                result,
                include_context=include_context,
                include_review_summary=include_review_summary,
            ),
            indent=2,
        ),
        encoding="utf-8",
    )
    private_file(path)
