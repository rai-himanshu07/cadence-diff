"""Machine-readable findings export (JSON) for headless runs and tooling."""

import datetime as dt
import json
from pathlib import Path

from qc_tool.engine import QCRunResult
from qc_tool.security import private_directory, private_file


def result_payload(result: QCRunResult, *, include_context: bool = False) -> dict:
    """A stable, versioned JSON payload for downstream tooling."""
    finding_excludes = (
        set()
        if include_context
        else {"baseline_excerpt", "current_excerpt"}
    )
    return {
        "schema_version": 1,
        "schema": "https://cadence-diff.local/schema/findings-v1.json",
        "generated_at": dt.datetime.now(dt.UTC).isoformat(timespec="seconds"),
        "context_included": include_context,
        "mode": result.mode.value,
        "profile": result.profile_name,
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


def write_json_report(
    result: QCRunResult, path: Path, *, include_context: bool = False
) -> None:
    private_directory(path.parent)
    path.write_text(
        json.dumps(result_payload(result, include_context=include_context), indent=2),
        encoding="utf-8",
    )
    private_file(path)
