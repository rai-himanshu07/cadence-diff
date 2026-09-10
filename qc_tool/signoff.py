"""Immutable review finalization and signed evidence publication."""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
from pathlib import Path

from pydantic import BaseModel

from qc_tool.attestation import (
    AttestationSignoff,
    create_attestation,
    load_or_create_attestation_key,
    verify_attestation,
)
from qc_tool.config.profile import (
    default_profile,
    load_profile,
    profile_path,
    profile_sha256,
)
from qc_tool.coverage import CoverageState
from qc_tool.engine import QCRunResult
from qc_tool.findings import Severity
from qc_tool.history.review_state import RunSignoff
from qc_tool.history.store import RunHistory, RunRecord, sha256_file
from qc_tool.report.excel_report import write_excel_report
from qc_tool.report.html_report import write_html_report
from qc_tool.security import private_directory, private_file


class SignoffNotReadyError(RuntimeError):
    """Raised when a run cannot be finalized without unresolved review work."""


class SignoffAssessment(BaseModel):
    undecided_finding_ids: tuple[str, ...] = ()
    required_acknowledgements: tuple[str, ...] = ()
    missing_acknowledgements: tuple[str, ...] = ()

    @property
    def ready(self) -> bool:
        return not self.undecided_finding_ids and not self.missing_acknowledgements


def required_acknowledgements(record: RunRecord) -> tuple[str, ...]:
    required = {
        f"coverage:{item.check_id}"
        for item in record.coverage
        if item.state is not CoverageState.CHECKED
    }
    mapping = record.mapping_coverage
    if mapping is not None:
        for field_name in ("mismatched", "unresolved", "unmapped", "unavailable"):
            if getattr(mapping, field_name):
                required.add(f"mapping:{field_name}")
    return tuple(sorted(required))


def assess_signoff(
    record: RunRecord,
    acknowledgements: set[str] | tuple[str, ...] = (),
) -> SignoffAssessment:
    decided = {
        finding.finding_id
        for finding in record.findings
        if finding.severity_overridden
        or finding.analyst_comment
        or finding.waiver_reason
    }
    undecided = tuple(
        finding.finding_id
        for finding in record.findings
        if finding.severity in {Severity.CRITICAL, Severity.WARNING}
        and finding.finding_id not in decided
    )
    required = required_acknowledgements(record)
    accepted = set(acknowledgements)
    missing = tuple(code for code in required if code not in accepted)
    return SignoffAssessment(
        undecided_finding_ids=undecided,
        required_acknowledgements=required,
        missing_acknowledgements=missing,
    )


def review_state_digest(
    record: RunRecord,
    acknowledgements: tuple[str, ...],
) -> str:
    decisions = [
        {
            "finding_id": finding.finding_id,
            "severity": finding.severity.value if finding.severity else None,
            "severity_overridden": finding.severity_overridden,
            "comment": finding.analyst_comment,
            "waiver_reason": finding.waiver_reason,
            "waiver_expires": finding.waiver_expires,
        }
        for finding in record.findings
        if finding.severity_overridden
        or finding.analyst_comment
        or finding.waiver_reason
    ]
    payload = {
        "run_id": record.run_id,
        "profile_sha256": record.profile_sha256,
        "file_hashes": dict(sorted(record.file_hashes.items())),
        "acknowledgements": sorted(acknowledgements),
        "decisions": sorted(decisions, key=lambda item: str(item["finding_id"])),
        "counterfactual_digest": getattr(record, "counterfactual_digest", ""),
    }
    if (
        record.package_manifest is not None
        and not record.package_manifest.is_legacy_projection
    ):
        payload["package_manifest"] = record.package_manifest.model_dump(mode="json")
    encoded = json.dumps(payload, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(encoded).hexdigest()


def _result_from_record(record: RunRecord) -> QCRunResult:
    return QCRunResult(
        profile_name=record.profile,
        mode=record.mode,
        requested_output_mode=record.requested_output_mode,
        resolved_output_policy=record.resolved_output_policy,
        files=record.files,
        findings=record.findings,
        disclosures=record.disclosures,
        coverage=record.coverage,
        mapping_coverage=record.mapping_coverage,
        mapping_suggestions=record.mapping_suggestions,
        verified_crosschecks=record.verified_crosschecks,
        comparison_scope=record.comparison_scope,
        package_manifest=record.package_manifest,
        alignment_trust=record.alignment_trust,
    )


def _current_profile(work_dir: Path, record: RunRecord):
    if record.profile == "default":
        return default_profile()
    return load_profile(profile_path(work_dir / "profiles", record.profile))


def finalize_run(
    work_dir: Path,
    run_id: int,
    acknowledgements: set[str] | tuple[str, ...],
) -> RunSignoff:
    """Publish reviewed reports and an attestation, then lock the run."""
    history = RunHistory(work_dir / "history.sqlite3")
    history.assert_mutable(run_id)
    record = history.get_run(run_id)
    accepted = tuple(sorted(set(acknowledgements)))
    assessment = assess_signoff(record, accepted)
    if not assessment.ready:
        raise SignoffNotReadyError(
            "review is incomplete: "
            f"{len(assessment.undecided_finding_ids)} undecided findings; "
            f"{len(assessment.missing_acknowledgements)} missing acknowledgements"
        )
    if record.profile_snapshot is None or not record.profile_sha256:
        raise SignoffNotReadyError(
            "this legacy run has no exact profile snapshot; submit a Re-QC run"
        )
    current_profile = _current_profile(work_dir, record)
    if profile_sha256(current_profile) != record.profile_sha256:
        raise SignoffNotReadyError(
            "the profile changed after this run; submit a Re-QC run before sign-off"
        )
    if set(record.file_paths) != set(record.file_hashes):
        raise SignoffNotReadyError("the run does not retain every source path")
    input_files = {role: Path(path) for role, path in record.file_paths.items()}
    for role, path in input_files.items():
        if not path.is_file() or sha256_file(path) != record.file_hashes[role]:
            raise SignoffNotReadyError(
                f"the saved bytes for role {role!r} no longer match this run"
            )

    finalized_at = dt.datetime.now(dt.UTC).isoformat(timespec="seconds")
    digest = review_state_digest(record, accepted)
    lineages = [
        lineage.model_dump(mode="json")
        for lineages in history.get_annotation_lineage(run_id).values()
        for lineage in lineages
    ]
    signoff_evidence = AttestationSignoff(
        finalized_at=finalized_at,
        acknowledgements=accepted,
        review_state_digest=digest,
        annotation_lineage=lineages,
    )
    result = _result_from_record(record)
    signoff_dir = private_directory(work_dir / "runs" / f"signoff-{run_id}")
    final_paths = {
        "excel": signoff_dir / "qc_report.final.xlsx",
        "html": signoff_dir / "qc_report.final.html",
    }
    attestation_path = signoff_dir / "run.final.qca"
    temporary_paths = {
        "excel": signoff_dir / ".qc_report.final.tmp.xlsx",
        "html": signoff_dir / ".qc_report.final.tmp.html",
        "attestation": signoff_dir / ".run.final.tmp.qca",
    }
    published: list[Path] = []
    try:
        write_excel_report(result, temporary_paths["excel"])
        write_html_report(result, temporary_paths["html"])
        _, key = load_or_create_attestation_key(work_dir)
        create_attestation(
            temporary_paths["attestation"],
            result=result,
            profile=record.profile_snapshot,
            input_files=input_files,
            report_paths={
                "excel": temporary_paths["excel"],
                "html": temporary_paths["html"],
            },
            key=key,
            signoff=signoff_evidence,
        )
        verification = verify_attestation(temporary_paths["attestation"], key=key)
        if not verification.valid:
            raise RuntimeError("new sign-off attestation did not verify")
        for source, destination in (
            (temporary_paths["excel"], final_paths["excel"]),
            (temporary_paths["html"], final_paths["html"]),
            (temporary_paths["attestation"], attestation_path),
        ):
            os.replace(source, destination)
            private_file(destination)
            published.append(destination)
        signoff = RunSignoff(
            run_id=run_id,
            finalized_at=finalized_at,
            acknowledgements=accepted,
            review_state_digest=digest,
            profile_sha256=record.profile_sha256,
            attestation_path=str(attestation_path),
            attestation_sha256=sha256_file(attestation_path),
            report_paths={kind: str(path) for kind, path in final_paths.items()},
        )
        history.record_signoff(signoff)
        # a finalized review is over; the recorded time must stop with it
        history.pause_review_sessions(run_id=run_id)
        return signoff
    except BaseException:
        for path in (*temporary_paths.values(), *published):
            path.unlink(missing_ok=True)
        raise
