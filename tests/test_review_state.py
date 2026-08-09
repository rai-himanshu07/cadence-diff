"""Review evidence fingerprints and exact run profile snapshots."""

from pathlib import Path

from qc_tool.config.profile import DeliverableProfile
from qc_tool.coverage import QCRunMode
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingEvidenceTag,
    FindingExpectedReason,
    SeriesAnchorV1,
    Severity,
)
from qc_tool.history.review_state import (
    FINDING_EVIDENCE_FIELDS,
    EvidenceFieldRole,
    finding_evidence_digest,
)
from qc_tool.history.store import RunHistory
from qc_tool.run_service import perform_run


def _finding() -> Finding:
    return Finding(
        finding_id="F1",
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        severity=Severity.CRITICAL,
        sheet="Data",
        location="B2",
        baseline_location="B1",
        slide_index=2,
        baseline_slide_index=1,
        baseline_value="10",
        current_value="11",
        message="value changed",
        impacts=["Data!C2", "Data!D2"],
        evidence_tags={
            FindingEvidenceTag.FORMULA_TEXT,
            FindingEvidenceTag.ADDED_REFERENCE,
        },
        waiver_reason="approved",
        waiver_expires="2026-12-31",
    )


def test_every_finding_field_has_an_evidence_role() -> None:
    assert set(FINDING_EVIDENCE_FIELDS) == set(Finding.model_fields)


def test_series_anchor_is_private_and_never_enters_evidence_or_public_json() -> None:
    finding = _finding()
    digest = finding_evidence_digest(finding)
    payload = finding.model_dump(mode="json")

    anchored = finding.model_copy(
        update={
            "series_anchor": SeriesAnchorV1(
                sheet="Data",
                current_region_id="Data!A1:D9",
                period_axis="rows",
                series_index=2,
                period_index=2,
            )
        }
    )

    assert FINDING_EVIDENCE_FIELDS["series_anchor"] is EvidenceFieldRole.EXCLUDED
    assert finding_evidence_digest(anchored) == digest
    assert anchored.model_dump(mode="json") == payload
    assert "series_anchor" not in payload


def test_evidence_digest_ignores_prose_ids_analyst_state_and_collection_order() -> None:
    finding = _finding()
    equivalent = finding.model_copy(
        update={
            "finding_id": "F999",
            "message": "reworded",
            "analyst_comment": "reviewed",
            "severity_overridden": True,
            "impacts": list(reversed(finding.impacts)),
            "evidence_tags": set(finding.evidence_tags),
        }
    )

    assert finding_evidence_digest(equivalent) == finding_evidence_digest(finding)


def test_evidence_digest_changes_for_locations_slides_values_and_waiver_boundary() -> None:
    finding = _finding()
    digest = finding_evidence_digest(finding)

    for update in (
        {"baseline_location": "B3"},
        {"slide_index": 3},
        {"baseline_slide_index": 4},
        {"current_value": "12"},
        {"waiver_expires": "2027-01-31"},
    ):
        assert finding_evidence_digest(finding.model_copy(update=update)) != digest


def test_effective_expected_distinguishes_reason_and_legacy_fallback() -> None:
    finding = _finding().model_copy(update={"expected_growth": False})
    legacy = finding.model_copy(update={"expected_growth": True})
    typed = finding.model_copy(
        update={"expected_reason": FindingExpectedReason.CADENCE_EXTENSION}
    )

    assert finding_evidence_digest(legacy) != finding_evidence_digest(finding)
    assert finding_evidence_digest(typed) != finding_evidence_digest(legacy)


def test_perform_run_records_exact_profile_snapshot(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    profile = DeliverableProfile(
        name="snapshot",
        severity={FindingClass.VALUE_CHANGED: Severity.WARNING},
    )
    artifacts = perform_run(
        tmp_path / "work",
        {"current_excel": fixture_dir / "current.xlsx"},
        {},
        profile,
        mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
    )

    record = RunHistory(tmp_path / "work" / "history.sqlite3").get_run(
        artifacts.run_id
    )
    assert record.profile_snapshot == profile
    assert len(record.profile_sha256) == 64
