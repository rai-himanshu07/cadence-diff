"""Explicit evidence-gated annotation carry-forward."""

from pathlib import Path

import pytest

from qc_tool.engine import QCRunResult
from qc_tool.findings import Finding, FindingClass, Severity
from qc_tool.history.carry_forward import apply_carry_forward, preview_carry_forward
from qc_tool.history.store import RunHistory


def _finding(
    finding_id: str,
    location: str,
    value: str,
    *,
    message: str = "changed",
) -> Finding:
    return Finding(
        finding_id=finding_id,
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        severity=Severity.CRITICAL,
        sheet="Data",
        location=location,
        baseline_value="0",
        current_value=value,
        message=message,
    )


def test_preview_and_apply_only_exact_selected_decisions(tmp_path: Path) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    previous = QCRunResult(
        profile_name="fixture",
        findings=[
            _finding("F1", "A1", "1"),
            _finding("F2", "A2", "2"),
            _finding("F3", "A3", "3"),
            _finding("F4", "A4", "4"),
        ],
    )
    previous_id = history.record_run(previous, file_hashes={}, report_paths={})
    history.set_annotations_bulk(
        previous_id,
        [
            ("F1", "info", "carry exact"),
            ("F2", "warning", "evidence changed"),
            ("F3", None, "resolved"),
            ("F4", "info", "ambiguous"),
        ],
    )
    current = QCRunResult(
        profile_name="fixture",
        findings=[
            _finding("N1", "A1", "1", message="reworded only"),
            _finding("N2", "A2", "99"),
            _finding("N4A", "A4", "4"),
            _finding("N4B", "A4", "4"),
        ],
    )
    current_id = history.record_run(
        current,
        file_hashes={},
        report_paths={},
        rerun_of=previous_id,
    )

    preview = preview_carry_forward(history, current_id)

    assert [candidate.finding_id for candidate in preview.exact] == ["N1"]
    assert preview.changed_evidence == ("N2",)
    assert preview.resolved == ("F3",)
    assert preview.ambiguous == ("F4",)
    assert not preview.source_finalized

    assert apply_carry_forward(history, current_id, {"N1"}) == 1
    annotated = history.get_run(current_id)
    exact = next(finding for finding in annotated.findings if finding.finding_id == "N1")
    changed = next(finding for finding in annotated.findings if finding.finding_id == "N2")
    assert exact.severity is Severity.INFO
    assert exact.analyst_comment == "carry exact"
    assert not changed.severity_overridden and not changed.analyst_comment
    lineage = history.get_annotation_lineage(current_id)["N1"]
    assert lineage.source_run_id == previous_id
    assert lineage.source_finding_id == "F1"
    assert len(lineage.evidence_digest) == 64


def test_apply_recomputes_and_rejects_nonexact_or_unknown_selection(
    tmp_path: Path,
) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    previous = QCRunResult(
        profile_name="fixture",
        findings=[_finding("F1", "A1", "1")],
    )
    previous_id = history.record_run(previous, file_hashes={}, report_paths={})
    history.set_annotation(previous_id, "F1", severity="info", comment="reviewed")
    current = QCRunResult(
        profile_name="fixture",
        findings=[_finding("N1", "A1", "2")],
    )
    current_id = history.record_run(
        current,
        file_hashes={},
        report_paths={},
        rerun_of=previous_id,
    )

    with pytest.raises(ValueError, match="not exact candidates"):
        apply_carry_forward(history, current_id, {"N1"})
    assert history.get_annotations(current_id) == {}
    assert history.get_annotation_lineage(current_id) == {}
