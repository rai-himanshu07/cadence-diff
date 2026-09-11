"""Explicit evidence-gated annotation carry-forward."""

from pathlib import Path

import pytest

from qc_tool.engine import QCRunResult, compare_findings
from qc_tool.findings import (
    Finding,
    FindingClass,
    MembershipCodec,
    PopulationEvidence,
    Severity,
)
from qc_tool.history.carry_forward import (
    apply_carry_forward,
    apply_population_carry_forward,
    preview_carry_forward,
)
from qc_tool.history.review_state import AnnotationLineageOutcome, RunSignoff
from qc_tool.history.store import RunHistory


def _finding(
    finding_id: str,
    location: str,
    value: str,
    *,
    message: str = "changed",
    finding_class: FindingClass = FindingClass.VALUE_CHANGED,
) -> Finding:
    return Finding(
        finding_id=finding_id,
        artifact="excel",
        finding_class=finding_class,
        severity=Severity.CRITICAL,
        sheet="Data",
        location=location,
        baseline_value="0",
        current_value=value,
        message=message,
    )


def _population(
    finding_id: str,
    pairs: tuple[tuple[str, str], ...],
    *,
    shape_before: str = "a" * 64,
    shape_after: str = "b" * 64,
) -> Finding:
    return Finding(
        finding_id=finding_id,
        artifact="excel",
        finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
        severity=Severity.WARNING,
        sheet="Data",
        location=f"{pairs[0][0]}:{pairs[-1][0]}",
        element="population",
        message=f"{len(pairs)} cells share one population",
        population=PopulationEvidence(
            member_count=len(pairs),
            membership=MembershipCodec(
                baseline_mode="pairs",
                pairs=pairs,
                member_count=len(pairs),
            ),
            first=pairs[0][0],
            last=pairs[-1][0],
            shape_before_digest=shape_before,
            shape_after_digest=shape_after,
        ),
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
    assert len(lineage) == 1
    assert lineage[0].source_run_id == previous_id
    assert lineage[0].source_finding_id == "F1"
    assert lineage[0].relation.value == "identity"
    assert lineage[0].outcome.value == "inherited"
    assert len(lineage[0].source_digest) == 64


def test_preview_discloses_a_cross_run_formula_engine_mismatch(tmp_path: Path) -> None:
    """Criterion 5: a shared excel role whose resolved formula engine
    differs between the source and Re-QC run must be disclosed -- the
    `exact`/`changed_evidence` digest check itself stays conservative either
    way, but the analyst should know evidence may not be directly
    comparable rather than silently assuming nothing changed.
    """
    history = RunHistory(tmp_path / "history.sqlite3")
    previous = QCRunResult(
        profile_name="fixture",
        findings=[_finding("F1", "A1", "1")],
        formula_engines={"current_excel": "excel:16.0.0000"},
    )
    previous_id = history.record_run(previous, file_hashes={}, report_paths={})
    current = QCRunResult(
        profile_name="fixture",
        findings=[_finding("N1", "A1", "1")],
        formula_engines={"current_excel": "native-biff12:1.2.3"},
    )
    current_id = history.record_run(
        current, file_hashes={}, report_paths={}, rerun_of=previous_id
    )

    preview = preview_carry_forward(history, current_id)

    assert preview.engine_provenance_mismatch == ("current_excel",)


def test_preview_reports_no_mismatch_when_engines_match_or_are_unrecorded(
    tmp_path: Path,
) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    previous = QCRunResult(
        profile_name="fixture",
        findings=[_finding("F1", "A1", "1")],
        formula_engines={"current_excel": "native-biff12:1.2.3"},
    )
    previous_id = history.record_run(previous, file_hashes={}, report_paths={})
    current = QCRunResult(
        profile_name="fixture",
        findings=[_finding("N1", "A1", "1")],
        formula_engines={"current_excel": "native-biff12:1.2.3"},
    )
    current_id = history.record_run(
        current, file_hashes={}, report_paths={}, rerun_of=previous_id
    )

    preview = preview_carry_forward(history, current_id)

    assert preview.engine_provenance_mismatch == ()


def test_preview_discloses_a_cross_run_values_engine_mismatch(tmp_path: Path) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    previous_id = history.record_run(
        QCRunResult(
            profile_name="fixture",
            findings=[_finding("F1", "A1", "1")],
            values_engines={"current_excel": "pyxlsb:1.0.10"},
        ),
        file_hashes={},
        report_paths={},
    )
    current_id = history.record_run(
        QCRunResult(
            profile_name="fixture",
            findings=[_finding("N1", "A1", "1")],
            values_engines={"current_excel": "native-biff12:1.2.3"},
        ),
        file_hashes={},
        report_paths={},
        rerun_of=previous_id,
    )

    preview = preview_carry_forward(history, current_id)

    assert preview.engine_provenance_mismatch == ("current_excel",)


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


def _finalize(history: RunHistory, run_id: int) -> None:
    history.record_signoff(
        RunSignoff(
            run_id=run_id,
            finalized_at="2026-09-08T00:00:00+00:00",
            review_state_digest="review-state",
            profile_sha256="profile-sha",
            attestation_path="run.qca",
            attestation_sha256="attestation-sha",
        )
    )


def test_atomic_sources_inherit_cleanly_into_a_new_population(tmp_path: Path) -> None:
    """Every member agreeing, full coverage, finalized source -> inherited."""
    history = RunHistory(tmp_path / "history.sqlite3")
    previous = QCRunResult(
        profile_name="fixture",
        findings=[
            _finding("F1", "B2", "x", finding_class=FindingClass.FORMULA_LOGIC_CHANGED),
            _finding("F2", "B3", "x", finding_class=FindingClass.FORMULA_LOGIC_CHANGED),
            _finding("F3", "B4", "x", finding_class=FindingClass.FORMULA_LOGIC_CHANGED),
        ],
    )
    previous_id = history.record_run(previous, file_hashes={}, report_paths={})
    history.set_annotations_bulk(
        previous_id,
        [
            ("F1", "info", "reviewed batch"),
            ("F2", "info", "reviewed batch"),
            ("F3", "info", "reviewed batch"),
        ],
    )
    _finalize(history, previous_id)
    current = QCRunResult(
        profile_name="fixture",
        findings=[_population("P1", (("B2", "B2"), ("B3", "B3"), ("B4", "B4")))],
    )
    current_id = history.record_run(
        current, file_hashes={}, report_paths={}, rerun_of=previous_id
    )

    preview = preview_carry_forward(history, current_id)

    assert len(preview.populations) == 1
    candidate = preview.populations[0]
    assert candidate.finding_id == "P1"
    assert candidate.outcome is AnnotationLineageOutcome.INHERITED
    assert candidate.severity == "info"
    assert candidate.comment == "reviewed batch"
    assert candidate.member_count == 3
    assert candidate.matched_member_count == 3

    assert apply_population_carry_forward(history, current_id, "P1") == 3
    annotated = history.get_run(current_id)
    population_finding = next(
        f for f in annotated.findings if f.finding_id == "P1"
    )
    assert population_finding.severity is Severity.INFO
    assert population_finding.analyst_comment == "reviewed batch"
    lineage = history.get_annotation_lineage(current_id)["P1"]
    assert len(lineage) == 3
    assert {row.source_finding_id for row in lineage} == {"F1", "F2", "F3"}
    assert all(row.relation.value == "member" for row in lineage)
    assert all(row.outcome.value == "inherited" for row in lineage)
    assert history.carried_annotation_count(current_id) == 1


def test_conflicting_severities_across_members_disclosed_not_applied(
    tmp_path: Path,
) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    previous = QCRunResult(
        profile_name="fixture",
        findings=[
            _finding("F1", "B2", "x", finding_class=FindingClass.FORMULA_LOGIC_CHANGED),
            _finding("F2", "B3", "x", finding_class=FindingClass.FORMULA_LOGIC_CHANGED),
        ],
    )
    previous_id = history.record_run(previous, file_hashes={}, report_paths={})
    history.set_annotations_bulk(
        previous_id,
        [("F1", "info", "fine"), ("F2", "critical", "actually a problem")],
    )
    _finalize(history, previous_id)
    current = QCRunResult(
        profile_name="fixture",
        findings=[_population("P1", (("B2", "B2"), ("B3", "B3")))],
    )
    current_id = history.record_run(
        current, file_hashes={}, report_paths={}, rerun_of=previous_id
    )

    preview = preview_carry_forward(history, current_id)

    assert len(preview.populations) == 1
    candidate = preview.populations[0]
    assert candidate.outcome is AnnotationLineageOutcome.CONFLICT_SEVERITY
    assert candidate.severity is None
    assert candidate.matched_member_count == 2
    with pytest.raises(ValueError, match="not an inherited carry-forward candidate"):
        apply_population_carry_forward(history, current_id, "P1")


def test_conflicting_comments_same_severity_disclosed_not_applied(
    tmp_path: Path,
) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    previous = QCRunResult(
        profile_name="fixture",
        findings=[
            _finding("F1", "B2", "x", finding_class=FindingClass.FORMULA_LOGIC_CHANGED),
            _finding("F2", "B3", "x", finding_class=FindingClass.FORMULA_LOGIC_CHANGED),
        ],
    )
    previous_id = history.record_run(previous, file_hashes={}, report_paths={})
    history.set_annotations_bulk(
        previous_id,
        [("F1", "info", "reason A"), ("F2", "info", "reason B")],
    )
    _finalize(history, previous_id)
    current = QCRunResult(
        profile_name="fixture",
        findings=[_population("P1", (("B2", "B2"), ("B3", "B3")))],
    )
    current_id = history.record_run(
        current, file_hashes={}, report_paths={}, rerun_of=previous_id
    )

    preview = preview_carry_forward(history, current_id)

    assert len(preview.populations) == 1
    assert preview.populations[0].outcome is AnnotationLineageOutcome.CONFLICT_COMMENT
    with pytest.raises(ValueError, match="not an inherited carry-forward candidate"):
        apply_population_carry_forward(history, current_id, "P1")


def test_partial_member_coverage_disclosed_not_applied(tmp_path: Path) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    previous = QCRunResult(
        profile_name="fixture",
        findings=[
            _finding("F1", "B2", "x", finding_class=FindingClass.FORMULA_LOGIC_CHANGED),
            _finding("F2", "B3", "x", finding_class=FindingClass.FORMULA_LOGIC_CHANGED),
        ],
    )
    previous_id = history.record_run(previous, file_hashes={}, report_paths={})
    # Only F1 is annotated; F2 was never reviewed in the source run.
    history.set_annotations_bulk(previous_id, [("F1", "info", "fine")])
    _finalize(history, previous_id)
    current = QCRunResult(
        profile_name="fixture",
        findings=[_population("P1", (("B2", "B2"), ("B3", "B3")))],
    )
    current_id = history.record_run(
        current, file_hashes={}, report_paths={}, rerun_of=previous_id
    )

    preview = preview_carry_forward(history, current_id)

    assert len(preview.populations) == 1
    candidate = preview.populations[0]
    assert candidate.outcome is AnnotationLineageOutcome.PARTIAL
    assert candidate.matched_member_count == 1
    assert candidate.member_count == 2
    with pytest.raises(ValueError, match="not an inherited carry-forward candidate"):
        apply_population_carry_forward(history, current_id, "P1")


def test_unfinalized_source_run_disclosed_not_applied(tmp_path: Path) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    previous = QCRunResult(
        profile_name="fixture",
        findings=[
            _finding("F1", "B2", "x", finding_class=FindingClass.FORMULA_LOGIC_CHANGED),
        ],
    )
    previous_id = history.record_run(previous, file_hashes={}, report_paths={})
    history.set_annotations_bulk(previous_id, [("F1", "info", "fine")])
    # Deliberately never finalized (no _finalize(history, previous_id) call).
    current = QCRunResult(
        profile_name="fixture",
        findings=[_population("P1", (("B2", "B2"),))],
    )
    current_id = history.record_run(
        current, file_hashes={}, report_paths={}, rerun_of=previous_id
    )

    preview = preview_carry_forward(history, current_id)

    assert len(preview.populations) == 1
    assert preview.populations[0].outcome is AnnotationLineageOutcome.UNFINALIZED_SOURCE
    with pytest.raises(ValueError, match="not an inherited carry-forward candidate"):
        apply_population_carry_forward(history, current_id, "P1")


def test_population_turned_atomic_is_disclosed_not_silently_resolved(
    tmp_path: Path,
) -> None:
    """Policy turned off later: a population's decision cannot map onto
    plain atomics, so it is disclosed as unresolved -- never silently
    treated as fixed.
    """
    history = RunHistory(tmp_path / "history.sqlite3")
    previous = QCRunResult(
        profile_name="fixture",
        findings=[_population("P1", (("B2", "B2"), ("B3", "B3")))],
    )
    previous_id = history.record_run(previous, file_hashes={}, report_paths={})
    history.set_annotation(previous_id, "P1", severity="info", comment="reviewed")
    current = QCRunResult(
        profile_name="fixture",
        findings=[
            _finding("N1", "B2", "x", finding_class=FindingClass.FORMULA_LOGIC_CHANGED),
            _finding("N2", "B3", "x", finding_class=FindingClass.FORMULA_LOGIC_CHANGED),
        ],
    )
    current_id = history.record_run(
        current, file_hashes={}, report_paths={}, rerun_of=previous_id
    )

    preview = preview_carry_forward(history, current_id)

    assert preview.ambiguous == ("P1",)
    assert preview.resolved == ()
    assert preview.exact == ()


def test_population_identity_survives_member_set_churn(tmp_path: Path) -> None:
    """Criterion 5: a population's identity excludes geometry, so inserting
    a member between runs still pairs it as the same persisting item, not
    a spurious resolved+new pair.
    """
    previous = [_population("P1", (("B2", "B2"), ("B3", "B3")))]
    current = [
        _population("P1", (("B2", "B2"), ("B3", "B3"), ("B4", "B4")))
    ]

    delta = compare_findings(previous, current)

    assert delta.resolved == 0
    assert delta.new == 0
    assert delta.persisting == 1

