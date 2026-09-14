"""Immutable run finalization and signed review evidence."""

from __future__ import annotations

import datetime as dt
import json
import os
import zipfile
from pathlib import Path

import pytest

import qc_tool.signoff as signoff_module
from qc_tool.attestation import (
    create_attestation,
    load_or_create_attestation_key,
    verify_attestation,
)
from qc_tool.config.profile import DeliverableProfile
from qc_tool.coverage import CoverageItem, CoverageState, MappingCoverage, QCRunMode
from qc_tool.engine import QCRunResult
from qc_tool.history.review_state import RunFinalizedError
from qc_tool.history.store import RunHistory, RunRecord
from qc_tool.run_service import perform_run
from qc_tool.signoff import (
    SignoffNotReadyError,
    finalize_run,
    required_acknowledgements,
    review_state_digest,
)


def test_review_state_digest_binds_the_resolved_input_digest() -> None:
    """plan-20260913 Step 4: two otherwise-identical runs (same run_id,
    profile, hashes, decisions) whose resolved logical configuration
    differs must sign a different review-state digest -- the digest is a
    real binding, not a cosmetic disclosure.
    """

    def _record(resolved_input_digest: str) -> RunRecord:
        return RunRecord(
            run_id=1,
            started_at=dt.datetime.now(dt.UTC),
            profile="fixture",
            mode=QCRunMode.CYCLE_COMPARISON,
            files={},
            file_hashes={},
            counts={},
            review_counts={},
            disclosures=[],
            verified_crosschecks=0,
            report_paths={},
            findings=[],
            resolved_input_digest=resolved_input_digest,
        )

    first = _record("a" * 64)
    second = _record("b" * 64)

    assert review_state_digest(first, ()) != review_state_digest(second, ())


def _reviewed_run(fixture_dir: Path, work_dir: Path) -> tuple[int, dict[str, Path]]:
    profile = DeliverableProfile(name="default")
    files = {"current_excel": fixture_dir / "current.xlsx"}
    artifacts = perform_run(
        work_dir,
        files,
        {},
        profile,
        mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
    )
    history = RunHistory(work_dir / "history.sqlite3")
    record = history.get_run(artifacts.run_id)
    updates = [
        (
            finding.finding_id,
            finding.severity.value if finding.severity is not None else None,
            "reviewed for sign-off",
        )
        for finding in record.findings
        if finding.severity is not None
        and finding.severity.value in {"critical", "warning"}
    ]
    history.set_annotations_bulk(artifacts.run_id, updates)
    return artifacts.run_id, files


def _finalize_ready(fixture_dir: Path, work_dir: Path):
    run_id, _ = _reviewed_run(fixture_dir, work_dir)
    record = RunHistory(work_dir / "history.sqlite3").get_run(run_id)
    signoff = finalize_run(
        work_dir,
        run_id,
        set(required_acknowledgements(record)),
    )
    return run_id, signoff


def test_not_included_coverage_does_not_require_signoff_acknowledgement(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    run_id, _files = _reviewed_run(fixture_dir, tmp_path / "work")
    record = RunHistory(tmp_path / "work" / "history.sqlite3").get_run(run_id)
    record.coverage.extend(
        [
            CoverageItem(
                check_id="omitted-ppt",
                label="PowerPoint comparison",
                artifact="ppt",
                state=CoverageState.NOT_INCLUDED,
            ),
            CoverageItem(
                check_id="failed-formulas",
                label="Formula comparison",
                artifact="excel",
                state=CoverageState.UNAVAILABLE,
            ),
        ]
    )

    acknowledgements = set(required_acknowledgements(record))
    assert "coverage:omitted-ppt" not in acknowledgements
    assert "coverage:failed-formulas" in acknowledgements


def test_finalization_stops_this_runs_review_timer(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    work_dir = tmp_path / "work"
    run_id, _files = _reviewed_run(fixture_dir, work_dir)
    history = RunHistory(work_dir / "history.sqlite3")
    history.start_review_session(run_id)
    record = history.get_run(run_id)

    finalize_run(work_dir, run_id, set(required_acknowledgements(record)))

    # the recorded review time is frozen with the sign-off
    assert history.active_review_run() is None
    frozen = history.review_seconds(run_id)
    assert frozen is not None
    later = dt.datetime.now(dt.UTC) + dt.timedelta(hours=1)
    assert history.review_seconds(run_id, now=later) == frozen


def test_finalization_creates_v2_attestation_and_locks_review_mutations(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    work_dir = tmp_path / "work"
    run_id, signoff = _finalize_ready(fixture_dir, work_dir)
    history = RunHistory(work_dir / "history.sqlite3")
    record = history.get_run(run_id)

    assert record.signoff == signoff
    assert all(Path(path).is_file() for path in signoff.report_paths.values())
    assert Path(signoff.attestation_path).is_file()
    _, key = load_or_create_attestation_key(work_dir)
    assert verify_attestation(Path(signoff.attestation_path), key=key).valid
    with zipfile.ZipFile(signoff.attestation_path) as archive:
        manifest = json.loads(archive.read("manifest.json"))
    assert manifest["schema_version"] == 2
    assert manifest["profile_sha256"] == record.profile_sha256
    assert manifest["signoff"]["review_state_digest"] == signoff.review_state_digest
    assert manifest["run"]["formula_engines"] == record.formula_engines
    assert manifest["run"]["values_engines"] == record.values_engines

    finding_id = record.findings[0].finding_id
    with pytest.raises(RunFinalizedError):
        history.set_annotation(run_id, finding_id, severity="info", comment="late")
    with pytest.raises(RunFinalizedError):
        history.set_annotations_bulk(run_id, [(finding_id, "info", "late")])
    with pytest.raises(RunFinalizedError):
        history.update_mapping_review(
            run_id,
            coverage=MappingCoverage(),
            suggestions=[],
            check_coverage=record.coverage,
        )


def test_v2_verifier_keeps_accepting_v1_bundles(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    work_dir = tmp_path / "work"
    profile = DeliverableProfile(name="default")
    files = {"current_excel": fixture_dir / "current.xlsx"}
    artifacts = perform_run(
        work_dir,
        files,
        {},
        profile,
        mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
    )
    _, key = load_or_create_attestation_key(work_dir)
    bundle = create_attestation(
        tmp_path / "v1.qca",
        result=artifacts.result,
        profile=profile,
        input_files=files,
        report_paths=artifacts.report_paths,
        key=key,
    )

    assert verify_attestation(bundle, key=key).valid
    with zipfile.ZipFile(bundle) as archive:
        assert json.loads(archive.read("manifest.json"))["schema_version"] == 1


def test_legacy_run_without_profile_snapshot_cannot_be_finalized(
    tmp_path: Path,
) -> None:
    work_dir = tmp_path / "work"
    history = RunHistory(work_dir / "history.sqlite3")
    run_id = history.record_run(
        QCRunResult(profile_name="default"),
        file_hashes={},
        report_paths={},
    )

    with pytest.raises(SignoffNotReadyError, match="no exact profile snapshot"):
        finalize_run(work_dir, run_id, set())


def test_finalization_succeeds_and_discloses_drift_when_the_named_profile_changed(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    """plan-20260913 Step 12 fix: a run finalizes from its own frozen
    profile_snapshot -- later editing the named profile on disk is
    disclosed on the signoff, never a reason to block finalization or
    invalidate this run's historical evidence.
    """
    from qc_tool.config.profile import profile_path, save_profile

    work_dir = tmp_path / "work"
    profiles_dir = work_dir / "profiles"
    original = DeliverableProfile(name="acme")
    save_profile(original, profile_path(profiles_dir, "acme"))
    files = {"current_excel": fixture_dir / "current.xlsx"}
    artifacts = perform_run(
        work_dir,
        files,
        {},
        original,
        mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
    )
    history = RunHistory(work_dir / "history.sqlite3")
    record = history.get_run(artifacts.run_id)
    updates = [
        (
            finding.finding_id,
            finding.severity.value if finding.severity is not None else None,
            "reviewed for sign-off",
        )
        for finding in record.findings
        if finding.severity is not None
        and finding.severity.value in {"critical", "warning"}
    ]
    history.set_annotations_bulk(artifacts.run_id, updates)

    # The profile changes AFTER the run but BEFORE finalization.
    changed = original.model_copy(
        update={"tolerance": original.tolerance.model_copy(update={"absolute": 5.0})}
    )
    save_profile(changed, profile_path(profiles_dir, "acme"))

    record = history.get_run(artifacts.run_id)
    signoff = finalize_run(
        work_dir, artifacts.run_id, set(required_acknowledgements(record))
    )

    assert signoff.profile_drifted is True
    # The signed attestation still binds the run's OWN frozen profile hash,
    # never the drifted current one.
    assert signoff.profile_sha256 == record.profile_sha256
    with zipfile.ZipFile(signoff.attestation_path) as archive:
        manifest = json.loads(archive.read("manifest.json"))
    assert manifest["profile_sha256"] == record.profile_sha256
    assert manifest["signoff"]["profile_drifted"] is True
    _, key = load_or_create_attestation_key(work_dir)
    assert verify_attestation(Path(signoff.attestation_path), key=key).valid

    # Reopening the finalized run reads the disclosure back correctly.
    stored = history.get_signoff(artifacts.run_id)
    assert stored is not None
    assert stored.profile_drifted is True


def test_finalization_reports_no_drift_when_the_profile_is_unchanged(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    run_id, signoff = _finalize_ready(fixture_dir, tmp_path / "work")
    assert signoff.profile_drifted is False
    history = RunHistory((tmp_path / "work") / "history.sqlite3")
    stored = history.get_signoff(run_id)
    assert stored is not None
    assert stored.profile_drifted is False


@pytest.mark.parametrize("failure", ["attestation", "replace", "database"])
def test_finalization_failures_leave_original_reports_and_run_mutable(
    failure: str,
    fixture_dir: Path,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    work_dir = tmp_path / failure
    run_id, _ = _reviewed_run(fixture_dir, work_dir)
    history = RunHistory(work_dir / "history.sqlite3")
    record = history.get_run(run_id)
    original_reports = [Path(path) for path in record.report_paths.values()]

    if failure == "attestation":
        monkeypatch.setattr(
            signoff_module,
            "create_attestation",
            lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("create failed")),
        )
    elif failure == "replace":
        real_replace = os.replace
        calls = {"count": 0}

        def fail_second_replace(source: Path, destination: Path) -> None:
            calls["count"] += 1
            if calls["count"] == 2:
                raise OSError("replace failed")
            real_replace(source, destination)

        monkeypatch.setattr(signoff_module.os, "replace", fail_second_replace)
    else:
        monkeypatch.setattr(
            RunHistory,
            "record_signoff",
            lambda self, signoff: (_ for _ in ()).throw(RuntimeError("db failed")),
        )

    with pytest.raises((OSError, RuntimeError)):
        finalize_run(
            work_dir,
            run_id,
            set(required_acknowledgements(record)),
        )

    assert RunHistory(work_dir / "history.sqlite3").get_signoff(run_id) is None
    assert all(path.is_file() for path in original_reports)
    RunHistory(work_dir / "history.sqlite3").set_annotation(
        run_id,
        record.findings[0].finding_id,
        severity="info",
        comment="still mutable",
    )
