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
from qc_tool.coverage import MappingCoverage, QCRunMode
from qc_tool.engine import QCRunResult
from qc_tool.history.review_state import RunFinalizedError
from qc_tool.history.store import RunHistory
from qc_tool.run_service import perform_run
from qc_tool.signoff import (
    SignoffNotReadyError,
    finalize_run,
    required_acknowledgements,
)


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
