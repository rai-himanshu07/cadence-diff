"""Run history tests (criterion 13)."""

import datetime as dt
import json
import sqlite3
import zipfile
from pathlib import Path

import pytest

from qc_tool.engine import QCRunResult
from qc_tool.history.store import RunHistory, export_runs_archive, sha256_file


def test_record_and_list_runs(
    qc_result: QCRunResult, fixture_dir: Path, tmp_path: Path
) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    hashes = {
        "baseline_excel": sha256_file(fixture_dir / "baseline.xlsx"),
        "current_excel": sha256_file(fixture_dir / "current.xlsx"),
    }
    first = history.record_run(
        qc_result, file_hashes=hashes, report_paths={"html": "runs/1/report.html"}
    )
    second = history.record_run(qc_result, file_hashes=hashes, report_paths={})
    assert second > first

    runs = history.list_runs()
    assert [run.run_id for run in runs] == [second, first]
    newest = runs[0]
    assert newest.profile == "fixture"
    assert newest.mode == qc_result.mode
    assert newest.coverage == qc_result.coverage
    assert newest.files["current_excel"] == "current.xlsx"
    assert newest.file_hashes == hashes
    assert newest.counts["critical"] > 0
    assert newest.review_counts["critical"] > 0
    assert newest.review_counts["critical"] <= newest.counts["critical"]
    assert newest.started_at.tzinfo is not None  # timezone-aware UTC
    assert newest.findings == []  # listing stays lightweight


def test_archive_hides_runs_without_losing_them(
    qc_result: QCRunResult, tmp_path: Path
) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    keep = history.record_run(qc_result, file_hashes={}, report_paths={})
    retire = history.record_run(qc_result, file_hashes={}, report_paths={})

    assert history.set_archived([retire], True) == 1

    active = history.list_runs(include_archived=False)
    assert [run.run_id for run in active] == [keep]
    assert {run.run_id: run.archived for run in history.list_runs()} == {
        keep: False,
        retire: True,
    }
    # The evidence itself survives archiving.
    assert history.get_run(retire).findings

    assert history.set_archived([retire], False) == 1
    assert [run.run_id for run in history.list_runs(include_archived=False)] == [
        retire,
        keep,
    ]


def test_delete_removes_records_annotations_and_managed_reports(
    qc_result: QCRunResult, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    runs_root = work_dir / "runs"
    report = runs_root / "abc" / "report.html"
    report.parent.mkdir(parents=True)
    report.write_text("<html></html>", encoding="utf-8")
    history = RunHistory(work_dir / "history.sqlite3")
    run_id = history.record_run(
        qc_result, file_hashes={}, report_paths={"html": str(report)}
    )
    finding_id = history.get_run(run_id).findings[0].finding_id
    history.set_annotation(run_id, finding_id, severity="info", comment="checked")

    assert history.delete_runs([run_id], managed_root=runs_root) == 1

    assert history.list_runs() == []
    assert history.get_annotations(run_id) == {}
    assert not report.exists()
    assert not report.parent.exists()  # the emptied run directory goes too


def test_delete_never_touches_files_outside_the_managed_root(
    qc_result: QCRunResult, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    runs_root = work_dir / "runs"
    runs_root.mkdir(parents=True)
    outsider = tmp_path / "not-ours.html"
    outsider.write_text("keep me", encoding="utf-8")
    history = RunHistory(work_dir / "history.sqlite3")
    run_id = history.record_run(
        qc_result, file_hashes={}, report_paths={"html": str(outsider)}
    )

    assert history.delete_runs([run_id], managed_root=runs_root) == 1

    assert history.list_runs() == []
    assert outsider.read_text(encoding="utf-8") == "keep me"


def test_export_archive_bundles_reports_and_a_path_free_manifest(
    qc_result: QCRunResult, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    runs_root = work_dir / "runs"
    report = runs_root / "abc" / "report.html"
    report.parent.mkdir(parents=True)
    report.write_text("<html>report</html>", encoding="utf-8")
    outsider = tmp_path / "escape.html"
    outsider.write_text("secret", encoding="utf-8")
    history = RunHistory(work_dir / "history.sqlite3")
    inside = history.record_run(
        qc_result, file_hashes={}, report_paths={"html": str(report)}
    )
    outside = history.record_run(
        qc_result, file_hashes={}, report_paths={"html": str(outsider)}
    )

    bundle = export_runs_archive(
        history.list_runs(), work_dir / "exports" / "runs.zip", managed_root=runs_root
    )

    with zipfile.ZipFile(bundle) as archive:
        names = set(archive.namelist())
        manifest = json.loads(archive.read("manifest.json"))
    assert f"run-{inside}/html.html" in names
    assert not any(name.startswith(f"run-{outside}/") for name in names)
    assert {entry["run_id"] for entry in manifest["runs"]} == {inside, outside}
    # A manifest is evidence, not a path leak.
    assert str(work_dir) not in json.dumps(manifest)
    assert bundle.stat().st_mode & 0o077 == 0


def test_get_run_roundtrips_findings(qc_result: QCRunResult, tmp_path: Path) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    run_id = history.record_run(
        qc_result, file_hashes={}, report_paths={"excel": "runs/1/report.xlsx"}
    )
    record = history.get_run(run_id)
    assert len(record.findings) == len(qc_result.findings)
    assert record.findings[0] == qc_result.findings[0]  # full pydantic round-trip
    assert record.report_paths == {"excel": "runs/1/report.xlsx"}
    assert record.mode == qc_result.mode
    assert record.coverage == qc_result.coverage


def test_get_missing_run_raises(tmp_path: Path) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    with pytest.raises(KeyError):
        history.get_run(999)


def test_history_survives_reopen(qc_result: QCRunResult, tmp_path: Path) -> None:
    db = tmp_path / "history.sqlite3"
    RunHistory(db).record_run(qc_result, file_hashes={}, report_paths={})
    reopened = RunHistory(db)
    runs = reopened.list_runs()
    assert len(runs) == 1
    assert runs[0].started_at <= dt.datetime.now(dt.UTC)


def test_legacy_run_rehydrates_after_optional_contract_expansion(tmp_path: Path) -> None:
    database = tmp_path / "legacy.sqlite3"
    with sqlite3.connect(database) as connection:
        connection.executescript(
            """
            CREATE TABLE runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                started_at TEXT NOT NULL,
                profile TEXT NOT NULL,
                files TEXT NOT NULL,
                file_hashes TEXT NOT NULL,
                counts TEXT NOT NULL,
                disclosures TEXT NOT NULL,
                verified_crosschecks INTEGER NOT NULL,
                findings TEXT NOT NULL,
                report_paths TEXT NOT NULL
            );
            """
        )
        connection.execute(
            """
            INSERT INTO runs (
                started_at, profile, files, file_hashes, counts,
                disclosures, verified_crosschecks, findings, report_paths
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                "2026-07-26T10:00:00+00:00",
                "legacy",
                json.dumps({"current_excel": "old.xlsx"}),
                "{}",
                json.dumps({"warning": 1}),
                "[]",
                0,
                json.dumps(
                    [
                        {
                            "artifact": "excel",
                            "finding_class": "value_changed",
                            "expected_growth": True,
                            "materiality": "recent_restatement",
                            "message": "legacy finding",
                        }
                    ]
                ),
                "{}",
            ),
        )

    record = RunHistory(database).get_run(1)

    assert record.mode.value == "cycle_comparison"
    assert record.coverage == []
    assert record.review_counts == {}
    assert len(record.findings) == 1
    finding = record.findings[0]
    assert finding.message == "legacy finding"
    assert finding.expected_growth is True
    assert finding.expected_reason is None
    assert finding.materiality is not None
    assert finding.materiality.value == "recent_restatement"
    assert finding.temporal_context is None
    assert finding.evidence_tags == set()
    assert finding.impacts == []
    assert finding.current_excerpt is None


def test_annotations_persist_and_apply(qc_result: QCRunResult, tmp_path: Path) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    run_id = history.record_run(qc_result, file_hashes={}, report_paths={})
    target = qc_result.findings[0].finding_id  # a critical finding

    history.set_annotation(
        run_id, target, severity="info", comment="expected per client email 07-25"
    )
    record = history.get_run(run_id)
    annotated = next(f for f in record.findings if f.finding_id == target)
    assert annotated.severity is not None and annotated.severity.value == "info"
    assert annotated.severity_overridden
    assert annotated.analyst_comment == "expected per client email 07-25"
    # Counts recomputed from findings reflect the override.
    from qc_tool.engine import QCRunResult as Result

    rebuilt = Result(profile_name=record.profile, findings=record.findings)
    assert rebuilt.counts != qc_result.counts

    # Upsert: comment-only update keeps engine severity for other findings.
    other = qc_result.findings[1].finding_id
    history.set_annotation(run_id, other, severity=None, comment="checked, fine")
    record = history.get_run(run_id)
    noted = next(f for f in record.findings if f.finding_id == other)
    assert not noted.severity_overridden
    assert noted.analyst_comment == "checked, fine"


def test_bulk_annotations_commit_as_one_group_decision(
    qc_result: QCRunResult, tmp_path: Path
) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    run_id = history.record_run(qc_result, file_hashes={}, report_paths={})
    targets = [finding.finding_id for finding in qc_result.findings[:3]]

    history.set_annotations_bulk(
        run_id,
        [(target, "expected", "reviewed as one range") for target in targets],
    )

    record = history.get_run(run_id)
    reviewed = [finding for finding in record.findings if finding.finding_id in targets]
    assert len(reviewed) == 3
    assert all(finding.severity is not None for finding in reviewed)
    assert all(finding.severity.value == "expected" for finding in reviewed if finding.severity)
    assert all(finding.severity_overridden for finding in reviewed)
    assert all(finding.analyst_comment == "reviewed as one range" for finding in reviewed)


def test_annotations_flow_into_reports(qc_result: QCRunResult, tmp_path: Path) -> None:
    from openpyxl import load_workbook

    from qc_tool.engine import QCRunResult as Result
    from qc_tool.report.excel_report import write_excel_report
    from qc_tool.report.html_report import render_html_report

    history = RunHistory(tmp_path / "history.sqlite3")
    run_id = history.record_run(qc_result, file_hashes={}, report_paths={})
    target = qc_result.findings[0].finding_id
    history.set_annotation(run_id, target, severity="info", comment="reviewed & accepted")
    record = history.get_run(run_id)
    result = Result(profile_name=record.profile, findings=record.findings)

    path = tmp_path / "annotated.xlsx"
    write_excel_report(result, path)
    sheet = load_workbook(path)["Findings"]
    assert sheet["B2"].value == "info *"  # override marker
    assert sheet["R2"].value == "reviewed & accepted"

    html = render_html_report(result)
    assert r'"comment": "reviewed \u0026 accepted"' in html
    assert '"overridden": true' in html
    assert "member.severity + (member.overridden ? ' *' : '')" in html
