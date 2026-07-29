"""Run history tests (criterion 13)."""

import datetime as dt
import json
import sqlite3
from pathlib import Path

import pytest

from qc_tool.engine import QCRunResult
from qc_tool.history.store import RunHistory, sha256_file


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
    assert newest.started_at.tzinfo is not None  # timezone-aware UTC
    assert newest.findings == []  # listing stays lightweight


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
    assert len(record.findings) == 1
    finding = record.findings[0]
    assert finding.message == "legacy finding"
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
    assert sheet["L2"].value == "reviewed & accepted"

    html = render_html_report(result)
    assert "reviewed &amp; accepted" in html
    assert "info *" in html
