"""Disposable setup-scan worker: safe synthetic end-to-end tests
(plan-20260913, Step 6).
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from qc_tool.history.store import sha256_file
from qc_tool.setup.preview_worker import SetupScanRequest, run_setup_scan_worker
from tests.fixtures.generate import encrypt_file


def _write_workbook(path: Path, *, rows: int = 3) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet.append(["ID", "Value"])
    for i in range(rows):
        sheet.append([f"A{i}", i])
    workbook.save(path)


def test_worker_scans_a_simple_pair_end_to_end(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path)
    _write_workbook(current_path, rows=4)

    request = SetupScanRequest(
        member_id="primary",
        baseline_path=str(baseline_path),
        current_path=str(current_path),
        baseline_hash=sha256_file(baseline_path),
        current_hash=sha256_file(current_path),
    )

    outcome = run_setup_scan_worker(request)

    assert outcome.disclosure == ""
    assert outcome.missing_credential_roles == ()
    assert outcome.result_payload is not None
    assert outcome.result_payload["member_id"] == "primary"
    current_sheets = outcome.result_payload["current_sheets"]
    assert isinstance(current_sheets, list)
    assert current_sheets[0]["sheet_name"] == "Data"
    assert current_sheets[0]["regions"][0]["region"]["max_row"] == 5


def test_worker_discloses_a_hash_mismatch(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path)
    _write_workbook(current_path)

    request = SetupScanRequest(
        member_id="primary",
        baseline_path=str(baseline_path),
        current_path=str(current_path),
        baseline_hash="0" * 64,  # deliberately wrong
        current_hash=sha256_file(current_path),
    )

    outcome = run_setup_scan_worker(request)

    assert outcome.result_payload is None
    assert "changed since it was uploaded" in outcome.disclosure


def test_worker_discloses_a_missing_source_file(tmp_path: Path) -> None:
    current_path = tmp_path / "current.xlsx"
    _write_workbook(current_path)

    request = SetupScanRequest(
        member_id="primary",
        baseline_path=str(tmp_path / "does-not-exist.xlsx"),
        current_path=str(current_path),
        baseline_hash="",
        current_hash=sha256_file(current_path),
    )

    outcome = run_setup_scan_worker(request)

    assert outcome.result_payload is None
    assert "no longer at its recorded location" in outcome.disclosure


def test_worker_reports_a_missing_credential_role_without_a_password(
    tmp_path: Path,
) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    current_plain = tmp_path / "current_plain.xlsx"
    _write_workbook(baseline_path)
    _write_workbook(current_plain)
    encrypt_file(current_plain, current_path, "hunter2")

    request = SetupScanRequest(
        member_id="primary",
        baseline_path=str(baseline_path),
        current_path=str(current_path),
        baseline_hash=sha256_file(baseline_path),
        current_hash=sha256_file(current_path),
    )

    outcome = run_setup_scan_worker(request)

    assert outcome.result_payload is None
    assert outcome.missing_credential_roles == ("current",)
    assert outcome.disclosure == ""


def test_worker_scans_successfully_with_the_right_password(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    current_plain = tmp_path / "current_plain.xlsx"
    _write_workbook(baseline_path)
    _write_workbook(current_plain)
    encrypt_file(current_plain, current_path, "hunter2")

    request = SetupScanRequest(
        member_id="primary",
        baseline_path=str(baseline_path),
        current_path=str(current_path),
        baseline_hash=sha256_file(baseline_path),
        current_hash=sha256_file(current_path),
        passwords={"current": "hunter2"},
    )

    outcome = run_setup_scan_worker(request)

    assert outcome.missing_credential_roles == ()
    assert outcome.disclosure == ""
    assert outcome.result_payload is not None
