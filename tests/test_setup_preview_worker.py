"""Disposable setup-scan worker: safe synthetic end-to-end tests
(plan-20260913, Step 6).
"""

from __future__ import annotations

import threading
from pathlib import Path

from openpyxl import Workbook

from qc_tool.history.store import sha256_file
from qc_tool.setup.preview_worker import (
    PreviewWindowRequest,
    SetupScanRequest,
    run_preview_window_worker,
    run_setup_scan_worker,
)
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


def test_worker_honors_a_pre_set_cancel_event(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path)
    _write_workbook(current_path)

    request = SetupScanRequest(
        member_id="primary",
        baseline_path=str(baseline_path),
        current_path=str(current_path),
        baseline_hash=sha256_file(baseline_path),
        current_hash=sha256_file(current_path),
    )
    cancel_event = threading.Event()
    cancel_event.set()

    outcome = run_setup_scan_worker(request, cancel_event=cancel_event)

    assert outcome.result_payload is None
    assert outcome.disclosure == "setup scan cancelled"


def test_worker_still_completes_when_cancel_event_is_never_set(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path)
    _write_workbook(current_path)

    request = SetupScanRequest(
        member_id="primary",
        baseline_path=str(baseline_path),
        current_path=str(current_path),
        baseline_hash=sha256_file(baseline_path),
        current_hash=sha256_file(current_path),
    )

    outcome = run_setup_scan_worker(request, cancel_event=threading.Event())

    assert outcome.disclosure == ""
    assert outcome.result_payload is not None


def test_preview_window_worker_fetches_a_bounded_grid(tmp_path: Path) -> None:
    current_path = tmp_path / "current.xlsx"
    _write_workbook(current_path, rows=5)

    request = PreviewWindowRequest(
        side="current",
        path=str(current_path),
        source_hash=sha256_file(current_path),
        sheet="Data",
        min_row=1,
        min_col=1,
        max_row=3,
        max_col=2,
    )

    outcome = run_preview_window_worker(request)

    assert outcome.disclosure == ""
    assert outcome.rows == [["ID", "Value"], ["A0", "0"], ["A1", "1"]]
    assert outcome.formula_cells == [[False, False], [False, False], [False, False]]
    assert outcome.formula_text == {}
    assert outcome.resolved_min_row == 1
    assert outcome.resolved_max_row == 3
    assert outcome.resolved_max_col == 2


def test_preview_window_worker_caps_a_huge_request_to_the_bounded_maximum(
    tmp_path: Path,
) -> None:
    current_path = tmp_path / "current.xlsx"
    _write_workbook(current_path, rows=5)

    request = PreviewWindowRequest(
        side="current",
        path=str(current_path),
        source_hash=sha256_file(current_path),
        sheet="Data",
        max_row=10_000,
        max_col=5_000,
    )

    outcome = run_preview_window_worker(request)

    assert outcome.disclosure == ""
    # bounded by the sheet's own extent, never the huge requested bound
    assert outcome.resolved_max_row <= 6
    assert outcome.resolved_max_col <= 40


def test_preview_window_worker_reveals_formula_text_only_when_asked(
    tmp_path: Path,
) -> None:
    current_path = tmp_path / "current.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet["A1"] = "=1+1"
    workbook.save(current_path)

    hidden_request = PreviewWindowRequest(
        side="current",
        path=str(current_path),
        source_hash=sha256_file(current_path),
        sheet="Data",
        max_row=1,
        max_col=1,
        reveal_formulas=False,
    )
    hidden_outcome = run_preview_window_worker(hidden_request)
    assert hidden_outcome.formula_cells == [[True]]
    assert hidden_outcome.formula_text == {}

    revealed_request = PreviewWindowRequest(
        side="current",
        path=str(current_path),
        source_hash=sha256_file(current_path),
        sheet="Data",
        max_row=1,
        max_col=1,
        reveal_formulas=True,
    )
    revealed_outcome = run_preview_window_worker(revealed_request)
    assert revealed_outcome.formula_text == {"1,1": "=1+1"}


def test_preview_window_worker_discloses_a_hash_mismatch(tmp_path: Path) -> None:
    current_path = tmp_path / "current.xlsx"
    _write_workbook(current_path)

    request = PreviewWindowRequest(
        side="current",
        path=str(current_path),
        source_hash="0" * 64,
        sheet="Data",
    )

    outcome = run_preview_window_worker(request)

    assert outcome.rows == []
    assert "changed since it was uploaded" in outcome.disclosure


def test_preview_window_worker_reports_a_missing_credential(tmp_path: Path) -> None:
    current_plain = tmp_path / "current_plain.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(current_plain)
    encrypt_file(current_plain, current_path, "hunter2")

    request = PreviewWindowRequest(
        side="current",
        path=str(current_path),
        source_hash=sha256_file(current_path),
        sheet="Data",
    )

    outcome = run_preview_window_worker(request)

    assert outcome.missing_credential is True
    assert outcome.rows == []
