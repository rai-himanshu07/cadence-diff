"""Bounded on-demand key-overlap query worker: safe synthetic end-to-end
tests (plan-20260913, Step 12 Fix 5).
"""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

from openpyxl import Workbook

from qc_tool.history.store import sha256_file
from qc_tool.setup.key_overlap_worker import (
    MAX_IDENTITY_COLUMNS,
    KeyOverlapRequest,
    run_key_overlap_worker,
)
from qc_tool.setup.preview_worker import SetupScanRequest, run_setup_scan_worker
from tests.fixtures.generate import encrypt_file


def _write_workbook(path: Path, rows: list[tuple[object, object]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet.append(["ID", "Amount"])
    for identity, amount in rows:
        sheet.append([identity, amount])
    workbook.save(path)


def _request(
    baseline_path: Path,
    current_path: Path,
    *,
    identity_columns: tuple[str, ...] = ("A",),
    baseline_identity_columns: tuple[str, ...] = (),
    baseline_password: str = "",
    current_password: str = "",
    trim_identity_whitespace: bool = False,
) -> KeyOverlapRequest:
    sidecar_path = baseline_path.parent / "setup-inspection.sqlite3"
    baseline_hash = sha256_file(baseline_path)
    current_hash = sha256_file(current_path)
    scan = run_setup_scan_worker(
        SetupScanRequest(
            member_id="primary",
            baseline_path=str(baseline_path),
            current_path=str(current_path),
            baseline_hash=baseline_hash,
            current_hash=current_hash,
            sidecar_path=str(sidecar_path),
            session_key="session-1",
            input_generation=1,
            passwords={
                "baseline": baseline_password,
                "current": current_password,
            },
        )
    )
    if baseline_password or current_password:
        assert scan.result_payload is not None
    return KeyOverlapRequest(
        sidecar_path=str(sidecar_path),
        session_key="session-1",
        input_generation=1,
        member_id="primary",
        baseline_hash=baseline_hash,
        current_hash=current_hash,
        baseline_sheet="Data",
        current_sheet="Data",
        baseline_first_row=2,  # row 1 is the "ID"/"Amount" header
        baseline_last_row=6,
        current_first_row=2,
        current_last_row=6,
        identity_columns=identity_columns,
        baseline_identity_columns=baseline_identity_columns,
        trim_identity_whitespace=trim_identity_whitespace,
    )


def test_worker_reports_full_overlap_for_identical_unique_keys(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path, [("R1", 1), ("R2", 2), ("R3", 3)])
    _write_workbook(current_path, [("R3", 30), ("R1", 10), ("R2", 20)])  # reordered

    request = _request(baseline_path, current_path)
    baseline_path.unlink()
    current_path.unlink()
    outcome = run_key_overlap_worker(request)

    assert outcome.ok
    assert outcome.ratio == 1.0
    assert outcome.baseline_unique_keys == 3
    assert outcome.current_unique_keys == 3
    assert outcome.overlapping_keys == 3
    assert outcome.disclosure == ""


def test_worker_reports_partial_overlap(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path, [("R1", 1), ("R2", 2), ("R3", 3), ("R4", 4)])
    _write_workbook(current_path, [("R3", 30), ("R4", 40), ("R5", 50), ("R6", 60)])

    outcome = run_key_overlap_worker(_request(baseline_path, current_path))

    # union = {R1,R2,R3,R4,R5,R6} = 6; overlap = {R3,R4} = 2
    assert outcome.ok
    assert outcome.baseline_unique_keys == 4
    assert outcome.current_unique_keys == 4
    assert outcome.overlapping_keys == 2
    assert outcome.ratio is not None
    assert abs(outcome.ratio - (2 / 6)) < 1e-9


def test_worker_excludes_duplicate_and_blank_keys_from_uniqueness(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    # R1 repeats on baseline (never unique there); blank row excluded on current.
    _write_workbook(baseline_path, [("R1", 1), ("R1", 2), ("R2", 3)])
    _write_workbook(current_path, [("R1", 10), ("", 20), ("R2", 30)])

    outcome = run_key_overlap_worker(_request(baseline_path, current_path))

    assert outcome.ok
    assert outcome.baseline_unique_keys == 1  # only R2 (R1 duplicated)
    assert outcome.current_unique_keys == 2  # R1 and R2 (blank row excluded)
    assert outcome.overlapping_keys == 1  # only R2 is unique on both sides
    assert outcome.ratio is not None
    assert abs(outcome.ratio - (1 / 2)) < 1e-9


def test_worker_respects_baseline_identity_column_override(tmp_path: Path) -> None:
    """Step 12 Fix 3 scenario: the confirmed identity column lives at a
    DIFFERENT physical letter on baseline vs current.
    """
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    base_wb = Workbook()
    base_ws = base_wb.active
    assert base_ws is not None
    base_ws.title = "Data"
    base_ws.append(["Amount", "ID"])  # ID is column B on baseline
    base_ws.append([1, "R1"])
    base_ws.append([2, "R2"])
    base_wb.save(baseline_path)

    curr_wb = Workbook()
    curr_ws = curr_wb.active
    assert curr_ws is not None
    curr_ws.title = "Data"
    curr_ws.append(["ID", "Amount"])  # ID is column A on current
    curr_ws.append(["R1", 10])
    curr_ws.append(["R2", 20])
    curr_wb.save(current_path)

    outcome = run_key_overlap_worker(
        _request(
            baseline_path,
            current_path,
            identity_columns=("A",),
            baseline_identity_columns=("B",),
        )
    )

    assert outcome.ok
    assert outcome.ratio == 1.0
    assert outcome.overlapping_keys == 2


def test_worker_trims_outer_whitespace_when_requested(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path, [("R1 ", 1), ("R2", 2)])
    _write_workbook(current_path, [("R1", 10), ("R2", 20)])

    untrimmed = run_key_overlap_worker(_request(baseline_path, current_path))
    assert untrimmed.ok
    assert untrimmed.overlapping_keys == 1  # "R1 " != "R1" without trim

    trimmed = run_key_overlap_worker(
        _request(baseline_path, current_path, trim_identity_whitespace=True)
    )
    assert trimmed.ok
    assert trimmed.overlapping_keys == 2


def test_worker_discloses_a_hash_mismatch(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path, [("R1", 1)])
    _write_workbook(current_path, [("R1", 1)])

    request = replace(
        _request(baseline_path, current_path),
        baseline_hash="0" * 64,
    )

    outcome = run_key_overlap_worker(request)

    assert not outcome.ok
    assert "stale" in outcome.disclosure


def test_worker_discloses_a_missing_sidecar(tmp_path: Path) -> None:
    request = KeyOverlapRequest(
        sidecar_path=str(tmp_path / "missing-sidecar.sqlite3"),
        session_key="session-1",
        input_generation=1,
        member_id="primary",
        baseline_hash="0" * 64,
        current_hash="1" * 64,
        baseline_sheet="Data",
        current_sheet="Data",
        baseline_first_row=1,
        baseline_last_row=2,
        current_first_row=1,
        current_last_row=2,
        identity_columns=("A",),
    )

    outcome = run_key_overlap_worker(request)

    assert not outcome.ok
    assert "sheet not found" in outcome.disclosure


def test_worker_reports_a_missing_credential_role_without_a_password(
    tmp_path: Path,
) -> None:
    baseline_plain = tmp_path / "baseline_plain.xlsx"
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_plain, [("R1", 1)])
    encrypt_file(baseline_plain, baseline_path, "hunter2")
    _write_workbook(current_path, [("R1", 1)])

    outcome = run_key_overlap_worker(_request(baseline_path, current_path))

    assert not outcome.ok
    assert outcome.missing_credential_roles == ()
    assert "sheet not found" in outcome.disclosure


def test_worker_scans_successfully_with_the_right_password(tmp_path: Path) -> None:
    baseline_plain = tmp_path / "baseline_plain.xlsx"
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_plain, [("R1", 1), ("R2", 2)])
    encrypt_file(baseline_plain, baseline_path, "hunter2")
    _write_workbook(current_path, [("R1", 10), ("R2", 20)])

    outcome = run_key_overlap_worker(
        _request(baseline_path, current_path, baseline_password="hunter2")
    )

    assert outcome.ok
    assert outcome.ratio == 1.0


def test_worker_rejects_empty_identity_columns(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path, [("R1", 1)])
    _write_workbook(current_path, [("R1", 1)])

    outcome = run_key_overlap_worker(_request(baseline_path, current_path, identity_columns=()))

    assert not outcome.ok
    assert "no identity columns" in outcome.disclosure


def test_worker_rejects_too_many_identity_columns(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path, [("R1", 1)])
    _write_workbook(current_path, [("R1", 1)])

    too_many = tuple(chr(ord("A") + i) for i in range(MAX_IDENTITY_COLUMNS + 1))
    outcome = run_key_overlap_worker(
        _request(baseline_path, current_path, identity_columns=too_many)
    )

    assert not outcome.ok
    assert "too many identity columns" in outcome.disclosure


def test_worker_discloses_a_missing_sheet(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path, [("R1", 1)])
    _write_workbook(current_path, [("R1", 1)])

    request = replace(
        _request(baseline_path, current_path),
        current_sheet="NoSuchSheet",
    )

    outcome = run_key_overlap_worker(request)

    assert not outcome.ok
    assert "sheet not found" in outcome.disclosure
