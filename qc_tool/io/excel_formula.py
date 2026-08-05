"""Parent-side Windows Excel formula extraction through a bounded worker."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import TypeAlias

from qc_tool.io.formula_enrichment import (
    FormulaEnrichmentError,
    FormulaExtraction,
    FormulaMap,
    validate_formula_extraction,
)
from qc_tool.io.xlsb_formula import XlsbFormulaScan

DEFAULT_EXCEL_TIMEOUT_SECONDS = 300.0
_WorkerRunner: TypeAlias = Callable[
    [Path, dict[str, object], float], FormulaExtraction
]


def _write_private_windows_file(path: Path, data: bytes) -> None:
    path.write_bytes(data)
    _restrict_windows_path(path, inherit=False)


def _restrict_windows_path(path: Path, *, inherit: bool) -> None:
    try:
        import ntsecuritycon  # pyright: ignore[reportMissingModuleSource]
        import win32api  # pyright: ignore[reportMissingModuleSource]
        import win32con  # pyright: ignore[reportMissingModuleSource]
        import win32security  # pyright: ignore[reportMissingModuleSource]
    except ImportError as exc:  # pragma: no cover - exercised on Windows deployment
        raise FormulaEnrichmentError("pywin32 is required for Windows formula enrichment") from exc

    token = win32security.OpenProcessToken(win32api.GetCurrentProcess(), win32con.TOKEN_QUERY)
    try:
        user_sid = win32security.GetTokenInformation(token, win32security.TokenUser)[0]
    finally:
        win32api.CloseHandle(token)
    dacl = win32security.ACL()
    inheritance = (
        win32con.CONTAINER_INHERIT_ACE | win32con.OBJECT_INHERIT_ACE
        if inherit
        else 0
    )
    dacl.AddAccessAllowedAceEx(
        win32security.ACL_REVISION_DS,
        inheritance,
        ntsecuritycon.FILE_ALL_ACCESS,
        user_sid,
    )
    descriptor = win32security.SECURITY_DESCRIPTOR()
    descriptor.SetSecurityDescriptorDacl(True, dacl, False)
    win32security.SetFileSecurity(
        str(path), win32security.DACL_SECURITY_INFORMATION, descriptor
    )


def _worker_request(work_dir: Path, scan: XlsbFormulaScan) -> dict[str, object]:
    return {
        "schema_version": 1,
        "parent_pid": os.getpid(),
        "input_path": str(work_dir / "input.xlsb"),
        "result_path": str(work_dir / "result.json"),
        "status_path": str(work_dir / "excel-status.json"),
        "formula_cells": {
            sheet: [[row, column] for row, column in sorted(cells)]
            for sheet, cells in scan.formula_cells.items()
            if cells
        },
        "formula_count": scan.formula_count,
    }


def _load_worker_result(path: Path) -> FormulaExtraction:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise FormulaEnrichmentError("Excel formula worker returned no valid result") from exc
    if not isinstance(payload, dict):
        raise FormulaEnrichmentError("Excel formula worker result is not an object")
    if payload.get("schema_version") != 1:
        raise FormulaEnrichmentError("Excel formula worker returned an unknown schema")
    if payload.get("ok") is not True:
        detail = payload.get("error")
        message = detail if isinstance(detail, str) else "unknown worker failure"
        raise FormulaEnrichmentError(f"Excel formula worker failed: {message[:300]}")

    raw_formulas = payload.get("formulas")
    if not isinstance(raw_formulas, dict):
        raise FormulaEnrichmentError("Excel formula worker omitted its formula map")
    formulas: FormulaMap = {}
    for sheet, raw_cells in raw_formulas.items():
        if not isinstance(sheet, str) or not isinstance(raw_cells, list):
            raise FormulaEnrichmentError("Excel formula worker returned malformed formula data")
        cells: dict[tuple[int, int], str] = {}
        for item in raw_cells:
            if not isinstance(item, dict):
                raise FormulaEnrichmentError("Excel formula worker returned malformed cell data")
            row = item.get("row")
            column = item.get("column")
            formula = item.get("formula")
            if (
                not isinstance(row, int)
                or isinstance(row, bool)
                or not isinstance(column, int)
                or isinstance(column, bool)
                or not isinstance(formula, str)
                or row < 1
                or column < 1
            ):
                raise FormulaEnrichmentError("Excel formula worker returned invalid cell data")
            coordinate = (row, column)
            if coordinate in cells:
                raise FormulaEnrichmentError(
                    f"Excel formula worker duplicated {sheet}!{coordinate}"
                )
            cells[coordinate] = formula
        if cells:
            formulas[sheet] = cells

    engine = payload.get("engine")
    detail = payload.get("detail")
    if not isinstance(engine, str) or not isinstance(detail, str):
        raise FormulaEnrichmentError("Excel formula worker omitted engine provenance")
    return FormulaExtraction(formulas=formulas, engine=engine, detail=detail)


def _status_identity(path: Path) -> tuple[int, float] | None:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(payload, dict):
        return None
    pid = payload.get("pid")
    created = payload.get("created")
    if (
        isinstance(pid, int)
        and not isinstance(pid, bool)
        and pid > 0
        and isinstance(created, int | float)
        and not isinstance(created, bool)
    ):
        return pid, float(created)
    return None


def _terminate_owned_excel(status_path: Path) -> None:
    identity = _status_identity(status_path)
    if identity is None:
        return
    pid, expected_created = identity
    try:
        import pywintypes  # pyright: ignore[reportMissingModuleSource]
        import win32api  # pyright: ignore[reportMissingModuleSource]
        import win32con  # pyright: ignore[reportMissingModuleSource]
        import win32event  # pyright: ignore[reportMissingModuleSource]
        import win32process  # pyright: ignore[reportMissingModuleSource]
    except ImportError:
        return
    try:
        handle = win32api.OpenProcess(
            win32con.PROCESS_QUERY_INFORMATION
            | win32con.PROCESS_TERMINATE
            | win32con.SYNCHRONIZE,
            False,
            pid,
        )
    except (OSError, pywintypes.error):
        return
    try:
        if win32event.WaitForSingleObject(handle, 10_000) != win32con.WAIT_TIMEOUT:
            return
        created = win32process.GetProcessTimes(handle)["CreationTime"].timestamp()
        if abs(created - expected_created) > 0.01:
            return
        if win32event.WaitForSingleObject(handle, 0) == win32con.WAIT_TIMEOUT:
            win32api.TerminateProcess(handle, 1)
            win32event.WaitForSingleObject(handle, 10_000)
    except (OSError, pywintypes.error):
        # Windows can deny process queries or termination after the handle opens.
        # Never turn best-effort owned-process cleanup into a failed QC run.
        return
    finally:
        win32api.CloseHandle(handle)


def _run_excel_worker(
    work_dir: Path, request: dict[str, object], timeout: float
) -> FormulaExtraction:
    status_path = Path(str(request["status_path"]))
    result_path = Path(str(request["result_path"]))
    creation_flags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    process = subprocess.Popen(
        [sys.executable, "-m", "qc_tool.io.excel_formula_worker"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=os.environ.copy(),
        close_fds=True,
        creationflags=creation_flags,
    )
    try:
        _, stderr = process.communicate(json.dumps(request), timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        _terminate_owned_excel(status_path)
        process.kill()
        process.communicate()
        raise FormulaEnrichmentError(
            f"Excel formula extraction exceeded {timeout:g} seconds"
        ) from exc
    finally:
        if process.poll() is not None:
            _terminate_owned_excel(status_path)
    if process.returncode != 0 and not result_path.exists():
        detail = stderr.strip().splitlines()[-1] if stderr.strip() else "no diagnostic"
        raise FormulaEnrichmentError(
            f"Excel formula worker exited with code {process.returncode}: {detail[:200]}"
        )
    return _load_worker_result(result_path)


def extract_formulas_with_excel(
    data: bytes,
    scan: XlsbFormulaScan,
    *,
    timeout: float = DEFAULT_EXCEL_TIMEOUT_SECONDS,
    worker_runner: _WorkerRunner | None = None,
) -> FormulaExtraction:
    """Extract formula text through desktop Excel without trusting Excel values."""
    if not scan.safe_for_external_engine:
        raise FormulaEnrichmentError(
            "XLSB contains active or external content: " + ", ".join(scan.risky_features)
        )
    if worker_runner is None and os.name != "nt":
        raise FormulaEnrichmentError("desktop Excel formula enrichment requires Windows")

    with tempfile.TemporaryDirectory(prefix="qc-tool-xlsb-") as temporary:
        work_dir = Path(temporary)
        if os.name == "nt":
            _restrict_windows_path(work_dir, inherit=True)
        input_path = work_dir / "input.xlsb"
        if os.name == "nt":
            _write_private_windows_file(input_path, data)
        else:
            input_path.write_bytes(data)
        request = _worker_request(work_dir, scan)
        extraction = (worker_runner or _run_excel_worker)(work_dir, request, timeout)
        validate_formula_extraction(scan, extraction)
        return extraction
