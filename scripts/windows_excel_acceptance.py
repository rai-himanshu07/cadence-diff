"""Run privacy-safe live Windows acceptance probes for XLSB Excel automation."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import subprocess
import sys
import tempfile
import time
from contextlib import suppress
from pathlib import Path
from typing import Any

import qc_tool.io.excel_formula as excel_formula_module
from qc_tool.io.excel_formula import (
    _run_excel_worker,
    _status_identity,
    _worker_request,
    _write_private_windows_file,
)
from qc_tool.io.excel_formula_worker import _excel_process
from qc_tool.io.formula_enrichment import FormulaEnrichmentError
from qc_tool.io.xlsb_formula import XlsbFormulaScan, scan_xlsb_formulas
from qc_tool.security import restrict_windows_path_to_current_user

_SCHEMA_VERSION = 1
_PARENT_DEATH_TIMEOUT_SECONDS = 45.0
_PROCESS_EXIT_TIMEOUT_SECONDS = 20.0
_TIMEOUT_SETUP_SECONDS = 45.0


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _write_json(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=True, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    os.replace(temporary, path)


def _dacl_summary(path: Path, *, expected_flags: int) -> dict[str, object]:
    import ntsecuritycon  # pyright: ignore[reportMissingModuleSource]
    import win32api  # pyright: ignore[reportMissingModuleSource]
    import win32con  # pyright: ignore[reportMissingModuleSource]
    import win32security  # pyright: ignore[reportMissingModuleSource]

    token = win32security.OpenProcessToken(
        win32api.GetCurrentProcess(), win32con.TOKEN_QUERY
    )
    try:
        current_sid = win32security.GetTokenInformation(
            token, win32security.TokenUser
        )[0]
    finally:
        win32api.CloseHandle(token)

    descriptor = win32security.GetFileSecurity(
        str(path), win32security.DACL_SECURITY_INFORMATION
    )
    dacl = descriptor.GetSecurityDescriptorDacl()
    if dacl is None:
        return {
            "ace_count": 0,
            "current_user_only": False,
            "expected_inheritance": False,
            "full_control": False,
            "passed": False,
        }

    ace_count = dacl.GetAceCount()
    current_user_only = ace_count == 1
    expected_inheritance = ace_count == 1
    full_control = ace_count == 1
    if ace_count == 1:
        header, access_mask, sid = dacl.GetAce(0)
        ace_type, ace_flags = header
        current_user_only = bool(
            ace_type == win32security.ACCESS_ALLOWED_ACE_TYPE
            and sid == current_sid
        )
        expected_inheritance = ace_flags == expected_flags
        full_control = (
            access_mask & ntsecuritycon.FILE_ALL_ACCESS
        ) == ntsecuritycon.FILE_ALL_ACCESS

    passed = current_user_only and expected_inheritance and full_control
    return {
        "ace_count": ace_count,
        "current_user_only": current_user_only,
        "expected_inheritance": expected_inheritance,
        "full_control": full_control,
        "passed": passed,
    }


def _probe_dacl(data: bytes) -> dict[str, object]:
    import win32con  # pyright: ignore[reportMissingModuleSource]

    with tempfile.TemporaryDirectory(prefix="qc-tool-dacl-probe-") as temporary:
        work_dir = Path(temporary)
        restrict_windows_path_to_current_user(work_dir, inherit=True)
        input_path = work_dir / "input.xlsb"
        _write_private_windows_file(input_path, data)
        directory = _dacl_summary(
            work_dir,
            expected_flags=(
                win32con.CONTAINER_INHERIT_ACE | win32con.OBJECT_INHERIT_ACE
            ),
        )
        source_copy = _dacl_summary(input_path, expected_flags=0)
    return {
        "directory": directory,
        "source_copy": source_copy,
        "passed": bool(directory["passed"] and source_copy["passed"]),
    }


def _process_state(identity: tuple[int, float]) -> str:
    import win32api  # pyright: ignore[reportMissingModuleSource]
    import win32con  # pyright: ignore[reportMissingModuleSource]
    import win32event  # pyright: ignore[reportMissingModuleSource]
    import win32process  # pyright: ignore[reportMissingModuleSource]

    pid, expected_created = identity
    try:
        pids = set(win32process.EnumProcesses())
    except OSError:
        return "unverifiable"
    if pid not in pids:
        return "exited"

    query_limited = getattr(win32con, "PROCESS_QUERY_LIMITED_INFORMATION", 0x1000)
    try:
        handle = win32api.OpenProcess(
            query_limited | win32con.SYNCHRONIZE,
            False,
            pid,
        )
    except OSError:
        return "unverifiable"
    try:
        created = win32process.GetProcessTimes(handle)["CreationTime"].timestamp()
        if abs(created - expected_created) > 0.01:
            return "pid_reused"
        if win32event.WaitForSingleObject(handle, 0) == win32con.WAIT_TIMEOUT:
            return "alive"
        return "exited"
    except OSError:
        return "unverifiable"
    finally:
        win32api.CloseHandle(handle)


def _wait_for_process_exit(
    identity: tuple[int, float], *, timeout: float = _PROCESS_EXIT_TIMEOUT_SECONDS
) -> str:
    deadline = time.monotonic() + timeout
    state = _process_state(identity)
    while state == "alive" and time.monotonic() < deadline:
        time.sleep(0.1)
        state = _process_state(identity)
    return state


def _prepare_worker(
    work_dir: Path, data: bytes, scan: XlsbFormulaScan
) -> dict[str, object]:
    work_dir.mkdir(parents=True, exist_ok=True)
    restrict_windows_path_to_current_user(work_dir, inherit=True)
    _write_private_windows_file(work_dir / "input.xlsb", data)
    return _worker_request(work_dir, scan)


def _run_forced_timeout(
    work_dir: Path, request: dict[str, object], timeout_seconds: float
) -> tuple[str, str]:
    """Start owned Excel before exercising the unchanged production timeout path."""
    original_popen = excel_formula_module.subprocess.Popen
    status_path = Path(str(request["status_path"]))
    setup_outcome = "status_not_observed"

    def recording_popen(
        command: object, *args: Any, **kwargs: Any
    ) -> subprocess.Popen[str]:
        nonlocal setup_outcome
        del command
        process = original_popen(
            [
                sys.executable,
                "-m",
                "scripts.windows_excel_acceptance",
                "--timeout-child",
                str(status_path),
            ],
            *args,
            **kwargs,
        )
        deadline = time.monotonic() + _TIMEOUT_SETUP_SECONDS
        while time.monotonic() < deadline:
            if _status_identity(status_path) is not None:
                setup_outcome = "owned_excel_started"
                break
            if process.poll() is not None:
                setup_outcome = "timeout_child_exited_before_status"
                break
            time.sleep(0.02)
        return process

    excel_formula_module.subprocess.Popen = recording_popen
    try:
        timeout_request = dict(request)
        timeout_request["formula_cells"] = {}
        timeout_request["formula_count"] = 0
        try:
            _run_excel_worker(work_dir, timeout_request, timeout_seconds)
        except FormulaEnrichmentError as exc:
            outcome = (
                "timeout"
                if "exceeded" in str(exc).lower()
                else "worker_failed_before_timeout"
            )
        else:
            outcome = "completed_before_timeout"
    finally:
        excel_formula_module.subprocess.Popen = original_popen
    return outcome, setup_outcome


def _probe_timeout(
    data: bytes, scan: XlsbFormulaScan, *, timeout_seconds: float
) -> dict[str, object]:
    with tempfile.TemporaryDirectory(prefix="qc-tool-timeout-probe-") as temporary:
        work_dir = Path(temporary)
        request = _prepare_worker(work_dir, data, scan)
        started = time.monotonic()
        outcome, setup_outcome = _run_forced_timeout(
            work_dir, request, timeout_seconds
        )
        elapsed_seconds = round(time.monotonic() - started, 3)
        identity = _status_identity(Path(str(request["status_path"])))
        process_state = (
            _wait_for_process_exit(identity) if identity is not None else "not_observed"
        )
    passed = bool(
        outcome == "timeout"
        and setup_outcome == "owned_excel_started"
        and process_state in {"exited", "pid_reused"}
    )
    return {
        "elapsed_seconds": elapsed_seconds,
        "outcome": outcome,
        "owned_excel_process": process_state,
        "status_observed": identity is not None,
        "timeout_setup": setup_outcome,
        "timeout_seconds": timeout_seconds,
        "passed": passed,
    }


def _timeout_child(status_path: Path) -> int:
    import pythoncom  # pyright: ignore[reportMissingModuleSource]
    import win32api  # pyright: ignore[reportMissingModuleSource]
    import win32com.client  # pyright: ignore[reportMissingModuleSource]

    pythoncom.CoInitialize()
    app = excel_handle = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.AutomationSecurity = 3
        app.EnableEvents = False
        app.DisplayAlerts = False
        app.Interactive = False
        pid, excel_handle, created = _excel_process(app)
        _write_json(status_path, {"pid": pid, "created": created})
        win32api.CloseHandle(excel_handle)
        excel_handle = None
        sys.stdin.read()
        time.sleep(600.0)
        return 4
    finally:
        if excel_handle is not None:
            win32api.CloseHandle(excel_handle)
        if app is not None:
            with suppress(Exception):
                app.Quit()
        pythoncom.CoUninitialize()


def _wait_for_status(
    status_path: Path,
    process: subprocess.Popen[bytes],
    *,
    timeout: float,
) -> tuple[int, float] | None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        identity = _status_identity(status_path)
        if identity is not None:
            return identity
        if process.poll() is not None:
            return None
        time.sleep(0.1)
    return None


def _parent_death_child(work_dir: Path) -> int:
    data = (work_dir / "input.xlsb").read_bytes()
    scan = scan_xlsb_formulas(data)
    request = _worker_request(work_dir, scan)
    try:
        _run_excel_worker(work_dir, request, 600.0)
    except FormulaEnrichmentError:
        return 3
    return 0


def _probe_parent_death(data: bytes, scan: XlsbFormulaScan) -> dict[str, object]:
    with tempfile.TemporaryDirectory(
        prefix="qc-tool-parent-death-probe-", ignore_cleanup_errors=True
    ) as temporary:
        work_dir = Path(temporary)
        request = _prepare_worker(work_dir, data, scan)
        command = [
            sys.executable,
            "-m",
            "scripts.windows_excel_acceptance",
            "--parent-death-child",
            str(work_dir),
        ]
        parent = subprocess.Popen(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            cwd=Path(__file__).resolve().parents[1],
            close_fds=True,
            creationflags=getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
        )
        identity: tuple[int, float] | None = None
        try:
            identity = _wait_for_status(
                Path(str(request["status_path"])),
                parent,
                timeout=_PARENT_DEATH_TIMEOUT_SECONDS,
            )
            parent_was_running = parent.poll() is None
            if identity is not None and parent_was_running:
                time.sleep(0.5)
                parent.kill()
                parent.wait(timeout=10.0)
            process_state = (
                _wait_for_process_exit(identity)
                if identity is not None
                else "not_observed"
            )
        finally:
            if parent.poll() is None:
                parent.kill()
                parent.wait(timeout=10.0)
        result_written = Path(str(request["result_path"])).exists()
    passed = bool(
        identity is not None
        and parent_was_running
        and process_state in {"exited", "pid_reused"}
        and not result_written
    )
    return {
        "owned_excel_process": process_state,
        "parent_was_running": parent_was_running,
        "result_written": result_written,
        "status_observed": identity is not None,
        "passed": passed,
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Run aggregate-only DACL, timeout, and parent-death probes against "
            "the production Windows Excel worker."
        )
    )
    parser.add_argument("--xlsb", type=Path, required=True)
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("windows-excel-acceptance.json"),
    )
    parser.add_argument("--timeout-seconds", type=float, default=2.0)
    parser.add_argument(
        "--probe",
        choices=("all", "dacl", "timeout", "parent-death"),
        default="all",
    )
    return parser


def _run_acceptance(args: argparse.Namespace) -> tuple[int, dict[str, object]]:
    if os.name != "nt":
        return 2, {
            "schema_version": _SCHEMA_VERSION,
            "overall": "unsupported_platform",
            "passed": False,
        }
    if args.timeout_seconds <= 0:
        return 2, {
            "schema_version": _SCHEMA_VERSION,
            "overall": "invalid_timeout",
            "passed": False,
        }
    try:
        data = args.xlsb.read_bytes()
        scan = scan_xlsb_formulas(data)
    except (OSError, ValueError):
        return 2, {
            "schema_version": _SCHEMA_VERSION,
            "overall": "invalid_source",
            "passed": False,
        }
    if not scan.safe_for_external_engine or scan.formula_count == 0:
        return 2, {
            "schema_version": _SCHEMA_VERSION,
            "formula_count": scan.formula_count,
            "overall": "source_not_probe_safe",
            "passed": False,
            "risky_feature_count": len(scan.risky_features),
        }

    source_hash = _sha256(data)
    probes: dict[str, dict[str, object]] = {}
    requested_probe = getattr(args, "probe", "all")
    requested = (
        ("dacl", "timeout", "parent_death")
        if requested_probe == "all"
        else (requested_probe.replace("-", "_"),)
    )
    for probe_name in requested:
        try:
            if probe_name == "dacl":
                probes[probe_name] = _probe_dacl(data)
            elif probe_name == "timeout":
                probes[probe_name] = _probe_timeout(
                    data, scan, timeout_seconds=args.timeout_seconds
                )
            else:
                probes[probe_name] = _probe_parent_death(data, scan)
        except Exception as exc:
            probes[probe_name] = {
                "error_type": type(exc).__name__,
                "passed": False,
                "stage": "probe_exception",
            }

    source_unchanged = _sha256(args.xlsb.read_bytes()) == source_hash
    passed = bool(
        source_unchanged
        and set(requested) == set(probes)
        and all(probe.get("passed") is True for probe in probes.values())
    )
    payload: dict[str, object] = {
        "schema_version": _SCHEMA_VERSION,
        "formula_count": scan.formula_count,
        "overall": "passed" if passed else "failed",
        "passed": passed,
        "probes": probes,
        "requested_probes": list(requested),
        "risky_feature_count": len(scan.risky_features),
        "source_sha256": source_hash,
        "source_unchanged": source_unchanged,
    }
    return (0 if passed else 1), payload


def main(argv: list[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if argv is None else argv)
    if arguments and arguments[0] == "--timeout-child":
        if len(arguments) != 2:
            return 2
        return _timeout_child(Path(arguments[1]))
    if arguments and arguments[0] == "--parent-death-child":
        if len(arguments) != 2:
            return 2
        return _parent_death_child(Path(arguments[1]))

    args = _parser().parse_args(arguments)
    exit_code, payload = _run_acceptance(args)
    _write_json(args.output, payload)
    print(f"Windows Excel acceptance: {payload['overall']}")
    print(f"Evidence: {args.output.name}")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
