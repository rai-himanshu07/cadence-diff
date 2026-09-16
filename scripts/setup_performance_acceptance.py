"""Aggregate-only setup-versus-full-preparation acceptance harness.

Private source paths are discovered internally and never accepted on the command
line, printed, logged, or serialized. Public fixtures can be selected by their
fixed synthetic format name. Output contains only hashed pair labels, counts,
timings, memory peaks, and source-unchanged booleans.
"""

from __future__ import annotations

import argparse
import contextlib
import hashlib
import json
import logging
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Literal

import psutil

_ROOT = Path(__file__).resolve().parents[1]
_PRIVATE_ROOT = _ROOT.parent / "ref_docs"
_PUBLIC_ROOT = _ROOT / "tests" / "fixtures" / "generated"
_EXCEL_SUFFIXES = frozenset({".xlsx", ".xlsm", ".xlsb"})


def _sha256(path: Path) -> bytes:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.digest()


def _safe_label(identity: str) -> str:
    return "pair-" + hashlib.sha256(identity.encode("utf-8")).hexdigest()[:12]


def _discover_private_pairs(root: Path = _PRIVATE_ROOT) -> list[tuple[str, Path, Path]]:
    """Discover pairs while returning only hashed labels to every caller."""
    if not root.is_dir():
        return []
    pairs: list[tuple[str, Path, Path]] = []
    for subdirectory in sorted(path for path in root.iterdir() if path.is_dir()):
        files = sorted(
            path
            for path in subdirectory.iterdir()
            if path.is_file() and path.suffix.casefold() in _EXCEL_SUFFIXES
        )
        if len(files) == 2:
            identity = f"directory:{subdirectory.relative_to(root)}"
            pairs.append((_safe_label(identity), files[0], files[1]))
    top_level = sorted(
        path
        for path in root.iterdir()
        if path.is_file() and path.suffix.casefold() in _EXCEL_SUFFIXES
    )
    by_prefix: dict[str, list[Path]] = {}
    for path in top_level:
        prefix = path.stem.rsplit("_", 1)[0].rsplit("-", 1)[0].strip()
        by_prefix.setdefault(prefix, []).append(path)
    for prefix, files in sorted(by_prefix.items()):
        if len(files) == 2:
            pairs.append((_safe_label(f"top:{prefix}"), files[0], files[1]))
    return pairs


def _public_pair(file_format: Literal["xlsx", "xlsb"]) -> tuple[str, Path, Path]:
    return (
        f"public-{file_format}",
        _PUBLIC_ROOT / f"baseline.{file_format}",
        _PUBLIC_ROOT / f"current.{file_format}",
    )


def _pair(
    *, private_index: int | None, public_format: Literal["xlsx", "xlsb"] | None
) -> tuple[str, Path, Path]:
    if public_format is not None:
        return _public_pair(public_format)
    pairs = _discover_private_pairs()
    if private_index is None or not 0 <= private_index < len(pairs):
        raise ValueError("private pair index is unavailable")
    return pairs[private_index]


def _measure_setup(label: str, baseline: Path, current: Path) -> dict[str, object]:
    from qc_tool.history.store import sha256_file
    from qc_tool.setup.preview_worker import (
        SetupScanProgress,
        SetupScanRequest,
        run_setup_scan_worker,
    )

    before = (_sha256(baseline), _sha256(current))
    baseline_hash = sha256_file(baseline)
    current_hash = sha256_file(current)
    started = time.perf_counter()
    cpu_started = time.process_time()
    first_inventory: float | None = None
    first_sheet: float | None = None
    phase_seconds: dict[str, dict[str, float]] = {}

    def progress(event: SetupScanProgress) -> None:
        nonlocal first_inventory, first_sheet
        elapsed = time.perf_counter() - started
        if event.phase == "inventory_ready" and first_inventory is None:
            first_inventory = elapsed
        if event.phase == "sheet_ready" and first_sheet is None:
            first_sheet = elapsed
        if event.phase == "sheet_ready" and event.side:
            phase_seconds[event.side] = {
                "source_open": event.source_open_seconds,
                "source_read": event.source_read_seconds,
                "sidecar_write": event.sidecar_write_seconds,
                "analysis": event.analysis_seconds,
                "region_detection": event.region_detection_seconds,
                "complexity": event.complexity_seconds,
                "ranked_candidate": event.ranked_candidate_seconds,
                "ranked_candidate_calls": float(event.ranked_candidate_calls),
            }

    with tempfile.TemporaryDirectory(prefix="qc-setup-acceptance-") as directory:
        outcome = run_setup_scan_worker(
            SetupScanRequest(
                member_id="primary",
                baseline_path=str(baseline),
                current_path=str(current),
                baseline_hash=baseline_hash,
                current_hash=current_hash,
                sidecar_path=str(Path(directory) / "setup.sqlite3"),
                session_key="acceptance",
                input_generation=1,
            ),
            on_progress=progress,
        )
    elapsed = time.perf_counter() - started
    cpu_seconds = time.process_time() - cpu_started + outcome.worker_cpu_seconds
    after = (_sha256(baseline), _sha256(current))
    payload = outcome.result_payload
    members = 1 if payload is not None else 0
    sheets = 0
    regions = 0
    if payload is not None:
        current_sheets = payload.get("current_sheets")
        if isinstance(current_sheets, list):
            sheets = len(current_sheets)
            regions = sum(
                len(sheet.get("regions", []))
                for sheet in current_sheets
                if isinstance(sheet, dict)
            )
    return {
        "label": label,
        "mode": "setup",
        "ok": payload is not None,
        "error_code": "" if payload is not None else "setup_unavailable",
        "elapsed_seconds": round(elapsed, 6),
        "cpu_seconds": round(cpu_seconds, 6),
        "first_inventory_seconds": (
            round(first_inventory, 6) if first_inventory is not None else None
        ),
        "first_editable_sheet_seconds": (
            round(first_sheet, 6) if first_sheet is not None else None
        ),
        "member_count": members,
        "sheet_count": sheets,
        "region_count": regions,
        "phase_seconds_by_side": {
            side: {
                phase: round(seconds, 6)
                for phase, seconds in sorted(timings.items())
            }
            for side, timings in sorted(phase_seconds.items())
        },
        "source_hashes_unchanged": before == after,
    }


def _measure_full_preparation(
    label: str, baseline: Path, current: Path
) -> dict[str, object]:
    from qc_tool.config.profile import default_profile
    from qc_tool.excel.align import align_workbooks
    from qc_tool.io.loader import load_workbook_snapshot

    before = (_sha256(baseline), _sha256(current))
    started = time.perf_counter()
    cpu_started = time.process_time()
    baseline_snapshot = load_workbook_snapshot(
        baseline, allow_large_workbook=True
    )
    current_snapshot = load_workbook_snapshot(
        current, allow_large_workbook=True
    )
    align_workbooks(
        baseline_snapshot,
        current_snapshot,
        default_profile(),
    )
    elapsed = time.perf_counter() - started
    cpu_seconds = time.process_time() - cpu_started
    after = (_sha256(baseline), _sha256(current))
    return {
        "label": label,
        "mode": "full_preparation",
        "ok": True,
        "error_code": "",
        "elapsed_seconds": round(elapsed, 6),
        "cpu_seconds": round(cpu_seconds, 6),
        "source_hashes_unchanged": before == after,
    }


def _worker(args: argparse.Namespace) -> int:
    logging.disable(logging.CRITICAL)
    try:
        label, baseline, current = _pair(
            private_index=args.private_index,
            public_format=args.public_format,
        )
        result = (
            _measure_setup(label, baseline, current)
            if args.measure == "setup"
            else _measure_full_preparation(label, baseline, current)
        )
    except BaseException as exc:
        result = {
            "label": "unavailable",
            "mode": args.measure,
            "ok": False,
            "error_code": type(exc).__name__,
            "source_hashes_unchanged": False,
        }
    print(json.dumps(result, sort_keys=True, separators=(",", ":")))
    return 0 if result["ok"] else 1


def _process_tree_rss(process: psutil.Process) -> int:
    total = 0
    candidates = [process]
    with contextlib.suppress(psutil.NoSuchProcess, psutil.AccessDenied):
        candidates.extend(process.children(recursive=True))
    for candidate in candidates:
        try:
            total += candidate.memory_info().rss
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return total


def _supervise(arguments: list[str]) -> dict[str, object]:
    process = subprocess.Popen(
        [
            sys.executable,
            "-m",
            "scripts.setup_performance_acceptance",
            "worker",
            *arguments,
        ],
        cwd=_ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    tracked = psutil.Process(process.pid)
    peak_rss = 0
    while process.poll() is None:
        peak_rss = max(peak_rss, _process_tree_rss(tracked))
        time.sleep(0.05)
    stdout, _stderr = process.communicate()
    peak_rss = max(peak_rss, _process_tree_rss(tracked))
    try:
        payload = json.loads(stdout)
    except json.JSONDecodeError:
        payload = {
            "label": "unavailable",
            "ok": False,
            "error_code": "invalid_worker_output",
        }
    if not isinstance(payload, dict):
        payload = {
            "label": "unavailable",
            "ok": False,
            "error_code": "invalid_worker_output",
        }
    payload["peak_process_tree_rss_bytes"] = peak_rss
    payload["worker_exit_code"] = process.returncode
    return payload


def _pair_arguments(args: argparse.Namespace) -> list[str]:
    if args.public_format is not None:
        return ["--public-format", args.public_format]
    return ["--private-index", str(args.private_index)]


def _run_pair(args: argparse.Namespace) -> dict[str, object]:
    pair_arguments = _pair_arguments(args)
    setup = _supervise([*pair_arguments, "--measure", "setup"])
    full = _supervise([*pair_arguments, "--measure", "full_preparation"])
    setup_seconds = setup.get("elapsed_seconds")
    full_seconds = full.get("elapsed_seconds")
    raw_peak_rss = setup.get("peak_process_tree_rss_bytes")
    peak_rss = raw_peak_rss if isinstance(raw_peak_rss, int) else 0
    setup_faster = (
        isinstance(setup_seconds, int | float)
        and isinstance(full_seconds, int | float)
        and setup_seconds < full_seconds
    )
    first_inventory = setup.get("first_inventory_seconds")
    first_sheet = setup.get("first_editable_sheet_seconds")
    setup_ok = setup.get("ok") is True
    full_ok = full.get("ok") is True
    sources_unchanged = bool(setup.get("source_hashes_unchanged")) and bool(
        full.get("source_hashes_unchanged")
    )
    within_budget = peak_rss <= 6 * 1024**3
    if args.public_format is not None:
        acceptance_passed = (
            setup_ok
            and full_ok
            and sources_unchanged
            and within_budget
            and isinstance(first_inventory, int | float)
            and first_inventory <= 5.0
            and isinstance(first_sheet, int | float)
            and first_sheet <= 5.0
            and isinstance(setup_seconds, int | float)
            and setup_seconds <= 8.0
        )
    else:
        acceptance_passed = (
            setup_ok
            and full_ok
            and sources_unchanged
            and within_budget
            and isinstance(first_inventory, int | float)
            and first_inventory <= 15.0
            and isinstance(first_sheet, int | float)
            and setup_faster
        )
    return {
        "schema_version": 1,
        "label": setup.get("label", full.get("label", "unavailable")),
        "setup": setup,
        "full_preparation": full,
        "setup_faster_than_full_preparation": setup_faster,
        "setup_within_worker_budget": within_budget,
        "sources_unchanged": sources_unchanged,
        "acceptance_passed": acceptance_passed,
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    commands = parser.add_subparsers(dest="command", required=True)
    list_command = commands.add_parser("list")
    list_command.set_defaults(handler=lambda _args: _list_pairs())
    run = commands.add_parser("run")
    selector = run.add_mutually_exclusive_group(required=True)
    selector.add_argument("--private-index", type=int)
    selector.add_argument("--public-format", choices=("xlsx", "xlsb"))
    run.set_defaults(handler=_run_command)
    worker = commands.add_parser("worker")
    worker_selector = worker.add_mutually_exclusive_group(required=True)
    worker_selector.add_argument("--private-index", type=int)
    worker_selector.add_argument("--public-format", choices=("xlsx", "xlsb"))
    worker.add_argument("--measure", choices=("setup", "full_preparation"), required=True)
    worker.set_defaults(handler=_worker)
    return parser


def _list_pairs() -> int:
    pairs = _discover_private_pairs()
    payload = {
        "count": len(pairs),
        "pairs": [
            {
                "index": index,
                "label": label,
                "formats": [baseline.suffix.casefold(), current.suffix.casefold()],
                "total_bytes": baseline.stat().st_size + current.stat().st_size,
            }
            for index, (label, baseline, current) in enumerate(pairs)
        ],
    }
    print(json.dumps(payload, sort_keys=True, separators=(",", ":")))
    return 0


def _run_command(args: argparse.Namespace) -> int:
    result = _run_pair(args)
    print(json.dumps(result, sort_keys=True, separators=(",", ":")))
    return 0 if result["acceptance_passed"] else 1


def main() -> int:
    args = _parser().parse_args()
    return int(args.handler(args))


if __name__ == "__main__":
    raise SystemExit(main())
