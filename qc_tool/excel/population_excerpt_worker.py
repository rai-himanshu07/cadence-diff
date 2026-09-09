"""Disposable, memory-bounded owned worker for population sample excerpts.

Reopening a run's recorded sources to build a handful of sample excerpts
requires loading a full workbook snapshot -- for a large real workbook, that
is exactly the multi-GB cost this project's own worker/UI-server memory
budgets exist to bound. Doing that reopen directly inside the NiceGUI server
process (even on a background thread) means the SERVER's own process memory
carries the full snapshot for as long as the load takes, violating the
"the UI server must stay bounded" / "never loads two full large workbooks
itself" constraints.

This module instead spawns a short-lived, disposable child process (the
server's own memory is never touched by the snapshot), waits for a small,
primitive-only excerpt payload back, and enforces both a wall-clock timeout
and forced termination if the child does not finish. It mirrors this
project's established owned-worker pattern (`qc_tool/worker.py`: primitive
-only messages, explicit cancellation) at a scope proportionate to one
excerpt lookup rather than a full QC run -- no persistent queue, no
progress reporting, no orphan-liveness polling.
"""

from __future__ import annotations

import multiprocessing as mp
import queue as queue_module
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from qc_tool.findings import GridExcerpt

#: Generous but bounded: real reopens of large workbooks take tens of
#: seconds (matching this project's own measured load times), never minutes.
DEFAULT_TIMEOUT_SECONDS = 60.0
#: How long to wait for a terminated/killed child to actually exit.
_JOIN_TIMEOUT_SECONDS = 5.0


@dataclass(frozen=True, slots=True)
class PopulationExcerptRequest:
    """Every input the child process needs, as primitives only."""

    baseline_path: str
    current_path: str
    baseline_hash: str
    current_hash: str
    sheet: str
    #: (current_location, baseline_location_or_None) pairs to build excerpts for.
    samples: tuple[tuple[str, str | None], ...]


@dataclass(slots=True)
class PopulationExcerpts:
    """On-demand sample excerpts, or a plain reason none were built."""

    excerpts: dict[str, tuple[GridExcerpt | None, GridExcerpt | None]] = field(
        default_factory=dict
    )
    disclosure: str = ""


def _worker_entry(payload: dict[str, Any], result_queue: mp.Queue[dict[str, Any]]) -> None:
    """Child-process entry point. Never touches UI/server state; the only
    channel back to the parent is one primitive-only dict on `result_queue`.
    """
    try:
        from qc_tool.excel.context import build_excerpt
        from qc_tool.history.store import sha256_file
        from qc_tool.io.decrypt import InvalidPasswordError, PasswordRequiredError
        from qc_tool.io.loader import (
            OOXMLWorkloadError,
            UnsupportedFormatError,
            XLSBWorkloadError,
            load_workbook_snapshot,
        )

        snapshots: dict[str, Any] = {}
        for side, path_str, expected_hash in (
            ("baseline", payload["baseline_path"], payload["baseline_hash"]),
            ("current", payload["current_path"], payload["current_hash"]),
        ):
            path = Path(path_str)
            if not path.exists():
                result_queue.put(
                    {
                        "ok": False,
                        "disclosure": f"{side} source is no longer at its recorded location",
                    }
                )
                return
            if expected_hash and sha256_file(path) != expected_hash:
                result_queue.put(
                    {"ok": False, "disclosure": f"{side} source changed since this run recorded it"}
                )
                return
            try:
                snapshots[side] = load_workbook_snapshot(path)
            except (PasswordRequiredError, InvalidPasswordError):
                result_queue.put(
                    {
                        "ok": False,
                        "disclosure": (
                            f"{side} source is password-protected; run history "
                            "never retains passwords"
                        ),
                    }
                )
                return
            except (OOXMLWorkloadError, XLSBWorkloadError):
                result_queue.put(
                    {
                        "ok": False,
                        "disclosure": f"{side} source is too large to reopen for excerpts",
                    }
                )
                return
            except UnsupportedFormatError:
                result_queue.put(
                    {
                        "ok": False,
                        "disclosure": f"{side} source format is not supported for reopening",
                    }
                )
                return
        sheet_name = payload["sheet"]
        try:
            baseline_sheet = snapshots["baseline"].sheet(sheet_name)
            current_sheet = snapshots["current"].sheet(sheet_name)
        except KeyError:
            result_queue.put(
                {
                    "ok": False,
                    "disclosure": f"sheet {sheet_name!r} was not found in the reopened sources",
                }
            )
            return
        excerpts: dict[str, list[dict[str, Any] | None]] = {}
        for current_location, baseline_location in payload["samples"]:
            baseline_excerpt = (
                build_excerpt(baseline_sheet, baseline_location) if baseline_location else None
            )
            current_excerpt = build_excerpt(current_sheet, current_location)
            excerpts[current_location] = [
                baseline_excerpt.model_dump(mode="json") if baseline_excerpt is not None else None,
                current_excerpt.model_dump(mode="json") if current_excerpt is not None else None,
            ]
        result_queue.put({"ok": True, "excerpts": excerpts})
    except BaseException as exc:  # the child must never crash silently
        result_queue.put(
            {"ok": False, "disclosure": f"excerpt worker failed ({type(exc).__name__})"}
        )


def run_population_excerpt_worker(
    request: PopulationExcerptRequest,
    *,
    timeout_seconds: float = DEFAULT_TIMEOUT_SECONDS,
) -> PopulationExcerpts:
    """Spawn a disposable child process to build sample excerpts.

    Blocking; callers on an event loop must run this via a thread/executor
    (mirrors the pre-existing `_load_population_sample_excerpts` contract).
    """
    ctx = mp.get_context("spawn")
    result_queue: mp.Queue[dict[str, Any]] = ctx.Queue(maxsize=1)
    payload: dict[str, Any] = {
        "baseline_path": request.baseline_path,
        "current_path": request.current_path,
        "baseline_hash": request.baseline_hash,
        "current_hash": request.current_hash,
        "sheet": request.sheet,
        "samples": [list(pair) for pair in request.samples],
    }
    process = ctx.Process(target=_worker_entry, args=(payload, result_queue), daemon=True)
    process.start()
    raw: dict[str, Any] | None = None
    try:
        raw = result_queue.get(timeout=timeout_seconds)
    except queue_module.Empty:
        raw = None
    finally:
        process.join(timeout=_JOIN_TIMEOUT_SECONDS)
        if process.is_alive():
            process.terminate()
            process.join(timeout=_JOIN_TIMEOUT_SECONDS)
            if process.is_alive():
                process.kill()
                process.join()
        result_queue.close()
    if raw is None:
        return PopulationExcerpts(disclosure="excerpt worker timed out")
    if not raw.get("ok"):
        return PopulationExcerpts(disclosure=str(raw.get("disclosure") or "excerpt worker failed"))
    excerpts: dict[str, tuple[GridExcerpt | None, GridExcerpt | None]] = {}
    for current_location, pair in raw["excerpts"].items():
        baseline_payload, current_payload = pair
        excerpts[current_location] = (
            GridExcerpt.model_validate(baseline_payload) if baseline_payload is not None else None,
            GridExcerpt.model_validate(current_payload) if current_payload is not None else None,
        )
    return PopulationExcerpts(excerpts=excerpts)
