"""Disposable, memory-bounded owned worker for the setup-analysis scan
(plan-20260913, Step 6).

Mirrors this project's established owned-worker pattern
(``qc_tool.excel.population_excerpt_worker``, itself mirroring
``qc_tool.worker``): a short-lived child process does the actual file
loading and structural analysis, so the hosting server process never holds
a full workbook snapshot for the duration of a scan. Primitive-only IPC (one
dict on the result queue); a wall-clock timeout plus forced
terminate/kill if the child does not finish in time.

The result crosses the process boundary as a plain, JSON-safe
``dict[str, object]`` (via ``dataclasses.asdict()`` on the pure
``SetupAnalysisResult`` tree ``qc_tool.setup.analysis`` produces) --
matching this project's own established convention for a structured but
primitive-only payload (e.g. ``RunStateRecord.action_required``), rather
than reconstructing typed objects back in the parent process.
"""

from __future__ import annotations

import dataclasses
import json
import multiprocessing as mp
import queue as queue_module
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

#: Generous but bounded: a real structural scan of a large workbook takes
#: tens of seconds (matching this project's own measured load times),
#: never minutes -- if it does, something is wrong and the scan degrades.
DEFAULT_TIMEOUT_SECONDS = 90.0
#: How long to wait for a terminated/killed child to actually exit.
_JOIN_TIMEOUT_SECONDS = 5.0
#: Poll granularity while waiting for the child so an external cancel
#: (plan-20260913, Step 7) is honored promptly instead of only at timeout.
_POLL_INTERVAL_SECONDS = 0.15


@dataclass(frozen=True, slots=True)
class SetupScanRequest:
    """Every input the child process needs, as primitives only."""

    member_id: str
    baseline_path: str
    current_path: str
    baseline_hash: str
    current_hash: str
    #: Role ("baseline"/"current") -> password; used only inside the child
    #: process, never returned, never persisted.
    passwords: dict[str, str] = field(default_factory=dict)
    allow_manual_review: bool = False


@dataclass(slots=True)
class SetupScanOutcome:
    """The scan's plain, JSON-safe result payload, or a reason it could not
    complete. ``result_payload`` is the ``dataclasses.asdict()`` form of a
    ``qc_tool.setup.models.SetupAnalysisResult`` for exactly one member.
    """

    result_payload: dict[str, object] | None = None
    #: Roles ("baseline"/"current") that need a password the caller did not
    #: supply -- never a path or filename.
    missing_credential_roles: tuple[str, ...] = ()
    disclosure: str = ""


def _worker_entry(payload: dict[str, Any], result_queue: mp.Queue[dict[str, Any]]) -> None:
    """Child-process entry point. Never touches UI/server state; the only
    channel back to the parent is one primitive-only dict on `result_queue`.
    """
    try:
        from qc_tool.history.store import sha256_file
        from qc_tool.io.decrypt import (
            InvalidPasswordError,
            PasswordRequiredError,
            open_decrypted,
        )
        from qc_tool.io.loader import (
            OOXMLWorkloadError,
            UnsupportedFormatError,
            XLSBWorkloadError,
            load_workbook_snapshot,
        )
        from qc_tool.io.xlsb_formula import XlsbFormulaScanError, scan_xlsb_formulas
        from qc_tool.setup.analysis import analyze_member
        from qc_tool.setup.models import XlsbRiskProfile

        passwords: dict[str, str] = payload["passwords"]
        snapshots: dict[str, Any] = {}
        xlsb_risks: dict[str, XlsbRiskProfile | None] = {"baseline": None, "current": None}
        missing_credentials: list[str] = []
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
                    {"ok": False, "disclosure": f"{side} source changed since it was uploaded"}
                )
                return
            password = passwords.get(side)
            try:
                snapshots[side] = load_workbook_snapshot(path, password=password)
            except (PasswordRequiredError, InvalidPasswordError):
                missing_credentials.append(side)
                continue
            except (OOXMLWorkloadError, XLSBWorkloadError):
                result_queue.put(
                    {"ok": False, "disclosure": f"{side} source is too large to scan"}
                )
                return
            except UnsupportedFormatError:
                result_queue.put(
                    {"ok": False, "disclosure": f"{side} source format is not supported"}
                )
                return
            if path.suffix.lower() == ".xlsb":
                try:
                    stream = open_decrypted(path, password)
                    scan = scan_xlsb_formulas(stream.getvalue())
                    xlsb_risks[side] = XlsbRiskProfile(
                        risky_features=scan.risky_features,
                        passive_features=scan.passive_features,
                        blocking_features=scan.blocking_features,
                        unknown_external_features=scan.unknown_external_features,
                        safe_for_external_engine=scan.safe_for_external_engine,
                    )
                except (
                    PasswordRequiredError,
                    InvalidPasswordError,
                    XlsbFormulaScanError,
                ):
                    xlsb_risks[side] = None
        if missing_credentials:
            result_queue.put(
                {"ok": False, "missing_credential_roles": missing_credentials}
            )
            return
        member_profile = analyze_member(
            member_id=payload["member_id"],
            baseline=snapshots["baseline"],
            current=snapshots["current"],
            baseline_hash=payload["baseline_hash"],
            current_hash=payload["current_hash"],
            baseline_xlsb_risk=xlsb_risks["baseline"],
            current_xlsb_risk=xlsb_risks["current"],
            allow_manual_review=bool(payload["allow_manual_review"]),
        )
        # `asdict()` preserves tuple-ness; round-trip through JSON so the
        # payload is uniformly list/dict/primitive regardless of transport
        # (multiprocessing pickling here, SQLite JSON once persisted).
        json_safe_payload = json.loads(json.dumps(dataclasses.asdict(member_profile)))
        result_queue.put({"ok": True, "result_payload": json_safe_payload})
    except BaseException as exc:  # the child must never crash silently
        result_queue.put(
            {"ok": False, "disclosure": f"setup scan worker failed ({type(exc).__name__})"}
        )


def run_setup_scan_worker(
    request: SetupScanRequest,
    *,
    timeout_seconds: float = DEFAULT_TIMEOUT_SECONDS,
    cancel_event: threading.Event | None = None,
) -> SetupScanOutcome:
    """Spawn a disposable child process to run the setup-analysis scan.

    Blocking; callers on an event loop must run this via a thread/executor
    (mirrors ``run_population_excerpt_worker``'s own contract). Callers are
    responsible for holding the shared exclusive slot
    (``qc_tool.runqueue.run_exclusive``/``ExclusiveWorkSlot``) for the
    duration of this call -- this function does not acquire it itself.

    ``cancel_event``, when supplied and set before the child finishes,
    escalates through the same terminate/kill ladder the timeout path uses
    and returns a cancellation disclosure instead of waiting out the full
    timeout (plan-20260913, Step 7's "auto-start/cancel" requirement).
    """
    ctx = mp.get_context("spawn")
    result_queue: mp.Queue[dict[str, Any]] = ctx.Queue(maxsize=1)
    payload: dict[str, Any] = {
        "member_id": request.member_id,
        "baseline_path": request.baseline_path,
        "current_path": request.current_path,
        "baseline_hash": request.baseline_hash,
        "current_hash": request.current_hash,
        "passwords": dict(request.passwords),
        "allow_manual_review": request.allow_manual_review,
    }
    process = ctx.Process(target=_worker_entry, args=(payload, result_queue), daemon=True)
    process.start()
    raw: dict[str, Any] | None = None
    cancelled = False
    elapsed = 0.0
    try:
        while True:
            if cancel_event is not None and cancel_event.is_set():
                cancelled = True
                break
            step = min(_POLL_INTERVAL_SECONDS, max(0.0, timeout_seconds - elapsed))
            try:
                raw = result_queue.get(timeout=step)
                break
            except queue_module.Empty:
                elapsed += step
                if elapsed >= timeout_seconds:
                    break
    finally:
        process.join(timeout=_JOIN_TIMEOUT_SECONDS)
        if process.is_alive():
            process.terminate()
            process.join(timeout=_JOIN_TIMEOUT_SECONDS)
            if process.is_alive():
                process.kill()
                process.join()
        result_queue.close()
    if cancelled:
        return SetupScanOutcome(disclosure="setup scan cancelled")
    if raw is None:
        return SetupScanOutcome(disclosure="setup scan worker timed out")
    if not raw.get("ok"):
        missing = raw.get("missing_credential_roles")
        if missing:
            return SetupScanOutcome(missing_credential_roles=tuple(missing))
        return SetupScanOutcome(disclosure=str(raw.get("disclosure") or "setup scan failed"))
    payload_result = raw.get("result_payload")
    return SetupScanOutcome(
        result_payload=payload_result if isinstance(payload_result, dict) else None
    )


#: Strict bounds on one preview-window fetch -- a large workbook must never
#: turn a "peek" into an unbounded materialization.
MAX_PREVIEW_ROWS = 200
MAX_PREVIEW_COLS = 40
#: A revealed formula's text is bounded per cell (matches this project's own
#: "content-free/bounded" disclosure convention elsewhere).
MAX_FORMULA_TEXT_CHARS = 256


@dataclass(frozen=True, slots=True)
class PreviewWindowRequest:
    """Every input the child process needs to fetch ONE bounded grid window
    from ONE side (baseline/current) of one already-uploaded source.
    """

    side: str  # "baseline" | "current"
    path: str
    source_hash: str
    sheet: str
    min_row: int = 1
    min_col: int = 1
    max_row: int = 40
    max_col: int = 12
    password: str = ""
    #: Formula text is opt-in ("reveal") -- absent by default; only the
    #: boolean formula-presence grid is returned otherwise.
    reveal_formulas: bool = False


@dataclass(slots=True)
class PreviewWindowOutcome:
    """A bounded grid of display strings plus a parallel formula-presence
    grid, or a plain reason none was built. Never a raw ``CellValue``, a
    filesystem path, or (unless explicitly revealed) formula text.
    """

    rows: list[list[str]] = field(default_factory=list)
    formula_cells: list[list[bool]] = field(default_factory=list)
    #: ``"r,c"`` (1-based, absolute) -> bounded formula text; populated only
    #: when the request set ``reveal_formulas``.
    formula_text: dict[str, str] = field(default_factory=dict)
    resolved_min_row: int = 1
    resolved_min_col: int = 1
    resolved_max_row: int = 0
    resolved_max_col: int = 0
    missing_credential: bool = False
    disclosure: str = ""


def _preview_worker_entry(
    payload: dict[str, Any], result_queue: mp.Queue[dict[str, Any]]
) -> None:
    try:
        from qc_tool.history.store import sha256_file
        from qc_tool.io.decrypt import InvalidPasswordError, PasswordRequiredError
        from qc_tool.io.loader import (
            OOXMLWorkloadError,
            UnsupportedFormatError,
            XLSBWorkloadError,
            load_workbook_snapshot,
        )
        from qc_tool.io.model import display_cell_value

        path = Path(str(payload["path"]))
        if not path.exists():
            result_queue.put(
                {"ok": False, "disclosure": "source is no longer at its recorded location"}
            )
            return
        expected_hash = str(payload["source_hash"])
        if expected_hash and sha256_file(path) != expected_hash:
            result_queue.put({"ok": False, "disclosure": "source changed since it was uploaded"})
            return
        password = str(payload.get("password") or "") or None
        try:
            snapshot = load_workbook_snapshot(path, password=password)
        except (PasswordRequiredError, InvalidPasswordError):
            result_queue.put({"ok": False, "missing_credential": True})
            return
        except (OOXMLWorkloadError, XLSBWorkloadError):
            result_queue.put({"ok": False, "disclosure": "source is too large to preview"})
            return
        except UnsupportedFormatError:
            result_queue.put({"ok": False, "disclosure": "source format is not supported"})
            return
        sheet_name = str(payload["sheet"])
        try:
            sheet = snapshot.sheet(sheet_name)
        except KeyError:
            result_queue.put({"ok": False, "disclosure": f"sheet not found: {sheet_name!r}"})
            return
        min_row = max(1, int(payload["min_row"]))
        min_col = max(1, int(payload["min_col"]))
        max_row = min(
            int(payload["max_row"]), min_row + MAX_PREVIEW_ROWS - 1, sheet.max_row or min_row
        )
        max_col = min(
            int(payload["max_col"]), min_col + MAX_PREVIEW_COLS - 1, sheet.max_column or min_col
        )
        max_row = max(max_row, min_row)
        max_col = max(max_col, min_col)
        reveal_formulas = bool(payload["reveal_formulas"])
        rows: list[list[str]] = []
        formula_cells: list[list[bool]] = []
        formula_text: dict[str, str] = {}
        for row in range(min_row, max_row + 1):
            value_row: list[str] = []
            formula_row: list[bool] = []
            for col in range(min_col, max_col + 1):
                record = sheet.cells.get((row, col))
                if record is None:
                    value_row.append("")
                    formula_row.append(False)
                    continue
                value_row.append(display_cell_value(record.value))
                has_formula = record.has_formula
                formula_row.append(has_formula)
                if reveal_formulas and has_formula and record.formula:
                    formula_text[f"{row},{col}"] = record.formula[:MAX_FORMULA_TEXT_CHARS]
            rows.append(value_row)
            formula_cells.append(formula_row)
        result_queue.put(
            {
                "ok": True,
                "rows": rows,
                "formula_cells": formula_cells,
                "formula_text": formula_text,
                "resolved_min_row": min_row,
                "resolved_min_col": min_col,
                "resolved_max_row": max_row,
                "resolved_max_col": max_col,
            }
        )
    except BaseException as exc:  # the child must never crash silently
        result_queue.put(
            {"ok": False, "disclosure": f"preview worker failed ({type(exc).__name__})"}
        )


def run_preview_window_worker(
    request: PreviewWindowRequest,
    *,
    timeout_seconds: float = DEFAULT_TIMEOUT_SECONDS,
) -> PreviewWindowOutcome:
    """Spawn a disposable child process to fetch one bounded grid window.

    Blocking; callers on an event loop must run this via a thread/executor.
    Loads only the ONE requested side, never both -- a preview toggle asks
    for one side at a time. Callers are responsible for holding the shared
    exclusive slot for the duration of this call.
    """
    ctx = mp.get_context("spawn")
    result_queue: mp.Queue[dict[str, Any]] = ctx.Queue(maxsize=1)
    payload: dict[str, Any] = {
        "path": request.path,
        "source_hash": request.source_hash,
        "sheet": request.sheet,
        "min_row": request.min_row,
        "min_col": request.min_col,
        "max_row": request.max_row,
        "max_col": request.max_col,
        "password": request.password,
        "reveal_formulas": request.reveal_formulas,
    }
    process = ctx.Process(
        target=_preview_worker_entry, args=(payload, result_queue), daemon=True
    )
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
        return PreviewWindowOutcome(disclosure="preview worker timed out")
    if not raw.get("ok"):
        if raw.get("missing_credential"):
            return PreviewWindowOutcome(missing_credential=True)
        return PreviewWindowOutcome(disclosure=str(raw.get("disclosure") or "preview failed"))
    return PreviewWindowOutcome(
        rows=list(raw.get("rows") or []),
        formula_cells=list(raw.get("formula_cells") or []),
        formula_text=dict(raw.get("formula_text") or {}),
        resolved_min_row=int(raw.get("resolved_min_row", 1)),
        resolved_min_col=int(raw.get("resolved_min_col", 1)),
        resolved_max_row=int(raw.get("resolved_max_row", 0)),
        resolved_max_col=int(raw.get("resolved_max_col", 0)),
    )
