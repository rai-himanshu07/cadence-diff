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
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

#: Bounded post-scan query timeout. Setup scans use an inactivity-aware
#: coordinator and therefore have no fixed total-wall timeout.
DEFAULT_TIMEOUT_SECONDS = 90.0
#: How long to wait for a terminated/killed child to actually exit.
_JOIN_TIMEOUT_SECONDS = 5.0
#: Poll granularity while waiting for the child so an external cancel
#: (plan-20260913, Step 7) is honored promptly instead of only at timeout.
_POLL_INTERVAL_SECONDS = 0.15
_PROGRESS_HEARTBEAT_SECONDS = 2.0
DEFAULT_INACTIVITY_TIMEOUT_SECONDS = 300.0


@dataclass(frozen=True, slots=True)
class SetupScanRequest:
    """Every input the child process needs, as primitives only."""

    member_id: str
    baseline_path: str
    current_path: str
    baseline_hash: str
    current_hash: str
    sidecar_path: str
    session_key: str
    input_generation: int
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
    worker_cpu_seconds: float = 0.0


@dataclass(frozen=True, slots=True)
class SetupScanProgress:
    """Aggregate-only progress; sheet names and source content stay in the sidecar."""

    phase: str
    side: str = ""
    processed: int = 0
    total: int = 0
    source_open_seconds: float = 0.0
    source_read_seconds: float = 0.0
    sidecar_write_seconds: float = 0.0
    analysis_seconds: float = 0.0
    region_detection_seconds: float = 0.0
    complexity_seconds: float = 0.0
    ranked_candidate_seconds: float = 0.0
    ranked_candidate_calls: int = 0
    profile_payload: dict[str, object] | None = field(
        default=None, repr=False, compare=False
    )


def _worker_entry(payload: dict[str, Any], result_queue: mp.Queue[dict[str, Any]]) -> None:
    """Child-process entry point. Never touches UI/server state; the only
    channel back to the parent is one primitive-only dict on `result_queue`.
    """
    cpu_started = time.process_time()

    def put_terminal(message: dict[str, Any]) -> None:
        message["worker_cpu_seconds"] = time.process_time() - cpu_started
        result_queue.put(message)

    try:
        from qc_tool.history.store import sha256_file
        from qc_tool.io.decrypt import (
            InvalidPasswordError,
            PasswordRequiredError,
        )
        from qc_tool.setup.analysis import StreamingMemberAnalyzer
        from qc_tool.setup.models import SheetSetupProfile, XlsbRiskProfile
        from qc_tool.setup.preview_store import SetupScanStore
        from qc_tool.setup.source_reader import (
            UnsupportedSetupSourceError,
            open_setup_source,
        )

        passwords: dict[str, str] = payload["passwords"]
        store = SetupScanStore(Path(str(payload["sidecar_path"])))
        session_key = str(payload["session_key"])
        member_id = str(payload["member_id"])
        input_generation = int(payload["input_generation"])
        store.delete_member(
            session_key,
            member_id,
            input_generation=input_generation,
        )
        xlsb_risks: dict[str, XlsbRiskProfile | None] = {
            "baseline": None,
            "current": None,
        }
        analyzer = StreamingMemberAnalyzer(
            member_id=member_id,
            baseline_hash=str(payload["baseline_hash"]),
            current_hash=str(payload["current_hash"]),
            allow_manual_review=bool(payload["allow_manual_review"]),
        )
        missing_credentials: list[str] = []
        source_specs = (
            ("baseline", payload["baseline_path"], payload["baseline_hash"]),
            ("current", payload["current_path"], payload["current_hash"]),
        )
        grouped_sources: dict[str, list[tuple[str, Path, str]]] = {}
        for side, path_str, expected_hash in source_specs:
            path = Path(path_str)
            if not path.exists():
                put_terminal(
                    {
                        "ok": False,
                        "disclosure": f"{side} source is no longer at its recorded location",
                    }
                )
                return
            if expected_hash and sha256_file(path) != expected_hash:
                put_terminal(
                    {"ok": False, "disclosure": f"{side} source changed since it was uploaded"}
                )
                return
            source_key = str(expected_hash) or str(path.resolve())
            grouped_sources.setdefault(source_key, []).append(
                (side, path, str(expected_hash))
            )

        for grouped in grouped_sources.values():
            sides = tuple(item[0] for item in grouped)
            path = grouped[0][1]
            password = next(
                (passwords[side] for side in sides if passwords.get(side)),
                None,
            )
            try:
                source_open_started = time.perf_counter()
                with open_setup_source(path, password=password) as source:
                    source_open_seconds = time.perf_counter() - source_open_started
                    for side in sides:
                        result_queue.put(
                            {
                                "kind": "progress",
                                "phase": "inventory_ready",
                                "side": side,
                                "processed": 0,
                                "total": source.sheet_count,
                            }
                        )
                    source_read_seconds = 0.0
                    sidecar_write_seconds = 0.0
                    analysis_seconds = 0.0
                    next_sheet_started = time.perf_counter()
                    for sheet_index, sheet in enumerate(source.iter_sheets(), start=1):
                        source_read_seconds += time.perf_counter() - next_sheet_started
                        sidecar_started = time.perf_counter()
                        for side, _side_path, expected_hash in grouped:
                            store.save_sheet(
                                session_key,
                                member_id,
                                side,
                                input_generation=input_generation,
                                source_hash=expected_hash,
                                sheet=sheet,
                            )
                        sidecar_write_seconds += time.perf_counter() - sidecar_started
                        analysis_started = time.perf_counter()
                        profiles: dict[str, SheetSetupProfile] = {}
                        if "baseline" in sides:
                            profiles["baseline"] = analyzer.add_baseline(sheet)
                        if "current" in sides:
                            baseline_sheet = sheet if "baseline" in sides else None
                            if baseline_sheet is None:
                                try:
                                    baseline_sheet = store.load_sheet(
                                        session_key,
                                        member_id,
                                        "baseline",
                                        sheet.name,
                                        expected_generation=input_generation,
                                        expected_source_hash=str(
                                            payload["baseline_hash"]
                                        ),
                                    )
                                except KeyError:
                                    baseline_sheet = None
                            profiles["current"] = analyzer.add_current(
                                sheet, baseline_sheet
                            )
                        profile_payloads = {
                            side: json.loads(json.dumps(dataclasses.asdict(profile)))
                            for side, profile in profiles.items()
                        }
                        analysis_seconds += time.perf_counter() - analysis_started
                        for side, _side_path, _expected_hash in grouped:
                            is_baseline = side == "baseline"
                            result_queue.put(
                                {
                                    "kind": "progress",
                                    "phase": "sheet_ready",
                                    "side": side,
                                    "processed": sheet_index,
                                    "total": source.sheet_count,
                                    "profile_payload": profile_payloads.get(side),
                                    "source_open_seconds": source_open_seconds,
                                    "source_read_seconds": source_read_seconds,
                                    "sidecar_write_seconds": sidecar_write_seconds,
                                    "analysis_seconds": analysis_seconds,
                                    "region_detection_seconds": (
                                        analyzer.baseline_region_seconds
                                        if is_baseline
                                        else analyzer.current_region_seconds
                                    ),
                                    "complexity_seconds": (
                                        analyzer.baseline_complexity_seconds
                                        if is_baseline
                                        else analyzer.current_complexity_seconds
                                    ),
                                    "ranked_candidate_seconds": (
                                        0.0
                                        if is_baseline
                                        else analyzer.ranked_candidate_seconds
                                    ),
                                    "ranked_candidate_calls": (
                                        0
                                        if is_baseline
                                        else analyzer.ranked_candidate_calls
                                    ),
                                }
                            )
                        next_sheet_started = time.perf_counter()
                    for side in sides:
                        xlsb_risks[side] = source.xlsb_risk
            except (PasswordRequiredError, InvalidPasswordError):
                missing_credentials.extend(sides)
                continue
            except UnsupportedSetupSourceError:
                put_terminal(
                    {
                        "ok": False,
                        "disclosure": "source format is not supported for setup",
                    }
                )
                return
        if missing_credentials:
            put_terminal(
                {
                    "ok": False,
                    "missing_credential_roles": sorted(set(missing_credentials)),
                }
            )
            return
        member_profile = analyzer.result(
            baseline_xlsb_risk=xlsb_risks["baseline"],
            current_xlsb_risk=xlsb_risks["current"],
        )
        # `asdict()` preserves tuple-ness; round-trip through JSON so the
        # payload is uniformly list/dict/primitive regardless of transport
        # (multiprocessing pickling here, SQLite JSON once persisted).
        json_safe_payload = json.loads(json.dumps(dataclasses.asdict(member_profile)))
        store.save_result(
            session_key,
            member_id,
            json_safe_payload,
            input_generation=input_generation,
        )
        put_terminal(
            {"kind": "result", "ok": True, "result_payload": json_safe_payload}
        )
    except BaseException as exc:  # the child must never crash silently
        put_terminal(
            {"ok": False, "disclosure": f"setup scan worker failed ({type(exc).__name__})"}
        )


def run_setup_scan_worker(
    request: SetupScanRequest,
    *,
    timeout_seconds: float | None = None,
    inactivity_timeout_seconds: float = DEFAULT_INACTIVITY_TIMEOUT_SECONDS,
    cancel_event: threading.Event | None = None,
    on_progress: Callable[[SetupScanProgress], None] | None = None,
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
        "sidecar_path": request.sidecar_path,
        "session_key": request.session_key,
        "input_generation": request.input_generation,
        "passwords": dict(request.passwords),
        "allow_manual_review": request.allow_manual_review,
    }
    process = ctx.Process(target=_worker_entry, args=(payload, result_queue), daemon=True)
    process.start()
    raw: dict[str, Any] | None = None
    cancelled = False
    inactive = False
    elapsed = 0.0
    last_child_progress = time.monotonic()
    last_heartbeat = last_child_progress
    last_progress = SetupScanProgress(phase="starting")
    try:
        while True:
            if cancel_event is not None and cancel_event.is_set():
                cancelled = True
                break
            step = (
                _POLL_INTERVAL_SECONDS
                if timeout_seconds is None
                else min(
                    _POLL_INTERVAL_SECONDS,
                    max(0.0, timeout_seconds - elapsed),
                )
            )
            try:
                message = result_queue.get(timeout=step)
                if message.get("kind") == "progress":
                    last_child_progress = time.monotonic()
                    if on_progress is not None:
                        profile_payload = message.get("profile_payload")
                        last_progress = SetupScanProgress(
                            phase=str(message.get("phase", "")),
                            side=str(message.get("side", "")),
                            processed=int(message.get("processed", 0)),
                            total=int(message.get("total", 0)),
                            source_open_seconds=float(
                                message.get("source_open_seconds", 0.0)
                            ),
                            source_read_seconds=float(
                                message.get("source_read_seconds", 0.0)
                            ),
                            sidecar_write_seconds=float(
                                message.get("sidecar_write_seconds", 0.0)
                            ),
                            analysis_seconds=float(
                                message.get("analysis_seconds", 0.0)
                            ),
                            region_detection_seconds=float(
                                message.get("region_detection_seconds", 0.0)
                            ),
                            complexity_seconds=float(
                                message.get("complexity_seconds", 0.0)
                            ),
                            ranked_candidate_seconds=float(
                                message.get("ranked_candidate_seconds", 0.0)
                            ),
                            ranked_candidate_calls=int(
                                message.get("ranked_candidate_calls", 0)
                            ),
                            profile_payload=(
                                dict(profile_payload)
                                if isinstance(profile_payload, dict)
                                else None
                            ),
                        )
                        on_progress(last_progress)
                        last_heartbeat = last_child_progress
                    continue
                raw = message
                break
            except queue_module.Empty:
                elapsed += step
                now = time.monotonic()
                if (
                    on_progress is not None
                    and now - last_heartbeat >= _PROGRESS_HEARTBEAT_SECONDS
                ):
                    on_progress(
                        dataclasses.replace(last_progress, profile_payload=None)
                    )
                    last_heartbeat = now
                if now - last_child_progress >= inactivity_timeout_seconds:
                    inactive = True
                    break
                if timeout_seconds is not None and elapsed >= timeout_seconds:
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
    if inactive:
        return SetupScanOutcome(
            disclosure="setup scan made no progress before the inactivity limit"
        )
    if raw is None:
        return SetupScanOutcome(disclosure="setup scan worker timed out")
    worker_cpu_seconds = float(raw.get("worker_cpu_seconds", 0.0))
    if not raw.get("ok"):
        missing = raw.get("missing_credential_roles")
        if missing:
            return SetupScanOutcome(
                missing_credential_roles=tuple(missing),
                worker_cpu_seconds=worker_cpu_seconds,
            )
        return SetupScanOutcome(
            disclosure=str(raw.get("disclosure") or "setup scan failed"),
            worker_cpu_seconds=worker_cpu_seconds,
        )
    payload_result = raw.get("result_payload")
    return SetupScanOutcome(
        result_payload=payload_result if isinstance(payload_result, dict) else None,
        worker_cpu_seconds=worker_cpu_seconds,
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
    sidecar_path: str
    session_key: str
    input_generation: int
    member_id: str
    source_hash: str
    sheet: str
    min_row: int = 1
    min_col: int = 1
    max_row: int = 40
    max_col: int = 12
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
        from qc_tool.io.model import display_cell_value
        from qc_tool.setup.preview_store import (
            SetupScanStore,
            SetupSidecarCorruptError,
            SetupSidecarStaleError,
        )

        store = SetupScanStore(Path(str(payload["sidecar_path"])))
        sheet_name = str(payload["sheet"])
        try:
            inventory = store.get_sheet_inventory(
                str(payload["session_key"]),
                str(payload["member_id"]),
                str(payload["side"]),
                sheet_name,
                expected_generation=int(payload["input_generation"]),
                expected_source_hash=str(payload["source_hash"]),
            )
        except SetupSidecarStaleError:
            result_queue.put(
                {"ok": False, "disclosure": "setup preview is stale; rerun analysis"}
            )
            return
        except SetupSidecarCorruptError:
            result_queue.put(
                {"ok": False, "disclosure": "setup preview sidecar is unreadable"}
            )
            return
        except KeyError:
            result_queue.put({"ok": False, "disclosure": "sheet is unavailable in setup"})
            return
        min_row = max(1, int(payload["min_row"]))
        min_col = max(1, int(payload["min_col"]))
        max_row = min(
            int(payload["max_row"]),
            min_row + MAX_PREVIEW_ROWS - 1,
            inventory.max_row or min_row,
        )
        max_col = min(
            int(payload["max_col"]),
            min_col + MAX_PREVIEW_COLS - 1,
            inventory.max_column or min_col,
        )
        max_row = max(max_row, min_row)
        max_col = max(max_col, min_col)
        window = store.query_window(
            str(payload["session_key"]),
            str(payload["member_id"]),
            str(payload["side"]),
            sheet_name,
            expected_generation=int(payload["input_generation"]),
            expected_source_hash=str(payload["source_hash"]),
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
        )
        reveal_formulas = bool(payload["reveal_formulas"])
        rows: list[list[str]] = []
        formula_cells: list[list[bool]] = []
        formula_text: dict[str, str] = {}
        for row in range(min_row, max_row + 1):
            value_row: list[str] = []
            formula_row: list[bool] = []
            for col in range(min_col, max_col + 1):
                record = window.cells.get((row, col))
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
        "side": request.side,
        "sidecar_path": request.sidecar_path,
        "session_key": request.session_key,
        "input_generation": request.input_generation,
        "member_id": request.member_id,
        "source_hash": request.source_hash,
        "sheet": request.sheet,
        "min_row": request.min_row,
        "min_col": request.min_col,
        "max_row": request.max_row,
        "max_col": request.max_col,
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
