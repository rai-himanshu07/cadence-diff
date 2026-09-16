"""Bounded, disposable, on-demand key-overlap query worker for confirmed
identity columns (plan-20260913, Step 12 Fix 5).

Mirrors ``qc_tool.setup.preview_worker``'s ``run_preview_window_worker``
pattern: a short-lived child process queries setup sidecar blocks and
computes ONLY the overlap ratio for the analyst's CONFIRMED
``identity_columns`` (never the auto-detector's own candidate --
``qc_tool.excel.ranked_identity._evaluate_combination()`` structurally
cannot report a low-overlap result, since it returns ``None`` the moment
either threshold misses, before ever constructing a candidate), and
returns a bounded ratio plus counts. Never returns a raw cell value or key.

Uses the SAME identity-key normalization the real alignment engine applies
to a confirmed region (exact typed equality, optional outer-whitespace
trim -- see ``qc_tool.excel.align._identity_key_component``, consumed via
``qc_tool.config.execution.region_as_row_identity_rule``'s
``exact_typed_equality=True`` convention), so the disclosed ratio matches
what will actually run at QC time, not the auto-detector's looser
casefold normalization.
"""

from __future__ import annotations

import multiprocessing as mp
import queue as queue_module
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

#: A bounded diagnostic query, not a full scan -- generous but bounded.
DEFAULT_TIMEOUT_SECONDS = 60.0
_JOIN_TIMEOUT_SECONDS = 5.0
#: A confirmed identity check is bounded to a handful of columns (mirrors
#: the auto-detector's own ``MAX_COMPOSITE_SIZE``); more than this signals
#: a caller bug, not a legitimate wide composite key.
MAX_IDENTITY_COLUMNS = 8


@dataclass(frozen=True, slots=True)
class KeyOverlapRequest:
    """Every input the child process needs to compute one region's
    confirmed-identity overlap ratio. Row bounds are resolved by the
    caller (the workspace already knows each side's resolved range); this
    worker never guesses a header boundary.
    """

    sidecar_path: str
    session_key: str
    input_generation: int
    member_id: str
    baseline_hash: str
    current_hash: str
    baseline_sheet: str
    current_sheet: str
    baseline_first_row: int
    baseline_last_row: int
    current_first_row: int
    current_last_row: int
    #: Current-side identity column letters.
    identity_columns: tuple[str, ...]
    #: Baseline-side letters, positionally corresponding to
    #: ``identity_columns``. Empty means "same letters as current" (today's
    #: default, unchanged for every region with no baseline-letter
    #: override -- Step 12 Fix 3).
    baseline_identity_columns: tuple[str, ...] = ()
    trim_identity_whitespace: bool = False


@dataclass(slots=True)
class KeyOverlapOutcome:
    """Bounded evidence only -- never a raw cell value or composite key."""

    ratio: float | None = None
    baseline_unique_keys: int = 0
    current_unique_keys: int = 0
    overlapping_keys: int = 0
    missing_credential_roles: tuple[str, ...] = field(default_factory=tuple)
    disclosure: str = ""

    @property
    def ok(self) -> bool:
        return self.ratio is not None


def _key_component(value: object, *, trim_identity_whitespace: bool) -> object:
    if isinstance(value, str):
        return value.strip() if trim_identity_whitespace else value
    return value


def _is_blank(value: object) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _unique_row_keys(
    cells: dict[tuple[int, int], Any],
    columns: list[int],
    *,
    first_row: int,
    last_row: int,
    trim_identity_whitespace: bool,
) -> set[tuple[object, ...]]:
    """Composite identity keys with count == 1 on this one side -- mirrors
    ``qc_tool.excel.ranked_identity._evaluate_combination``'s own
    unique-key filter, applied to the analyst's CONFIRMED columns instead
    of a detector-discovered candidate.
    """
    keys: list[tuple[object, ...]] = []
    for row in range(first_row, last_row + 1):
        parts: list[object] = []
        blank = False
        for column in columns:
            cell = cells.get((row, column))
            value = None if cell is None else cell.value
            if _is_blank(value):
                blank = True
                break
            parts.append(_key_component(value, trim_identity_whitespace=trim_identity_whitespace))
        if not blank:
            keys.append(tuple(parts))
    counts = Counter(keys)
    return {key for key, count in counts.items() if count == 1}


def _worker_entry(payload: dict[str, Any], result_queue: mp.Queue[dict[str, Any]]) -> None:
    """Child-process entry point. Never touches UI/server state; the only
    channel back to the parent is one primitive-only dict on `result_queue`.
    """
    try:
        from openpyxl.utils import column_index_from_string

        from qc_tool.setup.preview_store import (
            SetupScanStore,
            SetupSidecarCorruptError,
            SetupSidecarStaleError,
        )

        current_letters: list[str] = list(payload["identity_columns"])
        baseline_letters: list[str] = list(payload["baseline_identity_columns"]) or current_letters
        if len(baseline_letters) != len(current_letters):
            result_queue.put(
                {"ok": False, "disclosure": "identity column counts do not match between sides"}
            )
            return
        try:
            current_cols = [column_index_from_string(letter) for letter in current_letters]
            baseline_cols = [column_index_from_string(letter) for letter in baseline_letters]
        except ValueError:
            result_queue.put({"ok": False, "disclosure": "identity column letters are invalid"})
            return

        store = SetupScanStore(Path(str(payload["sidecar_path"])))
        try:
            baseline_cells = store.query_cells(
                str(payload["session_key"]),
                str(payload["member_id"]),
                "baseline",
                str(payload["baseline_sheet"]),
                expected_generation=int(payload["input_generation"]),
                expected_source_hash=str(payload["baseline_hash"]),
                min_row=int(payload["baseline_first_row"]),
                max_row=int(payload["baseline_last_row"]),
                columns=frozenset(baseline_cols),
            )
            current_cells = store.query_cells(
                str(payload["session_key"]),
                str(payload["member_id"]),
                "current",
                str(payload["current_sheet"]),
                expected_generation=int(payload["input_generation"]),
                expected_source_hash=str(payload["current_hash"]),
                min_row=int(payload["current_first_row"]),
                max_row=int(payload["current_last_row"]),
                columns=frozenset(current_cols),
            )
        except SetupSidecarStaleError:
            result_queue.put(
                {"ok": False, "disclosure": "setup key query is stale; rerun analysis"}
            )
            return
        except SetupSidecarCorruptError:
            result_queue.put(
                {"ok": False, "disclosure": "setup key sidecar is unreadable"}
            )
            return
        except KeyError:
            result_queue.put(
                {"ok": False, "disclosure": "sheet not found on one or both sides"}
            )
            return

        trim = bool(payload["trim_identity_whitespace"])
        base_unique = _unique_row_keys(
            baseline_cells,
            baseline_cols,
            first_row=int(payload["baseline_first_row"]),
            last_row=int(payload["baseline_last_row"]),
            trim_identity_whitespace=trim,
        )
        curr_unique = _unique_row_keys(
            current_cells,
            current_cols,
            first_row=int(payload["current_first_row"]),
            last_row=int(payload["current_last_row"]),
            trim_identity_whitespace=trim,
        )
        union = base_unique | curr_unique
        overlap = base_unique & curr_unique
        ratio = (len(overlap) / len(union)) if union else 0.0
        result_queue.put(
            {
                "ok": True,
                "ratio": ratio,
                "baseline_unique_keys": len(base_unique),
                "current_unique_keys": len(curr_unique),
                "overlapping_keys": len(overlap),
            }
        )
    except BaseException as exc:  # the child must never crash silently
        result_queue.put(
            {"ok": False, "disclosure": f"key overlap worker failed ({type(exc).__name__})"}
        )


def run_key_overlap_worker(
    request: KeyOverlapRequest,
    *,
    timeout_seconds: float = DEFAULT_TIMEOUT_SECONDS,
) -> KeyOverlapOutcome:
    """Spawn a disposable child process to compute one region's confirmed-
    identity key-overlap ratio.

    Blocking; callers on an event loop must run this via a thread/executor
    (mirrors ``run_preview_window_worker``'s own contract). Callers are
    responsible for holding the shared exclusive slot
    (``qc_tool.runqueue.run_exclusive``) for the duration of this call --
    this function does not acquire it itself.
    """
    if not request.identity_columns:
        return KeyOverlapOutcome(disclosure="no identity columns to query")
    if len(request.identity_columns) > MAX_IDENTITY_COLUMNS:
        return KeyOverlapOutcome(disclosure="too many identity columns for a bounded query")
    ctx = mp.get_context("spawn")
    result_queue: mp.Queue[dict[str, Any]] = ctx.Queue(maxsize=1)
    payload: dict[str, Any] = {
        "sidecar_path": request.sidecar_path,
        "session_key": request.session_key,
        "input_generation": request.input_generation,
        "member_id": request.member_id,
        "baseline_hash": request.baseline_hash,
        "current_hash": request.current_hash,
        "baseline_sheet": request.baseline_sheet,
        "current_sheet": request.current_sheet,
        "baseline_first_row": request.baseline_first_row,
        "baseline_last_row": request.baseline_last_row,
        "current_first_row": request.current_first_row,
        "current_last_row": request.current_last_row,
        "identity_columns": list(request.identity_columns),
        "baseline_identity_columns": list(request.baseline_identity_columns),
        "trim_identity_whitespace": request.trim_identity_whitespace,
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
        return KeyOverlapOutcome(disclosure="key overlap worker timed out")
    if not raw.get("ok"):
        missing = raw.get("missing_credential_roles")
        if missing:
            return KeyOverlapOutcome(missing_credential_roles=tuple(missing))
        disclosure = str(raw.get("disclosure") or "key overlap query failed")
        return KeyOverlapOutcome(disclosure=disclosure)
    return KeyOverlapOutcome(
        ratio=float(raw["ratio"]),
        baseline_unique_keys=int(raw["baseline_unique_keys"]),
        current_unique_keys=int(raw["current_unique_keys"]),
        overlapping_keys=int(raw["overlapping_keys"]),
    )
