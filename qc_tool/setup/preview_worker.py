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
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

#: Generous but bounded: a real structural scan of a large workbook takes
#: tens of seconds (matching this project's own measured load times),
#: never minutes -- if it does, something is wrong and the scan degrades.
DEFAULT_TIMEOUT_SECONDS = 90.0
#: How long to wait for a terminated/killed child to actually exit.
_JOIN_TIMEOUT_SECONDS = 5.0


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
) -> SetupScanOutcome:
    """Spawn a disposable child process to run the setup-analysis scan.

    Blocking; callers on an event loop must run this via a thread/executor
    (mirrors ``run_population_excerpt_worker``'s own contract). Callers are
    responsible for holding the shared exclusive slot
    (``qc_tool.runqueue.run_exclusive``/``ExclusiveWorkSlot``) for the
    duration of this call -- this function does not acquire it itself.
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
