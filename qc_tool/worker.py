"""Import-safe entry point and IPC contract for the owned QC worker process.

A spawn context imports this module in a fresh interpreter, so it must stay
free of UI, server, and network side effects. Every message crossing the
channel is a versioned dictionary of primitives; passwords, findings, and
source content never appear in one.
"""

from __future__ import annotations

import contextlib
import multiprocessing as mp
import queue as queue_module
import time
from pathlib import Path, PurePosixPath
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:  # pragma: no cover - typing only
    from multiprocessing.queues import Queue as IPCQueue
    from multiprocessing.synchronize import Event as CancelFlag

    from qc_tool.progress import CancellationFlag

#: Version of the primitive-only envelope exchanged with the queue manager.
ENVELOPE_VERSION = 1
#: Bounded channel capacity; progress is dropped rather than blocking QC.
IPC_QUEUE_CAPACITY = 64
#: How long a worker may block delivering its single terminal message.
TERMINAL_PUT_TIMEOUT_SECONDS = 5.0
#: How often an owned worker re-checks that its owner is still alive.
PARENT_LIVENESS_INTERVAL_SECONDS = 1.0
_MAX_ERROR_CHARS = 240


def progress_message(
    phase: str, processed: int, total: int, detail: str
) -> dict[str, Any]:
    return {
        "v": ENVELOPE_VERSION,
        "kind": "progress",
        "phase": phase,
        "processed": int(processed),
        "total": int(total),
        "detail": detail,
    }


def result_message(run_id: int, phases: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "v": ENVELOPE_VERSION,
        "kind": "result",
        "run_id": int(run_id),
        "phases": phases,
    }


def cancelled_message(phases: list[dict[str, Any]]) -> dict[str, Any]:
    return {"v": ENVELOPE_VERSION, "kind": "cancelled", "phases": phases}


def error_message(error: str, phases: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "v": ENVELOPE_VERSION,
        "kind": "error",
        "error": error,
        "phases": phases,
    }


def blocked_message(
    action_required: dict[str, Any], phases: list[dict[str, Any]]
) -> dict[str, Any]:
    """A terminal, non-failure outcome: the run cannot proceed as configured.

    ``action_required`` is the bounded, primitive-only
    ``RunActionRequired.model_dump(mode="json")`` payload -- never a value,
    formula, or path.
    """
    return {
        "v": ENVELOPE_VERSION,
        "kind": "blocked",
        "action_required": action_required,
        "phases": phases,
    }


_LEADING_PUNCTUATION = "'\"([<"
_TRAILING_PUNCTUATION = "'\").,;:)]>"


def _redact(token: str) -> str:
    """Replace a path-like token with its filename, keeping any punctuation."""
    if "/" not in token and "\\" not in token:
        return token
    lead = token[: len(token) - len(token.lstrip(_LEADING_PUNCTUATION))]
    end = len(token.rstrip(_TRAILING_PUNCTUATION))
    core = token[len(lead) : end].replace("\\", "/")
    return f"{lead}{PurePosixPath(core).name or 'path'}{token[end:]}"


def sanitize_error(error: BaseException | str) -> str:
    """Reduce a failure to one bounded line with filenames instead of paths."""
    text = (
        f"{type(error).__name__}: {error}"
        if isinstance(error, BaseException)
        else str(error)
    )
    first_line = text.strip().splitlines()[0] if text.strip() else ""
    redacted = " ".join(_redact(token) for token in first_line.split(" "))
    return redacted[:_MAX_ERROR_CHARS]


def _deliver(events: IPCQueue[dict[str, Any]], message: dict[str, Any]) -> None:
    # A dropped terminal message is recorded by the manager as a failure.
    with contextlib.suppress(queue_module.Full, OSError, ValueError):
        events.put(message, timeout=TERMINAL_PUT_TIMEOUT_SECONDS)


class OwnedCancellationFlag:
    """Cancels when the manager asks, or when the owning process is gone.

    A crashed or killed owner cannot run its cleanup, so the worker has to stop
    itself; otherwise it would finish and record a run nobody is waiting for.
    """

    def __init__(
        self,
        flag: CancellationFlag,
        parent: Any = None,
        *,
        interval: float = PARENT_LIVENESS_INTERVAL_SECONDS,
    ) -> None:
        self._flag = flag
        self._parent = parent
        self._interval = interval
        self._checked_at = 0.0
        self._orphaned = False

    @property
    def orphaned(self) -> bool:
        return self._orphaned

    def set(self) -> None:
        self._flag.set()

    def is_set(self) -> bool:
        if self._flag.is_set():
            return True
        if self._orphaned:
            return True
        if self._parent is None:
            return False
        now = time.monotonic()
        if now - self._checked_at < self._interval:
            return False
        self._checked_at = now
        self._orphaned = not self._parent.is_alive()
        return self._orphaned


def worker_main(
    payload: dict[str, Any],
    credentials: dict[str, str],
    events: IPCQueue[dict[str, Any]],
    cancel_flag: CancelFlag,
) -> None:
    """Execute one QC request and report exactly one terminal message."""
    from qc_tool.config.profile import DeliverableProfile
    from qc_tool.coverage import QCRunMode
    from qc_tool.progress import (
        CancellationToken,
        PhaseTelemetry,
        ProgressEvent,
        RunCancelled,
        RunPhase,
    )
    from qc_tool.run_action import RunBlockedError
    from qc_tool.run_service import perform_run

    telemetry = PhaseTelemetry()
    last_phase = RunPhase.PREPARING
    owned_flag = OwnedCancellationFlag(cancel_flag, mp.parent_process())

    def on_progress(event: ProgressEvent) -> None:
        nonlocal last_phase
        last_phase = event.phase
        telemetry(event)
        # Bounded channel: the newest progress is dropped rather than blocking QC.
        with contextlib.suppress(queue_module.Full):
            events.put_nowait(
                progress_message(
                    event.phase.value, event.processed, event.total, event.detail
                )
            )

    try:
        # Validate optional package manifest and pass member-scoped sheets
        from qc_tool.package import PackageManifest

        raw_manifest = payload.get("package_manifest") or None
        manifest_obj = None
        if raw_manifest:
            # Validate shape (primitive-only) and convert to model
            manifest_obj = PackageManifest.model_validate(raw_manifest)

        artifacts = perform_run(
            Path(payload["work_dir"]),
            {role: Path(path) for role, path in payload["files"].items()},
            credentials,
            DeliverableProfile.model_validate(payload["profile"]),
            mode=QCRunMode(payload["mode"]),
            rerun_of=payload["rerun_of"],
            allow_large_workbooks=bool(payload["allow_large_workbooks"]),
            allow_dependency_indexing=bool(payload["allow_dependency_indexing"]),
            acceptance_absolute=float(payload.get("acceptance_absolute", 0.0)),
            acceptance_relative=float(payload.get("acceptance_relative", 0.0)),
            compare_sheets=list(payload.get("compare_sheets") or []) or None,
            compare_slides=[int(i) for i in payload.get("compare_slides") or []] or None,
            cancellation_token=CancellationToken(owned_flag),
            on_progress=on_progress,
            package_manifest=manifest_obj,
            compare_member_sheets={
                k: tuple(v or ()) for k, v in (payload.get("compare_member_sheets") or {}).items()
            },
        )
    except RunCancelled:
        message = cancelled_message(telemetry.as_payload())
    except RunBlockedError as blocked:
        message = blocked_message(
            blocked.action_required.model_dump(mode="json"), telemetry.as_payload()
        )
    except Exception as exc:  # analyst-facing failure, never an unreported crash
        error = sanitize_error(exc)
        telemetry.fail(last_phase, error)
        message = error_message(error, telemetry.as_payload())
    else:
        message = result_message(artifacts.run_id, telemetry.as_payload())
    finally:
        credentials.clear()
    if owned_flag.orphaned:
        return
    _deliver(events, message)
