"""Process-global FIFO queue that runs QC in one owned worker process.

One heavy run executes at a time. Every transition is persisted to
`run_state`, so a browser refresh or a second tab reconnects to work already
in flight instead of starting duplicate work. The manager owns the worker
process and its channel: it cancels cooperatively first, then escalates
through terminate and kill, and always joins and closes both before the slot
is released.
"""

from __future__ import annotations

import contextlib
import logging
import multiprocessing as mp
import queue as queue_module
import secrets
import threading
import time
from collections import deque
from collections.abc import Callable
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Any

from qc_tool.history.run_state import (
    ACTIVE_STATUSES,
    RunStateRecord,
    RunStateStore,
    RunStatus,
)
from qc_tool.worker import ENVELOPE_VERSION, IPC_QUEUE_CAPACITY, sanitize_error, worker_main

if TYPE_CHECKING:  # pragma: no cover - typing only
    from multiprocessing.context import SpawnContext
    from multiprocessing.process import BaseProcess
    from multiprocessing.queues import Queue as IPCQueue
    from multiprocessing.synchronize import Event as CancelFlag

logger = logging.getLogger(__name__)

#: Cooperative cancellation grace before the terminate/kill ladder starts.
CANCEL_GRACE_SECONDS = 10.0
#: Join interval allowed after `terminate()` and again after `kill()`.
TERMINATE_JOIN_SECONDS = 5.0
_SUPERVISOR_POLL_SECONDS = 0.05
#: Time a dead worker is given to flush a terminal message already in flight.
_EXIT_SETTLE_SECONDS = 0.5


class QueueBusyError(RuntimeError):
    """A password-protected run cannot wait in the persisted queue."""


def new_request_id() -> str:
    return secrets.token_hex(8)


@dataclass(frozen=True, slots=True)
class RunRequest:
    """Primitive-only description of one QC request handed to a worker."""

    request_id: str
    work_dir: str
    mode: str
    profile_name: str
    profile: dict[str, Any]
    #: Role -> managed execution path, used only inside the owned process.
    files: dict[str, str]
    #: Role -> display filename shown in queue surfaces.
    display_files: dict[str, str]
    #: Run-level finding-output contract request (plan-20260910); a plain
    #: string mirroring `mode` ("profile"/"decision"/"atomic").
    requested_output_mode: str = "profile"
    allow_large_workbooks: bool = False
    allow_dependency_indexing: bool = False
    acceptance_absolute: float = 0.0
    acceptance_relative: float = 0.0
    #: Analyst comparison scope; None/empty = everything.
    compare_sheets: tuple[str, ...] = ()
    compare_slides: tuple[int, ...] = ()
    rerun_of: int | None = None
    # New fields for multi-workbook intake (primitive-only payloads)
    package_manifest: dict[str, Any] = field(default_factory=dict)
    compare_member_sheets: dict[str, tuple[str, ...]] = field(default_factory=dict)


@dataclass(slots=True)
class _ActiveJob:
    request: RunRequest
    process: BaseProcess
    events: IPCQueue[dict[str, Any]]
    cancel_flag: CancelFlag
    terminal_status: RunStatus | None = None
    run_id: int | None = None
    error: str = ""
    action_required: dict[str, Any] | None = None
    phases: list[dict[str, Any]] = field(default_factory=list)
    cancel_deadline: float | None = None
    terminated: bool = False
    killed: bool = False
    dead_since: float | None = None


class RunQueueManager:
    """Single-worker FIFO manager for one managed data directory."""

    def __init__(
        self,
        work_dir: Path,
        *,
        store: RunStateStore | None = None,
        context: SpawnContext | None = None,
        worker: Callable[..., None] = worker_main,
        cancel_grace_seconds: float = CANCEL_GRACE_SECONDS,
        terminate_join_seconds: float = TERMINATE_JOIN_SECONDS,
    ) -> None:
        self.work_dir = work_dir
        self.store = store or RunStateStore(work_dir / "history.sqlite3")
        self._context = context or mp.get_context("spawn")
        self._worker = worker
        self._cancel_grace = cancel_grace_seconds
        self._terminate_join = terminate_join_seconds
        self._lock = threading.RLock()
        self._wake = threading.Event()
        self._stopping = threading.Event()
        self._pending: deque[tuple[RunRequest, dict[str, str]]] = deque()
        self._active: _ActiveJob | None = None
        #: Request that owns the slot while its process is being spawned.
        self._starting_id: str | None = None
        self._cancel_requested: set[str] = set()
        self._thread: threading.Thread | None = None
        #: Set once the hosting server registers its shutdown hook.
        self.shutdown_hook_installed = False
        # A fresh manager means a fresh server process: nothing is resumed.
        self.store.mark_orphaned()

    @property
    def closed(self) -> bool:
        return self._stopping.is_set()

    def _slot_taken(self) -> bool:
        return self._active is not None or self._starting_id is not None

    # --- submission ---------------------------------------------------------

    def submit(
        self, request: RunRequest, credentials: dict[str, str] | None = None
    ) -> RunStateRecord:
        """Persist a queued request and start it as soon as the slot is free."""
        if self.closed:
            raise QueueBusyError("the run queue is shutting down")
        held = dict(credentials or {})
        with self._lock:
            if held and (self._slot_taken() or self._pending):
                held.clear()
                raise QueueBusyError(
                    "A run is already in progress. Password-protected files are "
                    "never queued — start this run again once the worker is free."
                )
            position = len(self._pending) + (1 if self._slot_taken() else 0)
            record = self.store.enqueue(
                request.request_id,
                mode=request.mode,
                profile=request.profile_name,
                files=dict(request.display_files),
                queue_position=position,
                profile_snapshot=dict(request.profile),
                requested_output_mode=request.requested_output_mode,
            )
            self._pending.append((request, held))
            self._ensure_supervisor()
        self._wake.set()
        return record

    def cancel(self, request_id: str) -> bool:
        """Cancel an active or queued request; queued work never starts."""
        with self._lock:
            job = self._active
            if job is not None and job.request.request_id == request_id:
                self._begin_cancel(job)
                return True
            if request_id == self._starting_id:
                # The slot is reserved but the process is still spawning.
                self._cancel_requested.add(request_id)
                self.store.request_cancel(request_id)
                self._wake.set()
                return True
            for index, (pending, held) in enumerate(self._pending):
                if pending.request_id != request_id:
                    continue
                held.clear()
                del self._pending[index]
                self._renumber()
                self.store.finish(request_id, RunStatus.CANCELLED)
                return True
        return False

    def _begin_cancel(self, job: _ActiveJob) -> None:
        if job.cancel_deadline is not None:
            return
        job.cancel_deadline = time.monotonic() + self._cancel_grace
        job.cancel_flag.set()
        self.store.request_cancel(job.request.request_id)
        self._wake.set()

    def wait(self, request_id: str, timeout: float = 120.0) -> RunStateRecord:
        """Block until a request reaches a terminal status (tests and CLI)."""
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            record = self.store.get(request_id)
            if record is not None and record.status not in ACTIVE_STATUSES:
                return record
            time.sleep(0.02)
        raise TimeoutError(f"request {request_id} did not finish within {timeout}s")

    def shutdown(self) -> None:
        """Stop the supervisor and retire any owned worker."""
        self._stopping.set()
        self._wake.set()
        thread = self._thread
        if thread is not None and thread is not threading.current_thread():
            # Long enough for a supervisor already inside the cancellation ladder.
            thread.join(timeout=self._cancel_grace + self._terminate_join * 3)
        with self._lock:
            job, self._active = self._active, None
            self._starting_id = None
            self._cancel_requested.clear()
            for _, held in self._pending:
                held.clear()
            self._pending.clear()
        if job is None:
            return
        self._stop_process(job)
        self._release(job)
        self.store.finish(
            job.request.request_id,
            RunStatus.ORPHANED,
            error="the server stopped before this request finished; submit it again",
        )

    # --- supervisor ---------------------------------------------------------

    def _ensure_supervisor(self) -> None:
        if self._thread is not None and self._thread.is_alive():
            return
        self._thread = threading.Thread(
            target=self._supervise, name="qc-run-queue", daemon=True
        )
        self._thread.start()

    def _supervise(self) -> None:
        while not self._stopping.is_set():
            try:
                self._service_active()
                self._start_next()
            except Exception:  # pragma: no cover - supervisor must never die
                logger.exception("run queue supervisor error")
            self._wake.wait(_SUPERVISOR_POLL_SECONDS)
            self._wake.clear()

    def _renumber(self) -> None:
        for index, (pending, _) in enumerate(self._pending):
            self.store.set_queue_position(pending.request_id, index + 1)

    def _start_next(self) -> None:
        with self._lock:
            if self._slot_taken() or not self._pending or self._stopping.is_set():
                return
            request, held = self._pending.popleft()
            self._starting_id = request.request_id
            self._renumber()
        self.store.set_queue_position(request.request_id, 0)
        self.store.mark_starting(request.request_id)
        try:
            events: IPCQueue[dict[str, Any]] = self._context.Queue(
                maxsize=IPC_QUEUE_CAPACITY
            )
            cancel_flag = self._context.Event()
            process = self._context.Process(
                target=self._worker,
                args=(asdict(request), held, events, cancel_flag),
                name=f"qc-worker-{request.request_id}",
                daemon=True,
            )
            process.start()
        except Exception as exc:
            logger.exception("could not start the QC worker process")
            with self._lock:
                self._starting_id = None
                self._cancel_requested.discard(request.request_id)
            self.store.finish(
                request.request_id, RunStatus.FAILED, error=sanitize_error(exc)
            )
            return
        finally:
            held.clear()
        with self._lock:
            job = _ActiveJob(
                request=request,
                process=process,
                events=events,
                cancel_flag=cancel_flag,
            )
            self._active = job
            self._starting_id = None
            cancel_now = request.request_id in self._cancel_requested
            self._cancel_requested.discard(request.request_id)
        self.store.mark_running(request.request_id)
        if cancel_now:
            with self._lock:
                self._begin_cancel(job)

    def _service_active(self) -> None:
        with self._lock:
            job = self._active
        if job is None:
            return
        self._drain(job)
        self._escalate_cancellation(job)
        if job.terminal_status is None and not job.process.is_alive():
            if job.dead_since is None:
                job.dead_since = time.monotonic()
            elif time.monotonic() - job.dead_since >= _EXIT_SETTLE_SECONDS:
                self._record_silent_exit(job)
        if job.terminal_status is not None:
            self._finish(job)

    def _record_silent_exit(self, job: _ActiveJob) -> None:
        if job.cancel_deadline is not None:
            job.terminal_status = RunStatus.CANCELLED
            job.error = job.error or "the QC worker was stopped before it reported"
            return
        job.terminal_status = RunStatus.FAILED
        job.error = (
            "the QC worker exited without a result "
            f"(exit code {job.process.exitcode})"
        )

    def _drain(self, job: _ActiveJob) -> None:
        while True:
            try:
                message = job.events.get_nowait()
            except queue_module.Empty:
                return
            except (OSError, ValueError, EOFError):
                if job.terminal_status is None:
                    job.terminal_status = RunStatus.FAILED
                    job.error = "the QC worker channel closed unexpectedly"
                return
            self._handle(job, message)

    def _handle(self, job: _ActiveJob, message: object) -> None:
        if not isinstance(message, dict) or message.get("v") != ENVELOPE_VERSION:
            logger.warning("discarding an unsupported QC worker message")
            return
        kind = message.get("kind")
        if kind == "progress":
            self.store.update_progress(
                job.request.request_id,
                phase=str(message.get("phase", "")),
                processed=int(message.get("processed", 0)),
                total=int(message.get("total", 0)),
                detail=str(message.get("detail", "")),
            )
            return
        if kind not in {"result", "cancelled", "error", "blocked"}:
            logger.warning("discarding an unsupported QC worker message")
            return
        phases = message.get("phases")
        job.phases = list(phases) if isinstance(phases, list) else []
        if kind == "result":
            job.terminal_status = RunStatus.SUCCEEDED
            job.run_id = int(message.get("run_id", 0)) or None
        elif kind == "cancelled":
            job.terminal_status = RunStatus.CANCELLED
        elif kind == "blocked":
            job.terminal_status = RunStatus.BLOCKED
            action_required = message.get("action_required")
            job.action_required = (
                action_required if isinstance(action_required, dict) else {}
            )
        else:
            job.terminal_status = RunStatus.FAILED
            job.error = sanitize_error(str(message.get("error", "")))

    def _escalate_cancellation(self, job: _ActiveJob) -> None:
        deadline = job.cancel_deadline
        if deadline is None or time.monotonic() < deadline:
            return
        if not job.process.is_alive():
            return
        now = time.monotonic()
        if not job.terminated:
            job.terminated = True
            job.cancel_deadline = now + self._terminate_join
            job.process.terminate()
            return
        if not job.killed:
            job.killed = True
            job.cancel_deadline = now + self._terminate_join
            kill = getattr(job.process, "kill", None)
            if callable(kill):
                kill()
            else:  # pragma: no cover - every supported platform exposes kill()
                job.process.terminate()

    def _stop_process(self, job: _ActiveJob) -> None:
        process = job.process
        if process.is_alive():
            process.terminate()
            process.join(self._terminate_join)
        if process.is_alive():
            kill = getattr(process, "kill", None)
            if callable(kill):
                kill()
            process.join(self._terminate_join)

    def _release(self, job: _ActiveJob) -> None:
        job.process.join(self._terminate_join)
        with contextlib.suppress(ValueError, AttributeError):
            job.process.close()
        with contextlib.suppress(OSError, ValueError):
            job.events.close()
            job.events.join_thread()

    def _finish(self, job: _ActiveJob) -> None:
        status = job.terminal_status or RunStatus.FAILED
        self._stop_process(job)
        self._release(job)
        try:
            if status is RunStatus.BLOCKED:
                self.store.finalize_blocked(
                    job.request.request_id,
                    job.action_required or {},
                    phases=job.phases,
                )
            else:
                self.store.finish(
                    job.request.request_id,
                    status,
                    run_id=job.run_id,
                    error=job.error,
                    phases=job.phases,
                )
        finally:
            # The slot is released even if the terminal write fails.
            with self._lock:
                self._active = None
            self._wake.set()


_MANAGERS: dict[Path, RunQueueManager] = {}
_MANAGERS_LOCK = threading.Lock()


def get_manager(work_dir: Path) -> RunQueueManager:
    """Return the process-global manager for one managed data directory."""
    key = work_dir.resolve()
    with _MANAGERS_LOCK:
        manager = _MANAGERS.get(key)
        if manager is None or manager.closed:
            manager = RunQueueManager(key)
            _MANAGERS[key] = manager
        return manager


def shutdown_managers() -> None:
    """Stop every process-global manager (server shutdown and tests)."""
    with _MANAGERS_LOCK:
        managers = list(_MANAGERS.values())
        _MANAGERS.clear()
    for manager in managers:
        manager.shutdown()
