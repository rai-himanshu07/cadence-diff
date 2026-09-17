"""Progressive, reconnectable lifecycle coordinator for setup inspection."""

from __future__ import annotations

import copy
import datetime as dt
import json
import sqlite3
import threading
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path

from qc_tool.runqueue import QueueBusyError, run_exclusive
from qc_tool.security import private_directory, private_file
from qc_tool.setup.preview_worker import (
    SetupScanOutcome,
    SetupScanProgress,
    SetupScanRequest,
    run_setup_scan_worker,
)

_RETRY_WAIT_SECONDS = 0.25

_SCHEMA = """
CREATE TABLE IF NOT EXISTS setup_jobs (
    session_id TEXT NOT NULL,
    input_generation INTEGER NOT NULL,
    status TEXT NOT NULL,
    phase TEXT NOT NULL DEFAULT '',
    processed INTEGER NOT NULL DEFAULT 0,
    total INTEGER NOT NULL DEFAULT 0,
    requires_credentials INTEGER NOT NULL DEFAULT 0,
    credential_roles TEXT NOT NULL DEFAULT '{}',
    updated_at TEXT NOT NULL,
    PRIMARY KEY (session_id, input_generation)
);
"""

_MIGRATIONS = {
    "credential_roles": (
        "ALTER TABLE setup_jobs ADD COLUMN credential_roles TEXT "
        "NOT NULL DEFAULT '{}'"
    ),
}


class SetupStatus(StrEnum):
    AWAITING_INPUTS = "awaiting_inputs"
    AWAITING_CREDENTIALS = "awaiting_credentials"
    WAITING_FOR_SLOT = "waiting_for_slot"
    INVENTORY_READY = "inventory_ready"
    SCANNING = "scanning"
    PARTIAL_READY = "partial_ready"
    COMPLETE = "complete"
    CANCELLING = "cancelling"
    CANCELLED = "cancelled"
    FAILED = "failed"
    STALE = "stale"


_TERMINAL_STATUSES = frozenset(
    {
        SetupStatus.COMPLETE,
        SetupStatus.CANCELLED,
        SetupStatus.FAILED,
    }
)


@dataclass(frozen=True, slots=True)
class SetupMemberInput:
    member_id: str
    baseline_path: str = field(repr=False)
    current_path: str = field(repr=False)
    baseline_hash: str
    current_hash: str


@dataclass(frozen=True, slots=True)
class SetupJobSnapshot:
    session_id: str
    input_generation: int
    status: SetupStatus
    phase: str
    processed: int
    total: int
    member_status: dict[str, str]
    member_disclosures: dict[str, str]
    missing_credential_roles: dict[str, tuple[str, ...]]
    credential_roles_required: dict[str, tuple[str, ...]]
    terminal: bool


@dataclass(frozen=True, slots=True)
class SetupCoordinatorRecord:
    session_id: str
    input_generation: int
    status: SetupStatus
    phase: str
    processed: int
    total: int
    requires_credentials: bool
    credential_roles_required: dict[str, tuple[str, ...]]
    updated_at: dt.datetime


class SetupCoordinatorStore:
    """Persist content-free setup lifecycle facts for reconnect surfaces."""

    def __init__(self, db_path: Path) -> None:
        private_directory(db_path.parent)
        self._db_path = db_path
        with self._connect() as connection:
            connection.executescript(_SCHEMA)
            existing = {
                row[1] for row in connection.execute(
                    "PRAGMA table_info(setup_jobs)"
                )
            }
            for column, statement in _MIGRATIONS.items():
                if column not in existing:
                    connection.execute(statement)
        private_file(db_path)

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self._db_path, timeout=30.0)
        connection.execute("PRAGMA journal_mode=WAL")
        connection.row_factory = sqlite3.Row
        return connection

    def save(self, snapshot: SetupJobSnapshot, *, requires_credentials: bool) -> None:
        now = dt.datetime.now(dt.UTC).isoformat(timespec="milliseconds")
        with self._connect() as connection:
            connection.execute(
                """
                INSERT INTO setup_jobs (
                    session_id, input_generation, status, phase, processed,
                    total, requires_credentials, credential_roles, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(session_id, input_generation) DO UPDATE SET
                    status = excluded.status,
                    phase = excluded.phase,
                    processed = excluded.processed,
                    total = excluded.total,
                    requires_credentials = excluded.requires_credentials,
                    credential_roles = excluded.credential_roles,
                    updated_at = excluded.updated_at
                """,
                (
                    snapshot.session_id,
                    snapshot.input_generation,
                    snapshot.status.value,
                    snapshot.phase,
                    snapshot.processed,
                    snapshot.total,
                    int(requires_credentials),
                    json.dumps(snapshot.credential_roles_required),
                    now,
                ),
            )

    def get(self, session_id: str, input_generation: int) -> SetupCoordinatorRecord | None:
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT * FROM setup_jobs
                WHERE session_id = ? AND input_generation = ?
                """,
                (session_id, input_generation),
            ).fetchone()
        if row is None:
            return None
        return SetupCoordinatorRecord(
            session_id=str(row["session_id"]),
            input_generation=int(row["input_generation"]),
            status=SetupStatus(str(row["status"])),
            phase=str(row["phase"]),
            processed=int(row["processed"]),
            total=int(row["total"]),
            requires_credentials=bool(row["requires_credentials"]),
            credential_roles_required={
                str(member_id): tuple(str(side) for side in sides)
                for member_id, sides in json.loads(
                    str(row["credential_roles"])
                ).items()
                if isinstance(sides, list)
            },
            updated_at=dt.datetime.fromisoformat(str(row["updated_at"])),
        )

    def delete_stale(self, older_than: dt.datetime) -> int:
        cutoff = older_than.isoformat(timespec="milliseconds")
        with self._connect() as connection:
            cursor = connection.execute(
                "DELETE FROM setup_jobs WHERE updated_at < ?", (cutoff,)
            )
            return cursor.rowcount

    def delete(self, session_id: str, input_generation: int) -> None:
        with self._connect() as connection:
            connection.execute(
                """
                DELETE FROM setup_jobs
                WHERE session_id = ? AND input_generation = ?
                """,
                (session_id, input_generation),
            )


@dataclass(slots=True)
class SetupCoordinatorJob:
    session_id: str
    input_generation: int
    status: SetupStatus = SetupStatus.AWAITING_INPUTS
    phase: str = ""
    processed: int = 0
    total: int = 0
    member_status: dict[str, str] = field(default_factory=dict)
    result_payloads: dict[str, dict[str, object]] = field(default_factory=dict)
    member_disclosures: dict[str, str] = field(default_factory=dict)
    missing_credential_roles: dict[str, tuple[str, ...]] = field(default_factory=dict)
    cancel_event: threading.Event = field(default_factory=threading.Event, repr=False)
    subscribers: set[str] = field(default_factory=set, repr=False)
    _passwords_by_member: dict[str, dict[str, str]] = field(
        default_factory=dict, repr=False
    )
    _progress: dict[tuple[str, str], tuple[int, int]] = field(
        default_factory=dict, repr=False
    )
    _lock: threading.RLock = field(default_factory=threading.RLock, repr=False)
    _condition: threading.Condition = field(init=False, repr=False)
    _thread: threading.Thread | None = field(default=None, repr=False)
    _lease_revoked: threading.Event = field(
        default_factory=threading.Event, repr=False
    )
    _started: bool = field(default=False, repr=False)
    _dispatched: bool = field(default=False, repr=False)
    _requires_credentials: bool = field(default=False, repr=False)
    _credential_roles_required: dict[str, tuple[str, ...]] = field(
        default_factory=dict, repr=False
    )
    _finished: bool = field(default=False, repr=False)
    _updated_monotonic: float = field(default_factory=time.monotonic, repr=False)

    def __post_init__(self) -> None:
        self._condition = threading.Condition(self._lock)

    def snapshot(self) -> SetupJobSnapshot:
        with self._lock:
            return SetupJobSnapshot(
                session_id=self.session_id,
                input_generation=self.input_generation,
                status=self.status,
                phase=self.phase,
                processed=self.processed,
                total=self.total,
                member_status=dict(self.member_status),
                member_disclosures=dict(self.member_disclosures),
                missing_credential_roles=dict(self.missing_credential_roles),
                credential_roles_required=dict(self._credential_roles_required),
                terminal=self._finished,
            )

    def result_payloads_snapshot(self) -> dict[str, dict[str, object]]:
        with self._lock:
            return copy.deepcopy(self.result_payloads)

    def wait_terminal(self, timeout: float) -> bool:
        deadline = time.monotonic() + timeout
        with self._condition:
            while not self._finished:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    return False
                self._condition.wait(remaining)
            return True

    def wait_for_status(self, status: SetupStatus, timeout: float) -> bool:
        deadline = time.monotonic() + timeout
        with self._condition:
            while self.status is not status:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    return False
                self._condition.wait(remaining)
            return True


Runner = Callable[..., SetupScanOutcome]
ExclusiveRunner = Callable[[Path, str, Callable[[], SetupScanOutcome]], SetupScanOutcome]


class SetupCoordinator:
    """Own setup attempts independently of any one NiceGUI page closure."""

    def __init__(
        self,
        work_dir: Path,
        *,
        store: SetupCoordinatorStore | None = None,
        runner: Runner = run_setup_scan_worker,
        exclusive_runner: ExclusiveRunner = run_exclusive,
    ) -> None:
        self.work_dir = work_dir
        self.store = store or SetupCoordinatorStore(work_dir / "history.sqlite3")
        self._runner = runner
        self._exclusive_runner = exclusive_runner
        self._jobs: dict[tuple[str, int], SetupCoordinatorJob] = {}
        self._lock = threading.RLock()
        self._persist_lock = threading.Lock()
        self.shutdown_hook_installed = False

    def _key(self, session_id: str, input_generation: int) -> tuple[str, int]:
        return session_id, input_generation

    def find(
        self, session_id: str, input_generation: int
    ) -> SetupCoordinatorJob | None:
        with self._lock:
            return self._jobs.get(self._key(session_id, input_generation))

    def get_or_create(
        self, session_id: str, input_generation: int
    ) -> SetupCoordinatorJob:
        key = self._key(session_id, input_generation)
        with self._lock:
            job = self._jobs.get(key)
            if job is not None:
                return job
            persisted = self.store.get(session_id, input_generation)
            status = SetupStatus.AWAITING_INPUTS
            phase = ""
            processed = total = 0
            if persisted is not None:
                phase = persisted.phase
                processed = persisted.processed
                total = persisted.total
                if persisted.status in _TERMINAL_STATUSES:
                    status = persisted.status
                elif (
                    persisted.status is SetupStatus.AWAITING_CREDENTIALS
                    or persisted.requires_credentials
                ):
                    status = SetupStatus.AWAITING_CREDENTIALS
                else:
                    status = SetupStatus.STALE
            job = SetupCoordinatorJob(
                session_id=session_id,
                input_generation=input_generation,
                status=status,
                phase=phase,
                processed=processed,
                total=total,
                _credential_roles_required=(
                    dict(persisted.credential_roles_required)
                    if persisted is not None
                    else {}
                ),
                _finished=status in _TERMINAL_STATUSES,
            )
            self._jobs[key] = job
            return job

    def attach(
        self, session_id: str, input_generation: int, subscriber_id: str
    ) -> SetupCoordinatorJob:
        job = self.get_or_create(session_id, input_generation)
        with job._condition:
            job.subscribers.add(subscriber_id)
            job._updated_monotonic = time.monotonic()
            job._condition.notify_all()
        return job

    def detach(self, session_id: str, input_generation: int, subscriber_id: str) -> None:
        job = self.find(session_id, input_generation)
        if job is None:
            return
        demote = False
        with job._condition:
            job.subscribers.discard(subscriber_id)
            if (
                not job.subscribers
                and job.status is SetupStatus.WAITING_FOR_SLOT
                and bool(job._passwords_by_member)
            ):
                job._passwords_by_member.clear()
                job._lease_revoked.set()
                job.status = SetupStatus.AWAITING_CREDENTIALS
                job.phase = "awaiting_credentials"
                job._started = False
                job._updated_monotonic = time.monotonic()
                job._condition.notify_all()
                demote = True
        if demote:
            self._persist(job)

    def start(
        self,
        session_id: str,
        input_generation: int,
        members: dict[str, SetupMemberInput],
        *,
        passwords_by_member: dict[str, dict[str, str]],
    ) -> SetupCoordinatorJob:
        job = self.get_or_create(session_id, input_generation)
        with job._condition:
            if job._started and job._thread is not None and job._thread.is_alive():
                return job
            if job.status is SetupStatus.COMPLETE:
                from qc_tool.setup.preview_store import SetupScanStore

                sidecar = SetupScanStore(self.work_dir / "setup-inspection.sqlite3")
                restored = {
                    member_id: payload
                    for member_id in members
                    if (
                        payload := sidecar.get_result(
                            session_id,
                            member_id,
                            input_generation=input_generation,
                        )
                    )
                    is not None
                }
                if len(restored) == len(members):
                    job.result_payloads = restored
                    return job
                job.status = SetupStatus.STALE
                job.phase = "stale"
                job._finished = False
            job.cancel_event = threading.Event()
            job._lease_revoked = threading.Event()
            job._passwords_by_member = {
                member_id: {
                    side: password
                    for side, password in passwords.items()
                    if password
                }
                for member_id, passwords in passwords_by_member.items()
                if any(passwords.values())
            }
            job._requires_credentials = bool(job._passwords_by_member)
            job._credential_roles_required = {
                member_id: tuple(sorted(passwords))
                for member_id, passwords in job._passwords_by_member.items()
            }
            job.member_status = dict.fromkeys(members, "queued")
            job.member_disclosures.clear()
            job.missing_credential_roles.clear()
            job.result_payloads = {
                member_id: {
                    "member_id": member.member_id,
                    "baseline_hash": member.baseline_hash,
                    "current_hash": member.current_hash,
                    "baseline_sheets": [],
                    "current_sheets": [],
                }
                for member_id, member in members.items()
            }
            job._progress.clear()
            job.processed = 0
            job.total = 0
            job.status = (
                SetupStatus.WAITING_FOR_SLOT
                if members
                else SetupStatus.AWAITING_INPUTS
            )
            job.phase = job.status.value
            job._started = bool(members)
            job._dispatched = False
            job._finished = False
            job._updated_monotonic = time.monotonic()
            if members:
                job._thread = threading.Thread(
                    target=self._run_job,
                    args=(job, dict(members)),
                    daemon=True,
                )
                job._thread.start()
            job._condition.notify_all()
        self._persist(job)
        return job

    def restart(
        self, session_id: str, input_generation: int
    ) -> SetupCoordinatorJob:
        key = self._key(session_id, input_generation)
        with self._lock:
            previous = self._jobs.get(key)
            subscribers = set(previous.subscribers) if previous is not None else set()
            if previous is not None:
                previous.cancel_event.set()
                previous._lease_revoked.set()
            replacement = SetupCoordinatorJob(
                session_id=session_id,
                input_generation=input_generation,
                subscribers=subscribers,
            )
            self._jobs[key] = replacement
        self._persist(replacement)
        return replacement

    def cancel(self, session_id: str, input_generation: int) -> bool:
        job = self.find(session_id, input_generation)
        if job is None:
            return False
        with job._condition:
            if job._finished:
                return False
            job.cancel_event.set()
            job._lease_revoked.set()
            job._passwords_by_member.clear()
            if job._thread is None or not job._thread.is_alive():
                job.status = SetupStatus.CANCELLED
                job.phase = "cancelled"
                job._finished = True
            else:
                job.status = SetupStatus.CANCELLING
                job.phase = "cancelling"
            job._updated_monotonic = time.monotonic()
            job._condition.notify_all()
        self._persist(job)
        return True

    def discard(
        self, session_id: str, input_generation: int
    ) -> threading.Thread | None:
        """Forget one generation after cancellation; selected files live elsewhere."""
        key = self._key(session_id, input_generation)
        with self._lock:
            job = self._jobs.pop(key, None)
        if job is not None:
            job.cancel_event.set()
            job._lease_revoked.set()
            job._passwords_by_member.clear()
        self.store.delete(session_id, input_generation)

        def reclaim_sidecar() -> None:
            thread = job._thread if job is not None else None
            if thread is not None and thread is not threading.current_thread():
                thread.join()
            from qc_tool.setup.preview_store import SetupScanStore

            SetupScanStore(
                self.work_dir / "setup-inspection.sqlite3"
            ).delete_generation(session_id, input_generation)

        thread = job._thread if job is not None else None
        if thread is not None and thread.is_alive():
            cleanup_thread = threading.Thread(
                target=reclaim_sidecar, daemon=True
            )
            cleanup_thread.start()
            return cleanup_thread
        reclaim_sidecar()
        return None

    def evict_terminal(self, *, max_age_seconds: float) -> int:
        now = time.monotonic()
        removed = 0
        with self._lock:
            for key, job in tuple(self._jobs.items()):
                with job._lock:
                    if (
                        job._finished
                        and not job.subscribers
                        and now - job._updated_monotonic >= max_age_seconds
                    ):
                        self._jobs.pop(key, None)
                        removed += 1
        return removed

    def shutdown(self) -> None:
        with self._lock:
            jobs = tuple(self._jobs.values())
        for job in jobs:
            self.cancel(job.session_id, job.input_generation)

    def _is_current(self, job: SetupCoordinatorJob) -> bool:
        with self._lock:
            return self._jobs.get(self._key(job.session_id, job.input_generation)) is job

    def _persist(self, job: SetupCoordinatorJob) -> None:
        with self._persist_lock:
            if self._is_current(job):
                self.store.save(
                    job.snapshot(), requires_credentials=job._requires_credentials
                )

    def _update_progress(
        self,
        job: SetupCoordinatorJob,
        member_id: str,
        progress: SetupScanProgress,
    ) -> None:
        if not self._is_current(job) or job.cancel_event.is_set():
            return
        with job._condition:
            job._progress[(member_id, progress.side)] = (
                progress.processed,
                progress.total,
            )
            job.processed = sum(value[0] for value in job._progress.values())
            job.total = sum(value[1] for value in job._progress.values())
            job.phase = progress.phase
            job.status = (
                SetupStatus.PARTIAL_READY
                if progress.phase == "sheet_ready"
                else SetupStatus.INVENTORY_READY
            )
            if progress.profile_payload is not None and progress.side in {
                "baseline",
                "current",
            }:
                member_payload = job.result_payloads.setdefault(
                    member_id,
                    {
                        "member_id": member_id,
                        "baseline_sheets": [],
                        "current_sheets": [],
                    },
                )
                key = f"{progress.side}_sheets"
                existing = member_payload.get(key)
                profiles = list(existing) if isinstance(existing, list) else []
                sheet_name = progress.profile_payload.get("sheet_name")
                profiles = [
                    profile
                    for profile in profiles
                    if not isinstance(profile, dict)
                    or profile.get("sheet_name") != sheet_name
                ]
                profiles.append(dict(progress.profile_payload))
                member_payload[key] = profiles
            job._updated_monotonic = time.monotonic()
            job._condition.notify_all()
        self._persist(job)

    def _run_job(
        self,
        job: SetupCoordinatorJob,
        members: dict[str, SetupMemberInput],
    ) -> None:
        for member_id, member in members.items():
            outcome: SetupScanOutcome | None = None
            if not self._is_current(job) or job.cancel_event.is_set():
                break
            while True:
                if not self._is_current(job) or job.cancel_event.is_set():
                    break
                with job._condition:
                    if job._lease_revoked.is_set():
                        return
                    passwords = dict(job._passwords_by_member.get(member_id, {}))
                    job.status = SetupStatus.WAITING_FOR_SLOT
                    job.phase = "waiting_for_slot"
                    job.member_status[member_id] = "waiting"
                    job._condition.notify_all()
                self._persist(job)

                def run_member(
                    member_input: SetupMemberInput = member,
                    current_member_id: str = member_id,
                    member_passwords: dict[str, str] = passwords,
                ) -> SetupScanOutcome:
                    with job._condition:
                        if job._lease_revoked.is_set() or job.cancel_event.is_set():
                            return SetupScanOutcome(disclosure="setup scan cancelled")
                        job._dispatched = True
                        job.status = SetupStatus.SCANNING
                        job.phase = "scanning"
                        job.member_status[current_member_id] = "running"
                        job._passwords_by_member.pop(current_member_id, None)
                        job._updated_monotonic = time.monotonic()
                        job._condition.notify_all()
                    self._persist(job)
                    request = SetupScanRequest(
                        member_id=member_input.member_id,
                        baseline_path=member_input.baseline_path,
                        current_path=member_input.current_path,
                        baseline_hash=member_input.baseline_hash,
                        current_hash=member_input.current_hash,
                        sidecar_path=str(self.work_dir / "setup-inspection.sqlite3"),
                        session_key=job.session_id,
                        input_generation=job.input_generation,
                        passwords=member_passwords,
                    )
                    return self._runner(
                        request,
                        cancel_event=job.cancel_event,
                        on_progress=lambda progress: self._update_progress(
                            job, current_member_id, progress
                        ),
                    )

                try:
                    outcome = self._exclusive_runner(
                        self.work_dir,
                        f"setup-scan:{job.session_id}:{member_id}",
                        run_member,
                    )
                except QueueBusyError:
                    if job._lease_revoked.wait(_RETRY_WAIT_SECONDS):
                        return
                    continue
                break
            if not self._is_current(job) or job.cancel_event.is_set():
                break
            if outcome is None:
                continue
            awaiting_credentials = False
            with job._condition:
                if outcome.missing_credential_roles:
                    job._requires_credentials = True
                    job._credential_roles_required[member_id] = tuple(
                        sorted(outcome.missing_credential_roles)
                    )
                    job.status = SetupStatus.AWAITING_CREDENTIALS
                    job.phase = "awaiting_credentials"
                    job.member_status[member_id] = "awaiting_credentials"
                    job.missing_credential_roles[member_id] = (
                        outcome.missing_credential_roles
                    )
                    job._passwords_by_member.clear()
                    job._started = False
                    job._updated_monotonic = time.monotonic()
                    job._condition.notify_all()
                    awaiting_credentials = True
                elif outcome.result_payload is None:
                    job.member_status[member_id] = "failed"
                    job.member_disclosures[member_id] = (
                        outcome.disclosure or "setup analysis failed"
                    )
                else:
                    job.member_status[member_id] = "done"
                    job.result_payloads[member_id] = outcome.result_payload
                job._updated_monotonic = time.monotonic()
                job._condition.notify_all()
            self._persist(job)
            if awaiting_credentials:
                return

        if not self._is_current(job):
            return
        with job._condition:
            if job.cancel_event.is_set():
                job.status = SetupStatus.CANCELLED
                job.phase = "cancelled"
                job.result_payloads.clear()
            elif job.member_status and all(
                status == "done" for status in job.member_status.values()
            ):
                job.status = SetupStatus.COMPLETE
                job.phase = "complete"
            elif job.result_payloads:
                job.status = SetupStatus.PARTIAL_READY
                job.phase = "partial_ready"
            else:
                job.status = SetupStatus.FAILED
                job.phase = "failed"
            job._started = False
            job._updated_monotonic = time.monotonic()
        self._persist(job)
        with job._condition:
            job._finished = True
            job._condition.notify_all()


_COORDINATORS: dict[Path, SetupCoordinator] = {}
_COORDINATORS_LOCK = threading.Lock()


def get_setup_coordinator(work_dir: Path) -> SetupCoordinator:
    key = work_dir.resolve()
    with _COORDINATORS_LOCK:
        coordinator = _COORDINATORS.get(key)
        if coordinator is None:
            coordinator = SetupCoordinator(key)
            _COORDINATORS[key] = coordinator
        return coordinator
