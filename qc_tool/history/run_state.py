"""Private lifecycle state for active and queued QC run requests.

Completed runs stay in the `runs` table; this table only tracks work that is
queued or in flight so a browser refresh can reconnect to it. It never stores
passwords, source paths, file contents, formulas, or findings.
"""

import datetime as dt
import json
import sqlite3
from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path

from qc_tool.security import private_directory, private_file

_SCHEMA = """
CREATE TABLE IF NOT EXISTS run_state (
    request_id TEXT PRIMARY KEY,
    created_at TEXT NOT NULL,
    status TEXT NOT NULL,
    queue_position INTEGER NOT NULL DEFAULT 0,
    mode TEXT NOT NULL,
    profile TEXT NOT NULL,
    files TEXT NOT NULL DEFAULT '{}',
    phase TEXT NOT NULL DEFAULT '',
    processed INTEGER NOT NULL DEFAULT 0,
    total INTEGER NOT NULL DEFAULT 0,
    detail TEXT NOT NULL DEFAULT '',
    started_at TEXT NOT NULL DEFAULT '',
    finished_at TEXT NOT NULL DEFAULT '',
    cancel_requested INTEGER NOT NULL DEFAULT 0,
    run_id INTEGER,
    error TEXT NOT NULL DEFAULT '',
    phases TEXT NOT NULL DEFAULT '[]',
    action_required TEXT NOT NULL DEFAULT '{}'
);
"""


class RunStatus(StrEnum):
    QUEUED = "queued"
    STARTING = "starting"
    RUNNING = "running"
    CANCELLING = "cancelling"
    SUCCEEDED = "succeeded"
    FAILED = "failed"
    CANCELLED = "cancelled"
    ORPHANED = "orphaned"
    #: Terminal: the run cannot proceed until the analyst takes an action
    #: outside this tool (see `qc_tool.run_action`). Never assigned a run id.
    BLOCKED = "blocked"


#: Statuses that still occupy or await the single worker slot.
ACTIVE_STATUSES = frozenset(
    {
        RunStatus.QUEUED,
        RunStatus.STARTING,
        RunStatus.RUNNING,
        RunStatus.CANCELLING,
    }
)


def _now() -> str:
    return dt.datetime.now(dt.UTC).isoformat(timespec="milliseconds")


#: New columns added after the table's initial release; existing databases
#: are migrated in place so a fresh row and a legacy row read identically.
_MIGRATIONS = {
    "action_required": (
        "ALTER TABLE run_state ADD COLUMN action_required TEXT NOT NULL DEFAULT '{}'"
    ),
}


@dataclass(slots=True)
class RunStateRecord:
    request_id: str
    created_at: dt.datetime
    status: RunStatus
    mode: str
    profile: str
    #: Role -> display filename only; never a source path.
    files: dict[str, str] = field(default_factory=dict)
    queue_position: int = 0
    phase: str = ""
    processed: int = 0
    total: int = 0
    detail: str = ""
    started_at: dt.datetime | None = None
    finished_at: dt.datetime | None = None
    cancel_requested: bool = False
    run_id: int | None = None
    error: str = ""
    phases: list[dict[str, object]] = field(default_factory=list)
    #: Bounded, primitive-only payload for `RunStatus.BLOCKED`; see
    #: `qc_tool.run_action.RunActionRequired`. Empty/`None` otherwise.
    action_required: dict[str, object] | None = None

    @property
    def is_active(self) -> bool:
        return self.status in ACTIVE_STATUSES

    def elapsed_seconds(self, now: dt.datetime | None = None) -> float:
        start = self.started_at or self.created_at
        end = self.finished_at or now or dt.datetime.now(dt.UTC)
        return max(0.0, (end - start).total_seconds())


def _timestamp(value: str) -> dt.datetime | None:
    return dt.datetime.fromisoformat(value) if value else None


def _decode_action_required(row: sqlite3.Row) -> dict[str, object] | None:
    try:
        raw = row["action_required"]
    except (KeyError, IndexError):
        return None
    try:
        decoded = json.loads(raw)
    except json.JSONDecodeError:
        return None
    return decoded or None


def _record(row: sqlite3.Row) -> RunStateRecord:
    return RunStateRecord(
        request_id=row["request_id"],
        created_at=dt.datetime.fromisoformat(row["created_at"]),
        status=RunStatus(row["status"]),
        mode=row["mode"],
        profile=row["profile"],
        files=json.loads(row["files"]),
        queue_position=row["queue_position"],
        phase=row["phase"],
        processed=row["processed"],
        total=row["total"],
        detail=row["detail"],
        started_at=_timestamp(row["started_at"]),
        finished_at=_timestamp(row["finished_at"]),
        cancel_requested=bool(row["cancel_requested"]),
        run_id=row["run_id"],
        error=row["error"],
        phases=json.loads(row["phases"]),
        action_required=(_decode_action_required(row)),
    )


class RunStateStore:
    """SQLite persistence for the single-worker run queue."""

    def __init__(self, db_path: Path) -> None:
        private_directory(db_path.parent)
        self._db_path = db_path
        with self._connect() as conn:
            conn.executescript(_SCHEMA)
            existing = {row[1] for row in conn.execute("PRAGMA table_info(run_state)")}
            for column, statement in _MIGRATIONS.items():
                if column not in existing:
                    conn.execute(statement)
        private_file(db_path)

    def _connect(self) -> sqlite3.Connection:
        # The worker process, the queue manager, and the UI share this file.
        conn = sqlite3.connect(self._db_path, timeout=30.0)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.row_factory = sqlite3.Row
        return conn

    def enqueue(
        self,
        request_id: str,
        *,
        mode: str,
        profile: str,
        files: dict[str, str],
        queue_position: int,
    ) -> RunStateRecord:
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO run_state (
                    request_id, created_at, status, queue_position, mode, profile, files
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    request_id,
                    _now(),
                    RunStatus.QUEUED.value,
                    queue_position,
                    mode,
                    profile,
                    json.dumps(files),
                ),
            )
        record = self.get(request_id)
        if record is None:  # pragma: no cover - the insert above just succeeded
            raise RuntimeError("run state row disappeared after insert")
        return record

    def get(self, request_id: str) -> RunStateRecord | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM run_state WHERE request_id = ?", (request_id,)
            ).fetchone()
        return None if row is None else _record(row)

    def pending(self) -> list[RunStateRecord]:
        """Active and queued requests in submission order."""
        placeholders = ", ".join("?" for _ in ACTIVE_STATUSES)
        statuses = sorted(status.value for status in ACTIVE_STATUSES)
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT * FROM run_state WHERE status IN ({placeholders}) "
                "ORDER BY queue_position, created_at, rowid",
                statuses,
            ).fetchall()
        return [_record(row) for row in rows]

    def set_queue_position(self, request_id: str, position: int) -> None:
        with self._connect() as conn:
            conn.execute(
                "UPDATE run_state SET queue_position = ? WHERE request_id = ?",
                (position, request_id),
            )

    def mark_starting(self, request_id: str) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                UPDATE run_state SET status = ?, started_at = ?, queue_position = 0
                WHERE request_id = ? AND status = ?
                """,
                (RunStatus.STARTING.value, _now(), request_id, RunStatus.QUEUED.value),
            )

    def mark_running(self, request_id: str) -> None:
        """No-op once a progress message has already advanced the status."""
        with self._connect() as conn:
            conn.execute(
                "UPDATE run_state SET status = ? WHERE request_id = ? AND status = ?",
                (RunStatus.RUNNING.value, request_id, RunStatus.STARTING.value),
            )

    def update_progress(
        self,
        request_id: str,
        *,
        phase: str,
        processed: int,
        total: int,
        detail: str,
    ) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                UPDATE run_state
                SET phase = ?, processed = ?, total = ?, detail = ?,
                    status = CASE status WHEN ? THEN ? ELSE status END
                WHERE request_id = ?
                """,
                (
                    phase,
                    processed,
                    total,
                    detail,
                    RunStatus.STARTING.value,
                    RunStatus.RUNNING.value,
                    request_id,
                ),
            )

    def request_cancel(self, request_id: str) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                UPDATE run_state
                SET cancel_requested = 1,
                    status = CASE WHEN status IN (?, ?, ?) THEN ? ELSE status END
                WHERE request_id = ?
                """,
                (
                    RunStatus.QUEUED.value,
                    RunStatus.STARTING.value,
                    RunStatus.RUNNING.value,
                    RunStatus.CANCELLING.value,
                    request_id,
                ),
            )

    def finish(
        self,
        request_id: str,
        status: RunStatus,
        *,
        run_id: int | None = None,
        error: str = "",
        phases: list[dict[str, object]] | None = None,
    ) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                UPDATE run_state
                SET status = ?, finished_at = ?, queue_position = 0,
                    run_id = ?, error = ?, phases = ?
                WHERE request_id = ?
                """,
                (
                    status.value,
                    _now(),
                    run_id,
                    error,
                    json.dumps(phases or []),
                    request_id,
                ),
            )

    def finalize_blocked(
        self,
        request_id: str,
        action_required: dict[str, object],
        *,
        phases: list[dict[str, object]] | None = None,
    ) -> None:
        """Terminal state for a run that cannot proceed; never assigns a run id.

        Stored in a dedicated `action_required` column rather than `error`,
        since this is not a failure -- it is a bounded, structured reason the
        analyst must act on outside this tool.
        """
        with self._connect() as conn:
            conn.execute(
                """
                UPDATE run_state
                SET status = ?, finished_at = ?, queue_position = 0,
                    action_required = ?, phases = ?
                WHERE request_id = ?
                """,
                (
                    RunStatus.BLOCKED.value,
                    _now(),
                    json.dumps(action_required),
                    json.dumps(phases or []),
                    request_id,
                ),
            )

    def mark_orphaned(self) -> list[str]:
        """Retire requests left incomplete by a previous server process."""
        pending = [record.request_id for record in self.pending()]
        if not pending:
            return []
        placeholders = ", ".join("?" for _ in pending)
        with self._connect() as conn:
            conn.execute(
                f"""
                UPDATE run_state
                SET status = ?, finished_at = ?, queue_position = 0, error = ?
                WHERE request_id IN ({placeholders})
                """,
                (
                    RunStatus.ORPHANED.value,
                    _now(),
                    "the server stopped before this request finished; submit it again",
                    *pending,
                ),
            )
        return pending
