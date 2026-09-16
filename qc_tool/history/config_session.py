"""Private, per-work-dir persistence for an in-progress configuration
session (plan-20260913-mode-aware-configuration-wizard.md, Step 5).

A configuration session holds the analyst's IN-PROGRESS setup choices for
one set of uploaded files -- e.g. confirmed member/sheet mappings, region
mode decisions -- so a browser refresh or reconnect restores them instead of
forcing the analyst to redo work. It is never a QC run and never creates a
`runs` history row.

Deliberately excluded from what is restored, by construction (no field
exists for them): preview grid content, formula text samples, and selector
candidate values. These are volatile, potentially large, and derived fresh
from the actual files every time a setup scan runs (Step 6); restoring them
would mean trusting stale bytes instead of the current file. Only DECISIONS
(small, primitive, analyst-confirmed choices) persist here; anything that
requires re-reading a file's content is always re-derived after restore.

The session key is derived only from the exact set of uploaded (role,
sha256) pairs -- never from a path or filename -- so re-uploading the SAME
files after a refresh resolves to the SAME session, and uploading DIFFERENT
files starts a fresh one. Managed-path/hash validation (confirming the
uploaded files this session refers to still exist and still match) is the
caller's job before trusting a restored session; this store only persists
and returns exactly what was saved.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import sqlite3
import uuid
from dataclasses import dataclass, field
from pathlib import Path

from qc_tool.security import private_directory, private_file

_SCHEMA = """
CREATE TABLE IF NOT EXISTS config_sessions (
    session_key TEXT PRIMARY KEY,
    session_id TEXT NOT NULL UNIQUE,
    source_set_digest TEXT NOT NULL DEFAULT '',
    revision INTEGER NOT NULL DEFAULT 1,
    input_generation INTEGER NOT NULL DEFAULT 1,
    input_signature TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    profile_name TEXT NOT NULL DEFAULT '',
    choices TEXT NOT NULL DEFAULT '{}'
);
"""


_MIGRATIONS = {
    "session_id": (
        "ALTER TABLE config_sessions ADD COLUMN session_id TEXT NOT NULL DEFAULT ''"
    ),
    "source_set_digest": (
        "ALTER TABLE config_sessions ADD COLUMN source_set_digest TEXT NOT NULL DEFAULT ''"
    ),
    "revision": (
        "ALTER TABLE config_sessions ADD COLUMN revision INTEGER NOT NULL DEFAULT 1"
    ),
    "input_generation": (
        "ALTER TABLE config_sessions ADD COLUMN input_generation INTEGER NOT NULL DEFAULT 1"
    ),
    "input_signature": (
        "ALTER TABLE config_sessions ADD COLUMN input_signature TEXT NOT NULL DEFAULT ''"
    ),
}


class ConfigSessionConflictError(RuntimeError):
    """A stale workspace attempted to overwrite newer session choices."""


def source_set_digest_for(file_hashes: dict[str, str]) -> str:
    """Return a stable, content-free digest for one role-to-hash mapping."""
    canonical = json.dumps(dict(sorted(file_hashes.items())), separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def session_key_for(file_hashes: dict[str, str]) -> str:
    """Return the legacy deterministic session alias for old links/tests."""
    return source_set_digest_for(file_hashes)


def _new_session_id() -> str:
    return uuid.uuid4().hex


def _string_mapping(value: object) -> dict[str, str]:
    if not isinstance(value, dict):
        return {}
    return {
        str(key): str(item)
        for key, item in value.items()
        if isinstance(key, str) and isinstance(item, str)
    }


def _input_signature(choices: dict[str, object]) -> str:
    """Digest only persisted inputs that invalidate setup-derived results."""
    files = choices.get("files")
    file_roles = sorted(str(role) for role in files) if isinstance(files, dict) else []
    payload = {
        "mode": str(choices.get("mode", "")),
        "file_roles": file_roles,
        "file_hashes": dict(sorted(_string_mapping(choices.get("file_hashes")).items())),
        "member_order": choices.get("member_order", {}),
    }
    canonical = json.dumps(payload, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


@dataclass(slots=True)
class ConfigSessionRecord:
    session_key: str
    session_id: str
    source_set_digest: str
    revision: int
    input_generation: int
    created_at: dt.datetime
    updated_at: dt.datetime
    profile_name: str = ""
    #: Primitive, analyst-confirmed choices only -- never preview content,
    #: formula text, or selector values (see module docstring).
    choices: dict[str, object] = field(default_factory=dict)


class ConfigSessionStore:
    """SQLite-backed store for one work dir's in-progress setup choices."""

    def __init__(self, db_path: Path) -> None:
        private_directory(db_path.parent)
        self._db_path = db_path
        with self._connect() as conn:
            conn.executescript(_SCHEMA)
            existing = {
                row[1] for row in conn.execute("PRAGMA table_info(config_sessions)")
            }
            for column, statement in _MIGRATIONS.items():
                if column not in existing:
                    conn.execute(statement)
            self._backfill_legacy_rows(conn)
            conn.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS "
                "idx_config_sessions_session_id ON config_sessions(session_id)"
            )
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_config_sessions_source_set_digest "
                "ON config_sessions(source_set_digest)"
            )
        private_file(db_path)

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self._db_path, timeout=30.0)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.row_factory = sqlite3.Row
        return conn

    @staticmethod
    def _backfill_legacy_rows(conn: sqlite3.Connection) -> None:
        rows = conn.execute(
            "SELECT session_key, choices FROM config_sessions WHERE session_id = ''"
        ).fetchall()
        for row in rows:
            try:
                choices = json.loads(row["choices"])
            except json.JSONDecodeError:
                choices = {}
            if not isinstance(choices, dict):
                choices = {}
            file_hashes = _string_mapping(choices.get("file_hashes"))
            source_set_digest = (
                source_set_digest_for(file_hashes) if file_hashes else row["session_key"]
            )
            conn.execute(
                """
                UPDATE config_sessions
                SET session_id = ?, source_set_digest = ?, input_signature = ?
                WHERE session_key = ?
                """,
                (
                    _new_session_id(),
                    source_set_digest,
                    _input_signature(choices),
                    row["session_key"],
                ),
            )

    @staticmethod
    def _record(row: sqlite3.Row) -> ConfigSessionRecord:
        return ConfigSessionRecord(
            session_key=row["session_key"],
            session_id=row["session_id"],
            source_set_digest=row["source_set_digest"],
            revision=row["revision"],
            input_generation=row["input_generation"],
            created_at=dt.datetime.fromisoformat(row["created_at"]),
            updated_at=dt.datetime.fromisoformat(row["updated_at"]),
            profile_name=row["profile_name"],
            choices=json.loads(row["choices"]),
        )

    def get(self, session_key: str) -> ConfigSessionRecord | None:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT * FROM config_sessions
                WHERE session_id = ? OR session_key = ?
                ORDER BY CASE WHEN session_id = ? THEN 0 ELSE 1 END
                LIMIT 1
                """,
                (session_key, session_key, session_key),
            ).fetchone()
        if row is None:
            return None
        return self._record(row)

    def latest_for_source_set(
        self, file_hashes: dict[str, str]
    ) -> ConfigSessionRecord | None:
        """Find the latest resumable session for these bytes, without authority."""
        source_set_digest = source_set_digest_for(file_hashes)
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT * FROM config_sessions
                WHERE source_set_digest = ?
                ORDER BY updated_at DESC, created_at DESC
                LIMIT 1
                """,
                (source_set_digest,),
            ).fetchone()
        return self._record(row) if row is not None else None

    def create_session(
        self,
        *,
        file_hashes: dict[str, str],
        profile_name: str = "",
        choices: dict[str, object],
    ) -> ConfigSessionRecord:
        """Create an independent session even when another uses the same files."""
        now = dt.datetime.now(dt.UTC).isoformat(timespec="milliseconds")
        session_id = _new_session_id()
        source_set_digest = source_set_digest_for(file_hashes)
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO config_sessions (
                    session_key, session_id, source_set_digest, revision,
                    input_generation, input_signature, created_at, updated_at,
                    profile_name, choices
                ) VALUES (?, ?, ?, 1, 1, ?, ?, ?, ?, ?)
                """,
                (
                    session_id,
                    session_id,
                    source_set_digest,
                    _input_signature(choices),
                    now,
                    now,
                    profile_name,
                    json.dumps(choices),
                ),
            )
        record = self.get(session_id)
        if record is None:  # pragma: no cover - the insert above just succeeded
            raise RuntimeError("config session row disappeared after creation")
        return record

    def save_choices(
        self,
        session_key: str,
        *,
        profile_name: str = "",
        choices: dict[str, object],
        expected_revision: int | None = None,
    ) -> ConfigSessionRecord:
        """Create or update the session's decisions. `choices` replaces the
        previously saved decisions wholesale -- callers own merging with any
        prior state they still want to keep.
        """
        now = dt.datetime.now(dt.UTC).isoformat(timespec="milliseconds")
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM config_sessions WHERE session_id = ? OR session_key = ?",
                (session_key, session_key),
            ).fetchone()
            if row is None:
                if expected_revision is not None:
                    raise ConfigSessionConflictError(
                        "configuration session no longer exists"
                    )
                session_id = _new_session_id()
                file_hashes = _string_mapping(choices.get("file_hashes"))
                source_set_digest = (
                    source_set_digest_for(file_hashes) if file_hashes else session_key
                )
                conn.execute(
                    """
                    INSERT INTO config_sessions (
                        session_key, session_id, source_set_digest, revision,
                        input_generation, input_signature, created_at, updated_at,
                        profile_name, choices
                    ) VALUES (?, ?, ?, 1, 1, ?, ?, ?, ?, ?)
                    """,
                    (
                        session_key,
                        session_id,
                        source_set_digest,
                        _input_signature(choices),
                        now,
                        now,
                        profile_name,
                        json.dumps(choices),
                    ),
                )
            else:
                current_revision = int(row["revision"])
                if expected_revision is not None and expected_revision != current_revision:
                    raise ConfigSessionConflictError(
                        "configuration session changed in another workspace"
                    )
                input_signature = _input_signature(choices)
                input_generation = int(row["input_generation"])
                if input_signature != row["input_signature"]:
                    input_generation += 1
                file_hashes = _string_mapping(choices.get("file_hashes"))
                source_set_digest = (
                    source_set_digest_for(file_hashes)
                    if file_hashes
                    else row["source_set_digest"]
                )
                cursor = conn.execute(
                    """
                    UPDATE config_sessions
                    SET updated_at = ?, profile_name = ?, choices = ?,
                        source_set_digest = ?, revision = revision + 1,
                        input_generation = ?, input_signature = ?
                    WHERE session_key = ? AND revision = ?
                    """,
                    (
                        now,
                        profile_name,
                        json.dumps(choices),
                        source_set_digest,
                        input_generation,
                        input_signature,
                        row["session_key"],
                        current_revision,
                    ),
                )
                if cursor.rowcount != 1:
                    raise ConfigSessionConflictError(
                        "configuration session changed in another workspace"
                    )
        record = self.get(session_key)
        if record is None:  # pragma: no cover - the upsert above just succeeded
            raise RuntimeError("config session row disappeared after save")
        return record

    def delete(self, session_key: str) -> None:
        with self._connect() as conn:
            conn.execute(
                "DELETE FROM config_sessions WHERE session_id = ? OR session_key = ?",
                (session_key, session_key),
            )

    def delete_stale(self, older_than: dt.datetime) -> int:
        """Bounded cleanup: remove sessions untouched since `older_than`."""
        cutoff = older_than.isoformat(timespec="milliseconds")
        with self._connect() as conn:
            cursor = conn.execute(
                "DELETE FROM config_sessions WHERE updated_at < ?", (cutoff,)
            )
            return cursor.rowcount
