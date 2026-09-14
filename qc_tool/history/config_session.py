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
from dataclasses import dataclass, field
from pathlib import Path

from qc_tool.security import private_directory, private_file

_SCHEMA = """
CREATE TABLE IF NOT EXISTS config_sessions (
    session_key TEXT PRIMARY KEY,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    profile_name TEXT NOT NULL DEFAULT '',
    choices TEXT NOT NULL DEFAULT '{}'
);
"""


def session_key_for(file_hashes: dict[str, str]) -> str:
    """Stable key derived only from role -> sha256 pairs -- never a path or
    filename, so it never carries private content into a log, error, or URL.
    """
    canonical = json.dumps(dict(sorted(file_hashes.items())), separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


@dataclass(slots=True)
class ConfigSessionRecord:
    session_key: str
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
        private_file(db_path)

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self._db_path, timeout=30.0)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.row_factory = sqlite3.Row
        return conn

    def get(self, session_key: str) -> ConfigSessionRecord | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM config_sessions WHERE session_key = ?",
                (session_key,),
            ).fetchone()
        if row is None:
            return None
        return ConfigSessionRecord(
            session_key=row["session_key"],
            created_at=dt.datetime.fromisoformat(row["created_at"]),
            updated_at=dt.datetime.fromisoformat(row["updated_at"]),
            profile_name=row["profile_name"],
            choices=json.loads(row["choices"]),
        )

    def save_choices(
        self,
        session_key: str,
        *,
        profile_name: str = "",
        choices: dict[str, object],
    ) -> ConfigSessionRecord:
        """Create or update the session's decisions. `choices` replaces the
        previously saved decisions wholesale -- callers own merging with any
        prior state they still want to keep.
        """
        now = dt.datetime.now(dt.UTC).isoformat(timespec="milliseconds")
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO config_sessions (
                    session_key, created_at, updated_at, profile_name, choices
                ) VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(session_key) DO UPDATE SET
                    updated_at = excluded.updated_at,
                    profile_name = excluded.profile_name,
                    choices = excluded.choices
                """,
                (session_key, now, now, profile_name, json.dumps(choices)),
            )
        record = self.get(session_key)
        if record is None:  # pragma: no cover - the upsert above just succeeded
            raise RuntimeError("config session row disappeared after save")
        return record

    def delete(self, session_key: str) -> None:
        with self._connect() as conn:
            conn.execute(
                "DELETE FROM config_sessions WHERE session_key = ?", (session_key,)
            )

    def delete_stale(self, older_than: dt.datetime) -> int:
        """Bounded cleanup: remove sessions untouched since `older_than`."""
        cutoff = older_than.isoformat(timespec="milliseconds")
        with self._connect() as conn:
            cursor = conn.execute(
                "DELETE FROM config_sessions WHERE updated_at < ?", (cutoff,)
            )
            return cursor.rowcount
