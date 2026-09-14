"""Private, compressed sidecar storage for setup-analysis scan results
(plan-20260913, Step 6) plus bounded post-scan queries against it.

Persists exactly what
``qc_tool.setup.preview_worker.SetupScanOutcome.result_payload`` already is:
a JSON-safe, primitive-only dict (never raw cell values, formula text, or
selector values -- see ``qc_tool.setup.models``' own docstring for what
this pipeline deliberately never observes). Compressed at rest (zlib),
mirroring this project's own established "private block sidecar"
convention (``qc_tool.history.run_state``'s focus blocks,
``qc_tool.findings_store``'s spill blocks) rather than storing a
potentially large JSON blob uncompressed.

``get_sheet_profile`` is the bounded post-scan query this pipeline's own
verify criterion calls for: a caller that only needs one sheet's
structural facts (to render one panel of a setup workspace, say) never
needs to know the whole result's shape.
"""

from __future__ import annotations

import datetime as dt
import json
import sqlite3
import zlib
from dataclasses import dataclass
from pathlib import Path

from qc_tool.security import private_directory, private_file

_SCHEMA = """
CREATE TABLE IF NOT EXISTS setup_scan_blocks (
    session_key TEXT NOT NULL,
    member_id TEXT NOT NULL,
    created_at TEXT NOT NULL,
    blob BLOB NOT NULL,
    PRIMARY KEY (session_key, member_id)
);
"""


@dataclass(slots=True)
class SetupScanSidecar:
    session_key: str
    member_id: str
    created_at: str
    #: Compressed byte size on disk -- a bounded, disclosable size fact,
    #: never the content itself.
    blob_bytes: int


class SetupScanStore:
    """SQLite-backed, compressed sidecar for one work dir's setup-scan
    results, keyed by ``(session_key, member_id)``.
    """

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

    def save_result(
        self,
        session_key: str,
        member_id: str,
        result_payload: dict[str, object],
    ) -> SetupScanSidecar:
        now = dt.datetime.now(dt.UTC).isoformat(timespec="milliseconds")
        blob = zlib.compress(
            json.dumps(result_payload, separators=(",", ":")).encode("utf-8"), level=6
        )
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO setup_scan_blocks (
                    session_key, member_id, created_at, blob
                ) VALUES (?, ?, ?, ?)
                ON CONFLICT(session_key, member_id) DO UPDATE SET
                    created_at = excluded.created_at,
                    blob = excluded.blob
                """,
                (session_key, member_id, now, blob),
            )
        return SetupScanSidecar(
            session_key=session_key,
            member_id=member_id,
            created_at=now,
            blob_bytes=len(blob),
        )

    def get_result(self, session_key: str, member_id: str) -> dict[str, object] | None:
        """The full decoded result payload. Prefer ``get_sheet_profile`` for
        a single-sheet lookup -- this bounded scan result is small enough
        that decoding it whole stays cheap, but a targeted query avoids
        handing a caller the entire tree when it only needs one sheet.
        """
        with self._connect() as conn:
            row = conn.execute(
                "SELECT blob FROM setup_scan_blocks"
                " WHERE session_key = ? AND member_id = ?",
                (session_key, member_id),
            ).fetchone()
        if row is None:
            return None
        decoded = json.loads(zlib.decompress(row["blob"]).decode("utf-8"))
        return decoded if isinstance(decoded, dict) else None

    def get_sheet_profile(
        self,
        session_key: str,
        member_id: str,
        sheet_name: str,
        *,
        side: str = "current",
    ) -> dict[str, object] | None:
        """Bounded post-profile query: one sheet's structural facts, without
        the caller needing to know the whole result's shape. ``side`` is
        ``"current"`` or ``"baseline"``.
        """
        result = self.get_result(session_key, member_id)
        if result is None:
            return None
        sheets = result.get(f"{side}_sheets")
        if not isinstance(sheets, list):
            return None
        for sheet in sheets:
            if isinstance(sheet, dict) and sheet.get("sheet_name") == sheet_name:
                return sheet
        return None

    def delete(self, session_key: str) -> None:
        with self._connect() as conn:
            conn.execute(
                "DELETE FROM setup_scan_blocks WHERE session_key = ?", (session_key,)
            )

    def delete_stale(self, older_than: dt.datetime) -> int:
        """Bounded cleanup: remove scan results untouched since ``older_than``."""
        cutoff = older_than.isoformat(timespec="milliseconds")
        with self._connect() as conn:
            cursor = conn.execute(
                "DELETE FROM setup_scan_blocks WHERE created_at < ?", (cutoff,)
            )
            return cursor.rowcount
