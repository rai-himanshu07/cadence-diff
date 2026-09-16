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
import hashlib
import json
import math
import sqlite3
import zlib
from dataclasses import dataclass
from pathlib import Path

from qc_tool.io.model import CellRecord, CellValue, SheetSnapshot
from qc_tool.security import private_directory, private_file

_SCHEMA = """
CREATE TABLE IF NOT EXISTS setup_scan_blocks (
    session_key TEXT NOT NULL,
    member_id TEXT NOT NULL,
    created_at TEXT NOT NULL,
    blob BLOB NOT NULL,
    PRIMARY KEY (session_key, member_id)
);
CREATE TABLE IF NOT EXISTS setup_sheet_inventory (
    session_key TEXT NOT NULL,
    member_id TEXT NOT NULL,
    side TEXT NOT NULL,
    sheet_key TEXT NOT NULL,
    input_generation INTEGER NOT NULL,
    source_hash TEXT NOT NULL,
    created_at TEXT NOT NULL,
    metadata_blob BLOB NOT NULL,
    block_count INTEGER NOT NULL,
    PRIMARY KEY (session_key, member_id, side, sheet_key)
);
CREATE TABLE IF NOT EXISTS setup_cell_blocks (
    session_key TEXT NOT NULL,
    member_id TEXT NOT NULL,
    side TEXT NOT NULL,
    sheet_key TEXT NOT NULL,
    block_index INTEGER NOT NULL,
    min_row INTEGER NOT NULL,
    max_row INTEGER NOT NULL,
    blob BLOB NOT NULL,
    PRIMARY KEY (
        session_key, member_id, side, sheet_key, block_index
    )
);
CREATE INDEX IF NOT EXISTS idx_setup_cell_blocks_rows
ON setup_cell_blocks (
    session_key, member_id, side, sheet_key, min_row, max_row
);
"""

_DEFAULT_BLOCK_ROWS = 512
_MAX_DECOMPRESSED_BLOCK_BYTES = 32 * 1024 * 1024
_MAX_FORMULA_TEXT_CHARS = 256


class SetupSidecarError(RuntimeError):
    """Base class for content-free setup-sidecar failures."""


class SetupSidecarStaleError(SetupSidecarError):
    """A query does not match the sidecar's source generation."""


class SetupSidecarCorruptError(SetupSidecarError):
    """A private sidecar block failed bounded validation."""


@dataclass(slots=True)
class SetupScanSidecar:
    session_key: str
    member_id: str
    created_at: str
    #: Compressed byte size on disk -- a bounded, disclosable size fact,
    #: never the content itself.
    blob_bytes: int


@dataclass(frozen=True, slots=True)
class SetupSheetInventory:
    sheet_name: str
    visibility: str
    max_row: int
    max_column: int
    hidden_rows: frozenset[int]
    hidden_columns: frozenset[int]
    block_count: int


@dataclass(frozen=True, slots=True)
class SetupWindow:
    cells: dict[tuple[int, int], CellRecord]
    decoded_blocks: int
    total_blocks: int


class SetupScanStore:
    """SQLite-backed, compressed sidecar for one work dir's setup-scan
    results, keyed by ``(session_key, member_id)``.
    """

    def __init__(self, db_path: Path, *, block_rows: int = _DEFAULT_BLOCK_ROWS) -> None:
        if block_rows < 1:
            raise ValueError("block_rows must be positive")
        private_directory(db_path.parent)
        self._db_path = db_path
        self._block_rows = block_rows
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
        *,
        input_generation: int = 0,
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
                (
                    session_key,
                    self._member_key(member_id, input_generation),
                    now,
                    blob,
                ),
            )
        return SetupScanSidecar(
            session_key=session_key,
            member_id=member_id,
            created_at=now,
            blob_bytes=len(blob),
        )

    def get_result(
        self,
        session_key: str,
        member_id: str,
        *,
        input_generation: int = 0,
    ) -> dict[str, object] | None:
        """The full decoded result payload. Prefer ``get_sheet_profile`` for
        a single-sheet lookup -- this bounded scan result is small enough
        that decoding it whole stays cheap, but a targeted query avoids
        handing a caller the entire tree when it only needs one sheet.
        """
        with self._connect() as conn:
            row = conn.execute(
                "SELECT blob FROM setup_scan_blocks"
                " WHERE session_key = ? AND member_id = ?",
                (session_key, self._member_key(member_id, input_generation)),
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
        input_generation: int = 0,
    ) -> dict[str, object] | None:
        """Bounded post-profile query: one sheet's structural facts, without
        the caller needing to know the whole result's shape. ``side`` is
        ``"current"`` or ``"baseline"``.
        """
        result = self.get_result(
            session_key, member_id, input_generation=input_generation
        )
        if result is None:
            return None
        sheets = result.get(f"{side}_sheets")
        if not isinstance(sheets, list):
            return None
        for sheet in sheets:
            if isinstance(sheet, dict) and sheet.get("sheet_name") == sheet_name:
                return sheet
        return None

    @staticmethod
    def _sheet_key(sheet_name: str, input_generation: int) -> str:
        identity = f"{input_generation}:{sheet_name}"
        return hashlib.sha256(identity.encode("utf-8")).hexdigest()

    @staticmethod
    def _member_key(member_id: str, input_generation: int) -> str:
        return f"{input_generation}:{member_id}"

    @staticmethod
    def _compress(payload: object) -> bytes:
        encoded = json.dumps(
            payload,
            ensure_ascii=True,
            allow_nan=False,
            separators=(",", ":"),
        ).encode("utf-8")
        return zlib.compress(encoded, level=6)

    @staticmethod
    def _decompress(blob: bytes) -> object:
        try:
            decompressor = zlib.decompressobj()
            decoded = decompressor.decompress(
                blob, _MAX_DECOMPRESSED_BLOCK_BYTES + 1
            )
            if (
                len(decoded) > _MAX_DECOMPRESSED_BLOCK_BYTES
                or decompressor.unconsumed_tail
                or not decompressor.eof
            ):
                raise SetupSidecarCorruptError(
                    "setup sidecar block exceeds its bounded size"
                )
            return json.loads(decoded.decode("utf-8"))
        except SetupSidecarCorruptError:
            raise
        except (UnicodeDecodeError, json.JSONDecodeError, zlib.error) as exc:
            raise SetupSidecarCorruptError(
                "setup sidecar block is unreadable"
            ) from exc

    @staticmethod
    def _encode_value(value: CellValue) -> list[object]:
        if value is None:
            return ["none", None]
        if isinstance(value, bool):
            return ["bool", value]
        if isinstance(value, int):
            return ["int", value]
        if isinstance(value, float):
            if math.isfinite(value):
                return ["float", value]
            marker = "nan" if math.isnan(value) else "inf" if value > 0 else "-inf"
            return ["special_float", marker]
        if isinstance(value, dt.datetime):
            return ["datetime", value.isoformat()]
        if isinstance(value, dt.date):
            return ["date", value.isoformat()]
        return ["str", value]

    @staticmethod
    def _decode_value(payload: object) -> CellValue:
        if not isinstance(payload, list) or len(payload) != 2:
            raise SetupSidecarCorruptError("setup sidecar block is unreadable")
        kind, value = payload
        try:
            if kind == "none":
                return None
            if kind == "bool" and isinstance(value, bool):
                return value
            if kind == "int" and isinstance(value, int) and not isinstance(value, bool):
                return value
            if kind == "float" and isinstance(value, int | float):
                return float(value)
            if kind == "special_float" and value in {"nan", "inf", "-inf"}:
                return {
                    "nan": float("nan"),
                    "inf": float("inf"),
                    "-inf": float("-inf"),
                }[value]
            if kind == "date" and isinstance(value, str):
                return dt.date.fromisoformat(value)
            if kind == "datetime" and isinstance(value, str):
                return dt.datetime.fromisoformat(value)
            if kind == "str" and isinstance(value, str):
                return value
        except ValueError as exc:
            raise SetupSidecarCorruptError(
                "setup sidecar block is unreadable"
            ) from exc
        raise SetupSidecarCorruptError("setup sidecar block is unreadable")

    @classmethod
    def _encode_cell(cls, cell: CellRecord) -> list[object]:
        return [
            cell.row,
            cell.column,
            cls._encode_value(cell.value),
            cell.is_formula,
            cell.formula[:_MAX_FORMULA_TEXT_CHARS] if cell.formula else None,
            cell.number_format,
            cell.style_key,
        ]

    @classmethod
    def _decode_cell(cls, payload: object) -> CellRecord:
        if not isinstance(payload, list) or len(payload) != 7:
            raise SetupSidecarCorruptError("setup sidecar block is unreadable")
        row, column, value, is_formula, formula, number_format, style_key = payload
        if (
            not isinstance(row, int)
            or isinstance(row, bool)
            or row < 1
            or not isinstance(column, int)
            or isinstance(column, bool)
            or column < 1
            or is_formula not in {True, False, None}
            or (formula is not None and not isinstance(formula, str))
            or (number_format is not None and not isinstance(number_format, str))
            or (style_key is not None and not isinstance(style_key, str))
        ):
            raise SetupSidecarCorruptError("setup sidecar block is unreadable")
        return CellRecord(
            row=row,
            column=column,
            value=cls._decode_value(value),
            formula=formula,
            is_formula=is_formula,
            number_format=number_format,
            style_key=style_key,
        )

    def save_sheet(
        self,
        session_key: str,
        member_id: str,
        side: str,
        *,
        input_generation: int,
        source_hash: str,
        sheet: SheetSnapshot,
    ) -> SetupSheetInventory:
        """Replace one sheet with compressed sparse row blocks."""
        if side not in {"baseline", "current"}:
            raise ValueError("side must be 'baseline' or 'current'")
        sheet_key = self._sheet_key(sheet.name, input_generation)
        metadata = {
            "sheet_name": sheet.name,
            "visibility": sheet.visibility,
            "max_row": sheet.max_row,
            "max_column": sheet.max_column,
            "hidden_rows": sorted(sheet.hidden_rows),
            "hidden_columns": sorted(sheet.hidden_columns),
        }
        now = dt.datetime.now(dt.UTC).isoformat(timespec="milliseconds")
        block_count = 0
        with self._connect() as conn:
            conn.execute("BEGIN IMMEDIATE")
            conn.execute(
                """
                DELETE FROM setup_cell_blocks
                WHERE session_key = ? AND member_id = ? AND side = ? AND sheet_key = ?
                """,
                (session_key, member_id, side, sheet_key),
            )
            current_block_index: int | None = None
            current_cells: list[list[object]] = []

            def flush_block() -> None:
                nonlocal block_count
                if current_block_index is None or not current_cells:
                    return
                block_count += 1
                min_row = current_block_index * self._block_rows + 1
                max_row = min_row + self._block_rows - 1
                conn.execute(
                    """
                    INSERT INTO setup_cell_blocks (
                        session_key, member_id, side, sheet_key, block_index,
                        min_row, max_row, blob
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        session_key,
                        member_id,
                        side,
                        sheet_key,
                        current_block_index,
                        min_row,
                        max_row,
                        self._compress(current_cells),
                    ),
                )

            for coordinate in sorted(sheet.cells):
                cell = sheet.cells[coordinate]
                block_index = (cell.row - 1) // self._block_rows
                if (
                    current_block_index is not None
                    and block_index != current_block_index
                ):
                    flush_block()
                    current_cells = []
                current_block_index = block_index
                current_cells.append(self._encode_cell(cell))
            flush_block()
            conn.execute(
                """
                INSERT INTO setup_sheet_inventory (
                    session_key, member_id, side, sheet_key, input_generation,
                    source_hash, created_at, metadata_blob, block_count
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(session_key, member_id, side, sheet_key) DO UPDATE SET
                    input_generation = excluded.input_generation,
                    source_hash = excluded.source_hash,
                    created_at = excluded.created_at,
                    metadata_blob = excluded.metadata_blob,
                    block_count = excluded.block_count
                """,
                (
                    session_key,
                    member_id,
                    side,
                    sheet_key,
                    input_generation,
                    source_hash,
                    now,
                    self._compress(metadata),
                    block_count,
                ),
            )
        return SetupSheetInventory(
            sheet_name=sheet.name,
            visibility=sheet.visibility,
            max_row=sheet.max_row,
            max_column=sheet.max_column,
            hidden_rows=sheet.hidden_rows,
            hidden_columns=sheet.hidden_columns,
            block_count=block_count,
        )

    def _inventory_row(
        self,
        session_key: str,
        member_id: str,
        side: str,
        sheet_name: str,
        *,
        expected_generation: int,
        expected_source_hash: str,
    ) -> sqlite3.Row:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT * FROM setup_sheet_inventory
                WHERE session_key = ? AND member_id = ? AND side = ? AND sheet_key = ?
                """,
                (
                    session_key,
                    member_id,
                    side,
                    self._sheet_key(sheet_name, expected_generation),
                ),
            ).fetchone()
        if row is None:
            with self._connect() as conn:
                candidates = conn.execute(
                    """
                    SELECT metadata_blob FROM setup_sheet_inventory
                    WHERE session_key = ? AND member_id = ? AND side = ?
                    """,
                    (session_key, member_id, side),
                ).fetchall()
            for candidate in candidates:
                metadata = self._decompress(bytes(candidate["metadata_blob"]))
                if isinstance(metadata, dict) and metadata.get("sheet_name") == sheet_name:
                    raise SetupSidecarStaleError(
                        "setup sidecar does not match the current input generation"
                    )
            raise KeyError("setup sidecar sheet is unavailable")
        if (
            int(row["input_generation"]) != expected_generation
            or str(row["source_hash"]) != expected_source_hash
        ):
            raise SetupSidecarStaleError(
                "setup sidecar does not match the current input generation"
            )
        return row

    @classmethod
    def _inventory_from_row(cls, row: sqlite3.Row) -> SetupSheetInventory:
        metadata = cls._decompress(bytes(row["metadata_blob"]))
        if not isinstance(metadata, dict):
            raise SetupSidecarCorruptError("setup sidecar block is unreadable")
        try:
            sheet_name = metadata["sheet_name"]
            visibility = metadata["visibility"]
            max_row = metadata["max_row"]
            max_column = metadata["max_column"]
            hidden_rows = metadata["hidden_rows"]
            hidden_columns = metadata["hidden_columns"]
            if (
                not isinstance(sheet_name, str)
                or not isinstance(visibility, str)
                or not isinstance(max_row, int)
                or not isinstance(max_column, int)
                or not isinstance(hidden_rows, list)
                or not isinstance(hidden_columns, list)
            ):
                raise TypeError
            return SetupSheetInventory(
                sheet_name=sheet_name,
                visibility=visibility,
                max_row=max_row,
                max_column=max_column,
                hidden_rows=frozenset(int(value) for value in hidden_rows),
                hidden_columns=frozenset(int(value) for value in hidden_columns),
                block_count=int(row["block_count"]),
            )
        except (KeyError, TypeError, ValueError) as exc:
            raise SetupSidecarCorruptError(
                "setup sidecar block is unreadable"
            ) from exc

    def list_sheets(
        self,
        session_key: str,
        member_id: str,
        side: str,
        *,
        expected_generation: int | None = None,
        expected_source_hash: str | None = None,
    ) -> tuple[SetupSheetInventory, ...]:
        with self._connect() as conn:
            if expected_generation is None:
                rows = conn.execute(
                    """
                    SELECT * FROM setup_sheet_inventory
                    WHERE session_key = ? AND member_id = ? AND side = ?
                    ORDER BY rowid
                    """,
                    (session_key, member_id, side),
                ).fetchall()
            else:
                rows = conn.execute(
                    """
                    SELECT * FROM setup_sheet_inventory
                    WHERE session_key = ? AND member_id = ? AND side = ?
                        AND input_generation = ? AND source_hash = ?
                    ORDER BY rowid
                    """,
                    (
                        session_key,
                        member_id,
                        side,
                        expected_generation,
                        expected_source_hash or "",
                    ),
                ).fetchall()
        return tuple(self._inventory_from_row(row) for row in rows)

    def get_sheet_inventory(
        self,
        session_key: str,
        member_id: str,
        side: str,
        sheet_name: str,
        *,
        expected_generation: int,
        expected_source_hash: str,
    ) -> SetupSheetInventory:
        row = self._inventory_row(
            session_key,
            member_id,
            side,
            sheet_name,
            expected_generation=expected_generation,
            expected_source_hash=expected_source_hash,
        )
        return self._inventory_from_row(row)

    def _query_cells(
        self,
        session_key: str,
        member_id: str,
        side: str,
        sheet_name: str,
        *,
        input_generation: int,
        min_row: int,
        max_row: int,
        min_col: int | None,
        max_col: int | None,
        columns: frozenset[int] | None,
    ) -> tuple[dict[tuple[int, int], CellRecord], int]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT blob FROM setup_cell_blocks
                WHERE session_key = ? AND member_id = ? AND side = ?
                    AND sheet_key = ? AND max_row >= ? AND min_row <= ?
                ORDER BY block_index
                """,
                (
                    session_key,
                    member_id,
                    side,
                    self._sheet_key(sheet_name, input_generation),
                    min_row,
                    max_row,
                ),
            ).fetchall()
        cells: dict[tuple[int, int], CellRecord] = {}
        for row in rows:
            payload = self._decompress(bytes(row["blob"]))
            if not isinstance(payload, list):
                raise SetupSidecarCorruptError("setup sidecar block is unreadable")
            for raw_cell in payload:
                cell = self._decode_cell(raw_cell)
                if not min_row <= cell.row <= max_row:
                    continue
                if columns is not None and cell.column not in columns:
                    continue
                if min_col is not None and cell.column < min_col:
                    continue
                if max_col is not None and cell.column > max_col:
                    continue
                cells[(cell.row, cell.column)] = cell
        return cells, len(rows)

    def query_window(
        self,
        session_key: str,
        member_id: str,
        side: str,
        sheet_name: str,
        *,
        expected_generation: int,
        expected_source_hash: str,
        min_row: int,
        max_row: int,
        min_col: int,
        max_col: int,
    ) -> SetupWindow:
        inventory_row = self._inventory_row(
            session_key,
            member_id,
            side,
            sheet_name,
            expected_generation=expected_generation,
            expected_source_hash=expected_source_hash,
        )
        cells, decoded_blocks = self._query_cells(
            session_key,
            member_id,
            side,
            sheet_name,
            input_generation=expected_generation,
            min_row=max(1, min_row),
            max_row=max(1, max_row),
            min_col=max(1, min_col),
            max_col=max(1, max_col),
            columns=None,
        )
        return SetupWindow(
            cells=cells,
            decoded_blocks=decoded_blocks,
            total_blocks=int(inventory_row["block_count"]),
        )

    def query_cells(
        self,
        session_key: str,
        member_id: str,
        side: str,
        sheet_name: str,
        *,
        expected_generation: int,
        expected_source_hash: str,
        min_row: int,
        max_row: int,
        columns: frozenset[int],
    ) -> dict[tuple[int, int], CellRecord]:
        self._inventory_row(
            session_key,
            member_id,
            side,
            sheet_name,
            expected_generation=expected_generation,
            expected_source_hash=expected_source_hash,
        )
        cells, _decoded_blocks = self._query_cells(
            session_key,
            member_id,
            side,
            sheet_name,
            input_generation=expected_generation,
            min_row=max(1, min_row),
            max_row=max(1, max_row),
            min_col=None,
            max_col=None,
            columns=columns,
        )
        return cells

    def load_sheet(
        self,
        session_key: str,
        member_id: str,
        side: str,
        sheet_name: str,
        *,
        expected_generation: int,
        expected_source_hash: str,
    ) -> SheetSnapshot:
        row = self._inventory_row(
            session_key,
            member_id,
            side,
            sheet_name,
            expected_generation=expected_generation,
            expected_source_hash=expected_source_hash,
        )
        inventory = self._inventory_from_row(row)
        cells, _decoded_blocks = self._query_cells(
            session_key,
            member_id,
            side,
            sheet_name,
            input_generation=expected_generation,
            min_row=1,
            max_row=max(1, inventory.max_row),
            min_col=None,
            max_col=None,
            columns=None,
        )
        return SheetSnapshot(
            name=inventory.sheet_name,
            visibility=inventory.visibility,
            max_row=inventory.max_row,
            max_column=inventory.max_column,
            cells=cells,
            hidden_rows=inventory.hidden_rows,
            hidden_columns=inventory.hidden_columns,
        )

    def delete(self, session_key: str) -> None:
        with self._connect() as conn:
            conn.execute(
                "DELETE FROM setup_cell_blocks WHERE session_key = ?", (session_key,)
            )
            conn.execute(
                "DELETE FROM setup_sheet_inventory WHERE session_key = ?",
                (session_key,),
            )
            conn.execute(
                "DELETE FROM setup_scan_blocks WHERE session_key = ?", (session_key,)
            )

    def delete_member(
        self,
        session_key: str,
        member_id: str,
        *,
        input_generation: int | None = None,
    ) -> None:
        """Delete one member before replacing its input generation."""
        with self._connect() as conn:
            generation_filter = (
                " AND input_generation = ?" if input_generation is not None else ""
            )
            parameters: tuple[object, ...] = (
                (session_key, member_id, input_generation)
                if input_generation is not None
                else (session_key, member_id)
            )
            sheet_keys = tuple(
                str(row[0])
                for row in conn.execute(
                    "SELECT sheet_key FROM setup_sheet_inventory "
                    "WHERE session_key = ? AND member_id = ?" + generation_filter,
                    parameters,
                )
            )
            for sheet_key in sheet_keys:
                conn.execute(
                    """
                    DELETE FROM setup_cell_blocks
                    WHERE session_key = ? AND member_id = ? AND sheet_key = ?
                    """,
                    (session_key, member_id, sheet_key),
                )
            conn.execute(
                "DELETE FROM setup_sheet_inventory WHERE session_key = ? "
                "AND member_id = ?" + generation_filter,
                parameters,
            )
            conn.execute(
                "DELETE FROM setup_scan_blocks WHERE session_key = ? AND member_id = ?",
                (
                    session_key,
                    self._member_key(member_id, input_generation or 0),
                ),
            )

    def delete_generation(self, session_key: str, input_generation: int) -> int:
        """Delete one superseded generation without touching newer rows."""
        with self._connect() as conn:
            stale_sheets = conn.execute(
                """
                SELECT member_id, side, sheet_key
                FROM setup_sheet_inventory
                WHERE session_key = ? AND input_generation = ?
                """,
                (session_key, input_generation),
            ).fetchall()
            for row in stale_sheets:
                conn.execute(
                    """
                    DELETE FROM setup_cell_blocks
                    WHERE session_key = ? AND member_id = ?
                        AND side = ? AND sheet_key = ?
                    """,
                    (
                        session_key,
                        row["member_id"],
                        row["side"],
                        row["sheet_key"],
                    ),
                )
            sheet_cursor = conn.execute(
                """
                DELETE FROM setup_sheet_inventory
                WHERE session_key = ? AND input_generation = ?
                """,
                (session_key, input_generation),
            )
            result_rows = conn.execute(
                "SELECT member_id FROM setup_scan_blocks WHERE session_key = ?",
                (session_key,),
            ).fetchall()
            removed_results = 0
            prefix = f"{input_generation}:"
            for row in result_rows:
                stored_member_id = str(row["member_id"])
                if not stored_member_id.startswith(prefix):
                    continue
                cursor = conn.execute(
                    """
                    DELETE FROM setup_scan_blocks
                    WHERE session_key = ? AND member_id = ?
                    """,
                    (session_key, stored_member_id),
                )
                removed_results += cursor.rowcount
            return sheet_cursor.rowcount + removed_results

    def delete_stale(self, older_than: dt.datetime) -> int:
        """Bounded cleanup: remove scan results untouched since ``older_than``."""
        cutoff = older_than.isoformat(timespec="milliseconds")
        with self._connect() as conn:
            stale_sheets = conn.execute(
                """
                SELECT session_key, member_id, side, sheet_key
                FROM setup_sheet_inventory
                WHERE created_at < ?
                """,
                (cutoff,),
            ).fetchall()
            for row in stale_sheets:
                conn.execute(
                    """
                    DELETE FROM setup_cell_blocks
                    WHERE session_key = ? AND member_id = ?
                        AND side = ? AND sheet_key = ?
                    """,
                    (
                        row["session_key"],
                        row["member_id"],
                        row["side"],
                        row["sheet_key"],
                    ),
                )
                conn.execute(
                    """
                    DELETE FROM setup_sheet_inventory
                    WHERE session_key = ? AND member_id = ?
                        AND side = ? AND sheet_key = ?
                    """,
                    (
                        row["session_key"],
                        row["member_id"],
                        row["side"],
                        row["sheet_key"],
                    ),
                )
            result_cursor = conn.execute(
                "DELETE FROM setup_scan_blocks WHERE created_at < ?",
                (cutoff,),
            )
            return len(stale_sheets) + result_cursor.rowcount
