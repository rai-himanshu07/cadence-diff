"""Private, content-addressed cache for external XLSB formula extraction.

Skips repeat external-engine work (LibreOffice/Excel conversion, or the
native BIFF12 kernel decode) plus the formula walk -- never cached-value
loading or XLSB style typing, both of which stay authoritative from the
direct ``pyxlsb`` read on every run.

An entry is keyed by the decrypted package's own SHA-256, a digest/count of
its scanned formula coordinates, this module's cache-schema and
extraction-contract versions, and the adapter family/fingerprint that
produced it -- so any schema, adapter, or content change is always a miss,
never a stale hit. A hit is always passed through the same
``validate_formula_extraction`` the direct extraction path already runs
before any snapshot mutation; validation failure quarantines the entry and
falls back to a normal extraction, changing no QC result.

Entries are immutable, block-compressed (reusing
``qc_tool.findings_store.encode_block``/``decode_block``), written to a
private temporary file, and atomically published via ``os.replace``. A race
between two writers for the same (deterministic) key is harmless -- both
would write byte-identical content. Bounded by entry count, total bytes, and
per-entry bytes; an over-cap entry is never admitted, and LRU eviction never
follows a symlink or touches a path outside the managed directory.

No cached content -- formula text, defined-name text/target, sheet name, or
source path -- is ever logged; only counts, sizes, and hex digests.
"""

from __future__ import annotations

import contextlib
import hashlib
import json
import logging
import os
import struct
import tempfile
import time
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import IO, Literal

from qc_tool.findings_store import FindingsStoreError, decode_block, encode_block
from qc_tool.io.formula_enrichment import (
    ExtractedDefinedName,
    FormulaEnrichmentError,
    FormulaExtraction,
    bounded_defined_names,
    validate_formula_extraction,
)
from qc_tool.io.xlsb_formula import XlsbFormulaScan
from qc_tool.security import private_directory, private_file

logger = logging.getLogger(__name__)

#: Bump when the on-disk entry envelope (magic/blocks/manifest shape) changes.
#: v2 (from v1): blocks now carry a `kind` (`"a1"`/`"r1c1"`/`"name"`) so a
#: canonical R1C1 map can be stored and reconstructed distinctly from A1 text
#: instead of always being dropped on read. Bumping this changes the key
#: digest, so every v1 entry becomes a deterministic miss -- never silently
#: reinterpreted as v2.
CACHE_SCHEMA_VERSION = 2
#: Bump when `FormulaExtraction`'s meaning changes in a way that would make an
#: old entry mean something different, independent of the file envelope.
EXTRACTION_CONTRACT_VERSION = 1

_MAGIC = b"QFXC1\n"
_FOOTER_STRUCT = struct.Struct("<Q")
_ENTRY_SUFFIX = ".qfxc"
#: A sheet's (or the defined-name collection's) rows are chunked this large so
#: one entry never holds a single monolithic in-memory block.
_ENTRY_BLOCK_ROWS = 20_000

DEFAULT_MAX_ENTRIES = 8
DEFAULT_MAX_TOTAL_BYTES = 1 * 1024 * 1024 * 1024
DEFAULT_MAX_ENTRY_BYTES = 512 * 1024 * 1024

Extractor = Callable[[bytes, XlsbFormulaScan], FormulaExtraction]


class FormulaCacheError(RuntimeError):
    """A cache entry could not be trusted; the caller must extract normally."""


@dataclass(frozen=True, slots=True)
class FormulaCacheKey:
    """Every field that must match for a cached extraction to be reusable."""

    package_sha256: str
    coordinate_digest: str
    coordinate_count: int
    adapter_family: str
    adapter_fingerprint: str
    schema_version: int = CACHE_SCHEMA_VERSION
    contract_version: int = EXTRACTION_CONTRACT_VERSION

    def digest(self) -> str:
        payload = "|".join(
            (
                str(self.schema_version),
                str(self.contract_version),
                self.adapter_family,
                self.adapter_fingerprint,
                self.package_sha256,
                self.coordinate_digest,
                str(self.coordinate_count),
            )
        )
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def coordinate_digest(scan: XlsbFormulaScan) -> str:
    """Order-independent digest of scanned formula coordinates -- never text."""
    parts = sorted(
        f"{sheet}\0{row}\0{column}"
        for sheet, cells in scan.formula_cells.items()
        for row, column in cells
    )
    digest = hashlib.sha256()
    for part in parts:
        digest.update(part.encode("utf-8"))
        digest.update(b"\n")
    return digest.hexdigest()


def package_digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def discover_adapter_fingerprint(
    engine: Literal["native", "excel", "libreoffice"],
) -> tuple[str, str] | None:
    """``(adapter_family, fingerprint)`` for the already-resolved ``engine``,
    without opening any workbook, or None.

    ``engine`` names which adapter will actually run (resolved from
    `formula_engine`'s ``native|excel|libreoffice|auto`` upstream -- ``auto``
    is never passed here); the cache key must reflect the real engine choice,
    not a `sys.platform` guess, so switching engines always misses instead of
    risking a cross-engine hit. ``None`` means the fingerprint cannot be
    trusted right now (adapter missing, version probe failed); caching is
    disabled for this call rather than guessed.
    """
    if engine == "native":
        from qc_tool.io.native_formula import native_adapter_fingerprint

        fingerprint = native_adapter_fingerprint()
        return ("native", fingerprint) if fingerprint else None
    if engine == "libreoffice":
        from qc_tool.io.libreoffice_formula import libreoffice_adapter_fingerprint

        fingerprint = libreoffice_adapter_fingerprint()
        return ("libreoffice", fingerprint) if fingerprint else None
    if engine == "excel":
        from qc_tool.io.excel_formula import excel_adapter_fingerprint

        fingerprint = excel_adapter_fingerprint()
        return ("excel", fingerprint) if fingerprint else None
    return None


def _entry_blocks(
    extraction: FormulaExtraction,
) -> list[tuple[str, str | None, list[list[object]]]]:
    """``(kind, sheet_name_or_None, bounded_rows)`` block plan for one entry.

    ``kind`` is ``"a1"``, ``"r1c1"``, or ``"name"`` so the reader can
    reconstruct each of `FormulaExtraction`'s independent maps distinctly
    instead of guessing from shape alone -- schema v1's collapse of A1 and
    R1C1 into one untagged shape is exactly what silently dropped
    `formulas_r1c1` on every cache hit.
    """
    blocks: list[tuple[str, str | None, list[list[object]]]] = []
    for sheet, cells in sorted(extraction.formulas.items()):
        rows = [
            [row, column, formula] for (row, column), formula in sorted(cells.items())
        ]
        for start in range(0, len(rows), _ENTRY_BLOCK_ROWS):
            blocks.append(("a1", sheet, rows[start : start + _ENTRY_BLOCK_ROWS]))
    if extraction.formulas_r1c1 is not None:
        for sheet, cells in sorted(extraction.formulas_r1c1.items()):
            rows = [
                [row, column, formula]
                for (row, column), formula in sorted(cells.items())
            ]
            for start in range(0, len(rows), _ENTRY_BLOCK_ROWS):
                blocks.append(("r1c1", sheet, rows[start : start + _ENTRY_BLOCK_ROWS]))
    name_rows = [
        [item.name, item.target, item.sheet, item.hidden]
        for item in extraction.defined_names
    ]
    for start in range(0, len(name_rows), _ENTRY_BLOCK_ROWS):
        blocks.append(("name", None, name_rows[start : start + _ENTRY_BLOCK_ROWS]))
    return blocks


def _write_entry(handle: IO[bytes], key: FormulaCacheKey, extraction: FormulaExtraction) -> int:
    handle.write(_MAGIC)
    block_infos: list[dict[str, object]] = []
    for kind, sheet, rows in _entry_blocks(extraction):
        blob = encode_block(rows)
        offset = handle.tell()
        handle.write(blob)
        block_infos.append(
            {
                "kind": kind,
                "sheet": sheet,
                "offset": offset,
                "length": len(blob),
                "count": len(rows),
            }
        )
    manifest = {
        "schema_version": key.schema_version,
        "contract_version": key.contract_version,
        "adapter_family": key.adapter_family,
        "adapter_fingerprint": key.adapter_fingerprint,
        "package_sha256": key.package_sha256,
        "coordinate_digest": key.coordinate_digest,
        "coordinate_count": key.coordinate_count,
        "engine": extraction.engine,
        "detail": extraction.detail,
        "defined_names_complete": extraction.defined_names_complete,
        # Distinguishes "this adapter never produces R1C1" (formulas_r1c1 is
        # None, e.g. Excel/LibreOffice) from "it produced an empty map" (e.g.
        # native on a workbook with formula cells nowhere r1c1 could be
        # validated) -- both round-trip to zero r1c1 blocks otherwise.
        "has_r1c1": extraction.formulas_r1c1 is not None,
        "blocks": block_infos,
        "written_at": round(time.time(), 3),
    }
    manifest_bytes = json.dumps(manifest, sort_keys=True).encode("utf-8")
    handle.write(manifest_bytes)
    handle.write(_FOOTER_STRUCT.pack(len(manifest_bytes)))
    handle.flush()
    os.fsync(handle.fileno())
    return handle.tell()


def _read_manifest(handle: IO[bytes]) -> dict[str, object]:
    handle.seek(0)
    magic = handle.read(len(_MAGIC))
    if magic != _MAGIC:
        raise FormulaCacheError("not a formula-cache entry")
    handle.seek(-_FOOTER_STRUCT.size, os.SEEK_END)
    (manifest_length,) = _FOOTER_STRUCT.unpack(handle.read(_FOOTER_STRUCT.size))
    handle.seek(-_FOOTER_STRUCT.size - manifest_length, os.SEEK_END)
    try:
        manifest = json.loads(handle.read(manifest_length).decode("utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise FormulaCacheError(f"corrupt manifest: {exc}") from exc
    if not isinstance(manifest, dict):
        raise FormulaCacheError("manifest is not an object")
    return manifest


def _read_entry(path: Path, key: FormulaCacheKey) -> FormulaExtraction:
    with path.open("rb") as handle:
        manifest = _read_manifest(handle)
        if manifest.get("schema_version") != CACHE_SCHEMA_VERSION:
            raise FormulaCacheError("unknown cache schema version")
        if manifest.get("contract_version") != EXTRACTION_CONTRACT_VERSION:
            raise FormulaCacheError("unknown extraction contract version")
        expectations = (
            ("adapter_family", key.adapter_family),
            ("adapter_fingerprint", key.adapter_fingerprint),
            ("package_sha256", key.package_sha256),
            ("coordinate_digest", key.coordinate_digest),
            ("coordinate_count", key.coordinate_count),
        )
        for field_name, expected in expectations:
            if manifest.get(field_name) != expected:
                raise FormulaCacheError(f"{field_name} does not match the requested key")
        blocks = manifest.get("blocks")
        if not isinstance(blocks, list):
            raise FormulaCacheError("manifest blocks are missing")
        formulas: dict[str, dict[tuple[int, int], str]] = {}
        formulas_r1c1: dict[str, dict[tuple[int, int], str]] = {}
        names: list[ExtractedDefinedName] = []
        for block in blocks:
            if not isinstance(block, dict):
                raise FormulaCacheError("manifest block is malformed")
            offset, length, count = (
                block.get("offset"),
                block.get("length"),
                block.get("count"),
            )
            if not (
                isinstance(offset, int)
                and isinstance(length, int)
                and isinstance(count, int)
            ):
                raise FormulaCacheError("manifest block metadata is malformed")
            handle.seek(offset)
            blob = handle.read(length)
            if len(blob) != length:
                raise FormulaCacheError("truncated cache block")
            try:
                rows = decode_block(blob)
            except FindingsStoreError as exc:
                raise FormulaCacheError(f"corrupt cache block: {exc}") from exc
            kind = block.get("kind")
            if kind == "name":
                for row in rows:
                    if not (isinstance(row, list) and len(row) == 4):
                        raise FormulaCacheError("malformed defined-name row")
                    name, target, name_sheet, hidden = row
                    names.append(
                        ExtractedDefinedName(
                            name=str(name),
                            target=str(target),
                            sheet=(str(name_sheet) if name_sheet is not None else None),
                            hidden=bool(hidden),
                        )
                    )
                continue
            sheet = block.get("sheet")
            if not isinstance(sheet, str):
                raise FormulaCacheError("manifest block sheet is malformed")
            if kind == "a1":
                target_map = formulas
            elif kind == "r1c1":
                target_map = formulas_r1c1
            else:
                raise FormulaCacheError("manifest block kind is unrecognized")
            sheet_map = target_map.setdefault(sheet, {})
            for row in rows:
                if not (isinstance(row, list) and len(row) == 3):
                    raise FormulaCacheError("malformed formula row")
                cell_row, cell_column, formula = row
                if not (
                    isinstance(cell_row, int)
                    and isinstance(cell_column, int)
                    and isinstance(formula, str)
                ):
                    raise FormulaCacheError("malformed formula row types")
                sheet_map[(cell_row, cell_column)] = formula
    has_r1c1 = bool(manifest.get("has_r1c1"))
    if has_r1c1:
        a1_coordinates = {
            (sheet, row, column)
            for sheet, cells in formulas.items()
            for row, column in cells
        }
        r1c1_coordinates = {
            (sheet, row, column)
            for sheet, cells in formulas_r1c1.items()
            for row, column in cells
        }
        if not r1c1_coordinates <= a1_coordinates:
            raise FormulaCacheError("r1c1 coordinates are not a subset of a1 coordinates")
    names_tuple, names_complete = bounded_defined_names(names)
    return FormulaExtraction(
        formulas=formulas,
        engine=str(manifest.get("engine", "")),
        detail=str(manifest.get("detail", "")),
        defined_names=names_tuple,
        defined_names_complete=names_complete and bool(manifest.get("defined_names_complete")),
        formulas_r1c1=(formulas_r1c1 if has_r1c1 else None),
    )


def _quarantine(path: Path) -> None:
    """A corrupt private cache entry has no value to preserve -- remove it."""
    with contextlib.suppress(OSError):
        path.unlink()


class FormulaExtractionCache:
    """Bounded, private, content-addressed cache of workbook formula text."""

    def __init__(
        self,
        root: Path,
        *,
        max_entries: int = DEFAULT_MAX_ENTRIES,
        max_total_bytes: int = DEFAULT_MAX_TOTAL_BYTES,
        max_entry_bytes: int = DEFAULT_MAX_ENTRY_BYTES,
    ) -> None:
        self.root = root
        self.max_entries = max_entries
        self.max_total_bytes = max_total_bytes
        self.max_entry_bytes = max_entry_bytes

    def _entry_path(self, key: FormulaCacheKey) -> Path:
        return self.root / f"{key.digest()}{_ENTRY_SUFFIX}"

    def lookup(self, key: FormulaCacheKey) -> FormulaExtraction | None:
        path = self._entry_path(key)
        if path.is_symlink() or not path.is_file():
            return None
        try:
            extraction = _read_entry(path, key)
        except (FormulaCacheError, OSError) as exc:
            # Never log `exc`'s own text -- it may echo a manifest field name
            # or filesystem path. Only the fixed reason and exception class
            # are safe to log.
            logger.warning(
                "formula-cache entry unusable (%s), quarantining",
                type(exc).__name__,
            )
            _quarantine(path)
            return None
        with contextlib.suppress(OSError):
            os.utime(path, None)  # LRU: a hit is a fresh access
        return extraction

    def store(self, key: FormulaCacheKey, extraction: FormulaExtraction) -> None:
        private_directory(self.root)
        tmp_path: Path | None = None
        try:
            fd, tmp_name = tempfile.mkstemp(
                dir=self.root, prefix=".tmp-", suffix=_ENTRY_SUFFIX
            )
            tmp_path = Path(tmp_name)
            with os.fdopen(fd, "wb") as handle:
                size = _write_entry(handle, key, extraction)
        except OSError as exc:
            # Never log `exc`'s own text -- OSError messages embed the
            # filesystem path involved. Only the fixed reason and exception
            # class are safe to log.
            logger.warning("formula-cache store failed (%s)", type(exc).__name__)
            if tmp_path is not None:
                with contextlib.suppress(OSError):
                    tmp_path.unlink()
            return
        if size > self.max_entry_bytes:
            with contextlib.suppress(OSError):
                tmp_path.unlink()
            return
        private_file(tmp_path)
        # Atomic publish; a concurrent identical writer racing this rename is
        # harmless because both would produce byte-identical content for the
        # same immutable key.
        os.replace(tmp_path, self._entry_path(key))
        self._evict_over_cap()

    def status(self) -> dict[str, int]:
        entries = self._entries()
        return {
            "entry_count": len(entries),
            "total_bytes": sum(size for _, size, _ in entries),
        }

    def clear(self) -> int:
        entries = self._entries()
        for path, _, _ in entries:
            with contextlib.suppress(OSError):
                path.unlink()
        return len(entries)

    def _entries(self) -> list[tuple[Path, int, float]]:
        if not self.root.is_dir():
            return []
        out: list[tuple[Path, int, float]] = []
        for path in self.root.iterdir():
            if path.is_symlink() or not path.is_file():
                continue
            if path.suffix != _ENTRY_SUFFIX or path.name.startswith(".tmp-"):
                continue
            try:
                stat = path.stat()
            except OSError:
                continue
            out.append((path, stat.st_size, stat.st_atime))
        return out

    def _evict_over_cap(self) -> None:
        entries = self._entries()
        entries.sort(key=lambda item: item[2])  # oldest access first
        total = sum(size for _, size, _ in entries)
        while entries and (
            len(entries) > self.max_entries or total > self.max_total_bytes
        ):
            path, size, _ = entries.pop(0)
            with contextlib.suppress(OSError):
                path.unlink()
            total -= size


def extract_with_cache(
    data: bytes,
    scan: XlsbFormulaScan,
    *,
    cache: FormulaExtractionCache | None,
    engine: Literal["native", "excel", "libreoffice"],
    extractor: Extractor,
) -> FormulaExtraction:
    """Try a cache hit first; always fall back to ``extractor`` on any miss.

    ``engine`` is the already-resolved adapter ``extractor`` will actually
    run (never ``"auto"``) -- it keys the cache entry so switching engines
    always misses. A cache hit, miss, corrupt-entry fallback, or ``cache is
    None`` (disabled) all reach the identical ``extractor`` output for a
    genuine miss and the identical validated extraction for a hit -- callers
    see no difference in the returned `FormulaExtraction` shape.
    """
    if cache is None:
        return extractor(data, scan)
    fingerprint = discover_adapter_fingerprint(engine)
    if fingerprint is None:
        return extractor(data, scan)
    adapter_family, adapter_fingerprint = fingerprint
    key = FormulaCacheKey(
        package_sha256=package_digest(data),
        coordinate_digest=coordinate_digest(scan),
        coordinate_count=scan.formula_count,
        adapter_family=adapter_family,
        adapter_fingerprint=adapter_fingerprint,
    )
    cached = cache.lookup(key)
    if cached is not None:
        try:
            validate_formula_extraction(scan, cached)
        except FormulaEnrichmentError:
            # Never log the exception's own text: `validate_formula_extraction`
            # may embed a real sheet name and coordinate in its message (for
            # example "Sheet1!A1: ..."). Only this fixed reason is safe to log.
            logger.warning(
                "formula-cache hit failed live validation; quarantining entry"
            )
            _quarantine(cache._entry_path(key))
        else:
            return cached
    extraction = extractor(data, scan)
    cache.store(key, extraction)
    return extraction
