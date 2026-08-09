"""Block-compressed findings storage with lazy sequence access.

Blocks are compact JSONL compressed once with stdlib zlib (measured 17x on
real findings). The same blobs serve the in-run spill, the history table, and
paged UI reads; digests are always computed over uncompressed payloads.
"""

from __future__ import annotations

import heapq
import json
import struct
import zlib
from collections import OrderedDict
from collections.abc import Callable, Iterable, Iterator, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import IO, Any, Protocol

from qc_tool.findings import Finding

try:  # parse-side only: encoded bytes are digest-pinned to stdlib json
    from orjson import loads as _json_loads
except ImportError:  # pragma: no cover - orjson ships in the project env
    _json_loads = json.loads

BLOCK_FINDINGS = 5_000
_ZLIB_LEVEL = 6
_MAGIC = b"QCFB1\n"
_FOOTER_STRUCT = struct.Struct("<Q")
_MERGE_FAN_IN = 24


class FindingsStoreError(RuntimeError):
    """Typed failure for corrupted or inconsistent block storage."""


def finding_payload(finding: Finding) -> dict[str, Any]:
    """Serialize one finding, preserving the private excluded fields.

    ``model_dump`` drops ``exclude=True`` fields (focus shape ids, the
    counterfactual basis, the series anchor), but history sidecars and focus
    targets read them from the in-memory result, so a spill round-trip must
    carry them. ``model_validate`` accepts them back as regular fields.
    """
    payload = finding.model_dump(mode="json")
    if finding.focus_shape_id is not None:
        payload["focus_shape_id"] = finding.focus_shape_id
    if finding.baseline_focus_shape_id is not None:
        payload["baseline_focus_shape_id"] = finding.baseline_focus_shape_id
    if finding.counterfactual_basis is not None:
        payload["counterfactual_basis"] = finding.counterfactual_basis.model_dump(
            mode="json"
        )
    if finding.series_anchor is not None:
        payload["series_anchor"] = finding.series_anchor.model_dump(mode="json")
    return payload


_PRIVATE_PAYLOAD_KEYS = frozenset(
    {
        "focus_shape_id",
        "baseline_focus_shape_id",
        "counterfactual_basis",
        "series_anchor",
    }
)


def public_payload(payload: dict[str, Any]) -> dict[str, Any]:
    """The stored-history projection: a spill payload minus private fields.

    History keeps counterfactual bases, series anchors, and focus locators in
    their own sidecar tables; the findings blocks must match the legacy
    ``runs.findings`` JSON, which never carried them.
    """
    return {
        key: value
        for key, value in payload.items()
        if key not in _PRIVATE_PAYLOAD_KEYS
    }


def finding_ordinal(finding_id: str) -> int:
    """Position encoded by a triaged ``F%04d`` id; 0 when not parseable."""
    if len(finding_id) < 2 or not finding_id.startswith("F"):
        return 0
    try:
        return int(finding_id[1:])
    except ValueError:
        return 0


def finding_by_id(
    findings: Sequence[Finding], finding_id: str
) -> Finding | None:
    """O(1) lookup exploiting triage's id-equals-position contract.

    Falls back to a scan when ids are not the standard sequential form.
    """
    ordinal = finding_ordinal(finding_id)
    if 1 <= ordinal <= len(findings):
        candidate = findings[ordinal - 1]
        if candidate.finding_id == finding_id:
            return candidate
    for candidate in findings:
        if candidate.finding_id == finding_id:
            return candidate
    return None


def encode_block(rows: Sequence[object], *, level: int = _ZLIB_LEVEL) -> bytes:
    lines = "\n".join(json.dumps(row, separators=(",", ":")) for row in rows)
    return zlib.compress(lines.encode("utf-8"), level)


def decode_block(blob: bytes) -> list[object]:
    try:
        text = zlib.decompress(blob)
    except zlib.error as exc:
        raise FindingsStoreError(f"block payload is corrupted: {exc}") from exc
    try:
        return [_json_loads(line) for line in text.splitlines() if line]
    except (json.JSONDecodeError, ValueError, UnicodeDecodeError) as exc:
        raise FindingsStoreError(f"block row is not valid JSON: {exc}") from exc


@dataclass(frozen=True, slots=True)
class BlockInfo:
    offset: int
    length: int
    count: int


class BlockSource(Protocol):
    """Read-side contract shared by spill files and history rows."""

    def block_infos(self) -> Sequence[BlockInfo]: ...

    def read_block(self, index: int) -> bytes: ...


class BlockFile:
    """Ordered block container: concatenated blobs plus a JSON footer."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self._infos: list[BlockInfo] = []
        self._handle: IO[bytes] | None = None

    # -- writing ---------------------------------------------------------
    def __enter__(self) -> BlockFile:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._handle = self.path.open("wb")
        self._handle.write(_MAGIC)
        return self

    def append_rows(self, rows: Sequence[object]) -> None:
        if self._handle is None:
            raise FindingsStoreError("block file is not open for writing")
        if not rows:
            return
        blob = encode_block(rows)
        offset = self._handle.tell()
        self._handle.write(blob)
        self._infos.append(BlockInfo(offset=offset, length=len(blob), count=len(rows)))

    def __exit__(self, *exc_info: object) -> None:
        if self._handle is None:
            return
        footer = json.dumps(
            [[info.offset, info.length, info.count] for info in self._infos],
            separators=(",", ":"),
        ).encode("utf-8")
        self._handle.write(footer)
        self._handle.write(_FOOTER_STRUCT.pack(len(footer)))
        self._handle.close()
        self._handle = None

    # -- reading ---------------------------------------------------------
    @classmethod
    def open(cls, path: Path) -> BlockFile:
        source = cls(path)
        with path.open("rb") as handle:
            magic = handle.read(len(_MAGIC))
            if magic != _MAGIC:
                raise FindingsStoreError(f"{path.name}: not a findings block file")
            handle.seek(-_FOOTER_STRUCT.size, 2)
            (footer_length,) = _FOOTER_STRUCT.unpack(handle.read(_FOOTER_STRUCT.size))
            handle.seek(-_FOOTER_STRUCT.size - footer_length, 2)
            try:
                table = json.loads(handle.read(footer_length).decode("utf-8"))
            except (json.JSONDecodeError, UnicodeDecodeError) as exc:
                raise FindingsStoreError(
                    f"{path.name}: block footer is corrupted: {exc}"
                ) from exc
        source._infos = [
            BlockInfo(offset=int(offset), length=int(length), count=int(count))
            for offset, length, count in table
        ]
        return source

    def block_infos(self) -> Sequence[BlockInfo]:
        return self._infos

    def read_block(self, index: int) -> bytes:
        info = self._infos[index]
        with self.path.open("rb") as handle:
            handle.seek(info.offset)
            blob = handle.read(info.length)
        if len(blob) != info.length:
            raise FindingsStoreError(f"{self.path.name}: truncated block {index}")
        return blob


class SpillWriter:
    """Buffered spill of ``(sort_key, payload)`` rows into sorted blocks."""

    def __init__(self, path: Path, *, block_rows: int = BLOCK_FINDINGS) -> None:
        self._file = BlockFile(path)
        self._buffer: list[tuple[list[object], object]] = []
        self._block_rows = block_rows
        self._total = 0

    def __enter__(self) -> SpillWriter:
        self._file.__enter__()
        return self

    def append(self, sort_key: Sequence[object], payload: object) -> None:
        self._buffer.append((list(sort_key), payload))
        self._total += 1
        if len(self._buffer) >= self._block_rows:
            self._flush()

    def _flush(self) -> None:
        if not self._buffer:
            return
        self._buffer.sort(key=lambda row: row[0])
        self._file.append_rows([[key, payload] for key, payload in self._buffer])
        self._buffer.clear()

    def __exit__(self, *exc_info: object) -> None:
        self._flush()
        self._file.__exit__(*exc_info)

    @property
    def total_rows(self) -> int:
        return self._total

    @property
    def path(self) -> Path:
        return self._file.path


def _block_rows(source: BlockSource, index: int) -> Iterator[list[Any]]:
    """Stream one block's rows without materializing the decoded block.

    ``heapq.merge`` holds ``fan_in`` of these iterators open at once, so each
    must hold only a bounded decompression buffer and one row — a full
    ``decode_block`` per iterator multiplies to gigabytes on monster spills.
    Lines are split per decompressed chunk (carrying only the partial tail)
    because per-line deletion from one growing buffer is quadratic.
    """
    decompressor = zlib.decompressobj()
    tail = b""
    data = b""
    view = memoryview(source.read_block(index))

    def parse(line: bytes) -> list[Any]:
        try:
            row = _json_loads(line)
        except (json.JSONDecodeError, ValueError) as exc:
            raise FindingsStoreError(f"block row is not valid JSON: {exc}") from exc
        if not isinstance(row, list) or len(row) != 2:
            raise FindingsStoreError("spill row must be a [key, payload] pair")
        return row

    try:
        for start in range(0, len(view), _DECOMPRESS_CHUNK):
            data = decompressor.decompress(view[start : start + _DECOMPRESS_CHUNK])
            if not data:
                continue
            lines = (tail + data).split(b"\n")
            tail = lines.pop()
            for line in lines:
                if line:
                    yield parse(line)
        data = decompressor.flush()
    except zlib.error as exc:
        raise FindingsStoreError(f"block payload is corrupted: {exc}") from exc
    for line in (tail + data).split(b"\n"):
        if line:
            yield parse(line)


_DECOMPRESS_CHUNK = 256 * 1024


def _run_rows(source: BlockSource, blocks: Sequence[int]) -> Iterator[list[Any]]:
    """One sorted run: its blocks streamed back to back."""
    for index in blocks:
        yield from _block_rows(source, index)


def merge_spill(
    path: Path,
    *,
    fan_in: int = _MERGE_FAN_IN,
) -> Iterator[object]:
    """Yield payloads in global sort-key order with bounded fan-in.

    Merging tracks sorted RUNS, not blocks: each merge pass fuses ``fan_in``
    runs into one longer run spanning many scratch blocks, so the run count
    divides by ``fan_in`` every pass and always converges. (Counting blocks
    instead loops forever once total rows exceed ``fan_in`` blocks' worth —
    a repacked scratch file holds just as many blocks as its input.)
    """
    source = BlockFile.open(path)
    runs: list[list[int]] = [[index] for index in range(len(source.block_infos()))]
    scratches: list[Path] = []
    suffix = 0
    try:
        while len(runs) > fan_in:
            scratch = path.with_suffix(f".merge{suffix}")
            suffix += 1
            new_runs: list[list[int]] = []
            with BlockFile(scratch) as staged:
                block_index = 0
                for start in range(0, len(runs), fan_in):
                    group = runs[start : start + fan_in]
                    run_blocks: list[int] = []
                    batch: list[list[Any]] = []
                    for row in heapq.merge(
                        *(_run_rows(source, run) for run in group),
                        key=lambda row: row[0],
                    ):
                        batch.append(row)
                        if len(batch) >= BLOCK_FINDINGS:
                            staged.append_rows(batch)
                            run_blocks.append(block_index)
                            block_index += 1
                            batch = []
                    if batch:
                        staged.append_rows(batch)
                        run_blocks.append(block_index)
                        block_index += 1
                    new_runs.append(run_blocks)
            source = BlockFile.open(scratch)
            runs = new_runs
            scratches.append(scratch)
            if len(scratches) > 1:
                scratches[-2].unlink(missing_ok=True)
        for row in heapq.merge(
            *(_run_rows(source, run) for run in runs),
            key=lambda row: row[0],
        ):
            yield row[1]
    finally:
        for scratch in scratches:
            scratch.unlink(missing_ok=True)


def write_finding_blocks(
    path: Path,
    payloads: Iterable[object],
    *,
    block_rows: int = BLOCK_FINDINGS,
) -> BlockFile:
    """Write finding payloads in their final order and return the container."""
    with BlockFile(path) as container:
        batch: list[object] = []
        for payload in payloads:
            batch.append(payload)
            if len(batch) >= block_rows:
                container.append_rows(batch)
                batch = []
        if batch:
            container.append_rows(batch)
    return BlockFile.open(path)


class FindingSequence(Sequence[Finding]):
    """Lazy ``Sequence[Finding]`` over a block source with an LRU page cache."""

    def __init__(
        self,
        source: BlockSource,
        *,
        cache_blocks: int = 4,
        decorate: Callable[[Finding], None] | None = None,
    ) -> None:
        self._source = source
        self._cache_blocks = cache_blocks
        self._decorate = decorate
        self._cache: OrderedDict[int, list[Finding]] = OrderedDict()
        self._starts: list[int] = []
        total = 0
        for info in source.block_infos():
            self._starts.append(total)
            total += info.count
        self._total = total

    def __len__(self) -> int:
        return self._total

    def __eq__(self, other: object) -> bool:
        """Element-wise equality with any findings sequence (list semantics)."""
        if isinstance(other, FindingSequence | list | tuple):
            if self._total != len(other):
                return False
            return all(mine == theirs for mine, theirs in zip(self, other, strict=True))
        return NotImplemented

    def _load_block(self, block_index: int) -> list[Finding]:
        cached = self._cache.get(block_index)
        if cached is not None:
            self._cache.move_to_end(block_index)
            return cached
        rows = decode_block(self._source.read_block(block_index))
        try:
            findings = [Finding.model_validate(row) for row in rows]
        except Exception as exc:  # pydantic ValidationError and shape errors
            raise FindingsStoreError(
                f"stored finding failed validation in block {block_index}: {exc}"
            ) from exc
        if self._decorate is not None:
            for finding in findings:
                self._decorate(finding)
        self._cache[block_index] = findings
        if len(self._cache) > self._cache_blocks:
            self._cache.popitem(last=False)
        return findings

    def _locate(self, index: int) -> tuple[int, int]:
        if index < 0:
            index += self._total
        if not 0 <= index < self._total:
            raise IndexError(index)
        low, high = 0, len(self._starts) - 1
        while low < high:
            mid = (low + high + 1) // 2
            if self._starts[mid] <= index:
                low = mid
            else:
                high = mid - 1
        return low, index - self._starts[low]

    def __getitem__(self, index):  # type: ignore[override]
        if isinstance(index, slice):
            return [self[i] for i in range(*index.indices(self._total))]
        block_index, row_index = self._locate(int(index))
        return self._load_block(block_index)[row_index]

    def __iter__(self) -> Iterator[Finding]:
        for block_index in range(len(self._starts)):
            yield from self._load_block(block_index)

    def iter_trusted(self) -> Iterator[Finding]:
        """Sequential pass using the trusted constructor, bypassing the cache.

        Only for passes over data this process wrote itself (record time):
        validation was already paid when the findings were produced, and
        skipping it plus the LRU roughly halves a full pass. Every shared
        read surface keeps the validating ``__iter__``.
        """
        for block_index in range(len(self._starts)):
            cached = self._cache.get(block_index)
            if cached is not None:
                yield from cached
                continue
            for row in decode_block(self._source.read_block(block_index)):
                finding = Finding.from_trusted_payload(row)  # type: ignore[arg-type]
                if self._decorate is not None:
                    self._decorate(finding)
                yield finding

    def iter_payloads(self) -> Iterator[object]:
        """Raw JSON payloads in order, skipping model validation."""
        for block_index in range(len(self._starts)):
            yield from decode_block(self._source.read_block(block_index))
