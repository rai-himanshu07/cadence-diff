"""VBA project reading for module inventory and module-text comparison.

``xl/vbaProject.bin`` is an MS-CFB compound file. Its ``VBA`` storage holds one
stream per module: a performance cache followed by the module source compressed
with the MS-OVBA run-length format. Everything here decodes bytes only. No macro
is executed, evaluated, interpreted, or resolved, and nothing outside the
supplied package bytes is read.
"""

from __future__ import annotations

import hashlib
import io
import zipfile
from dataclasses import dataclass, field

_VBA_PART = "xl/vbaproject.bin"
_CFB_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_ENDOFCHAIN = 0xFFFFFFFE
_FREESECT = 0xFFFFFFFF
_MAX_SECTOR_COUNT = 1 << 22
_MAX_DIRECTORY_ENTRIES = 1 << 16
_DIRECTORY_ENTRY_SIZE = 128
_NOSTREAM = 0xFFFFFFFF
_STORAGE = 1
_STREAM = 2
_ROOT = 5
#: Streams that the format reserves; every other VBA child stream is a module.
_RESERVED_STREAMS = {"dir", "_vba_project"}
_RESERVED_PREFIX = "__srp_"
#: A false compressed-container start is cheap to reject but expensive to expand,
#: so only this many candidate offsets are ever fully decompressed per stream.
_MAX_CONTAINER_ATTEMPTS = 64


class VbaReadError(ValueError):
    """The VBA project cannot be decoded without guessing."""


@dataclass(frozen=True, slots=True)
class VbaModule:
    """One VBA module, identified by name and compared by source text."""

    name: str
    line_count: int
    digest: str
    text: str = field(default="", compare=False, repr=False)


@dataclass(frozen=True, slots=True)
class VbaProjectScan:
    """What a package declares about its VBA project."""

    modules: tuple[VbaModule, ...] = ()
    present: bool = False
    available: bool = False
    protected: bool = False
    unreadable_modules: int = 0
    detail: str = ""


# --- MS-OVBA compressed container ---------------------------------------


def decompress_ovba(data: bytes, start: int = 0) -> bytes:
    """Expand one MS-OVBA compressed container, refusing anything malformed."""
    if start >= len(data) or data[start] != 0x01:
        raise VbaReadError("compressed container does not start with the 0x01 signature")
    out = bytearray()
    position = start + 1
    while position < len(data):
        if position + 2 > len(data):
            raise VbaReadError("truncated chunk header")
        header = int.from_bytes(data[position : position + 2], "little")
        position += 2
        if (header >> 12) & 0x07 != 0b011:
            raise VbaReadError("chunk header signature is not 0b011")
        end = position + (header & 0x0FFF) + 1
        if end > len(data):
            raise VbaReadError("chunk extends past the end of the stream")
        if not header & 0x8000:
            out += data[position:end]
            position = end
            continue
        chunk_start = len(out)
        while position < end:
            flags = data[position]
            position += 1
            for bit in range(8):
                if position >= end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[position])
                    position += 1
                    continue
                if position + 2 > end:
                    raise VbaReadError("truncated copy token")
                token = int.from_bytes(data[position : position + 2], "little")
                position += 2
                difference = len(out) - chunk_start
                bit_count = max((difference - 1).bit_length() if difference else 0, 4)
                length = (token & (0xFFFF >> bit_count)) + 3
                source = len(out) - ((token >> (16 - bit_count)) + 1)
                if source < chunk_start:
                    raise VbaReadError("copy token points before its chunk")
                for _ in range(length):
                    out.append(out[source])
                    source += 1
        position = end
    return bytes(out)


# --- MS-CFB compound file ------------------------------------------------


@dataclass(frozen=True, slots=True)
class _DirectoryEntry:
    name: str
    kind: int
    left: int
    right: int
    child: int
    start: int
    size: int


class _CompoundFile:
    """Minimal read-only MS-CFB reader with bounded chain following."""

    def __init__(self, data: bytes) -> None:
        if len(data) < 512 or not data.startswith(_CFB_SIGNATURE):
            raise VbaReadError("not an MS-CFB compound file")
        self._data = data
        self._sector_size = 1 << int.from_bytes(data[30:32], "little")
        self._mini_sector_size = 1 << int.from_bytes(data[32:34], "little")
        if self._sector_size not in (512, 4096) or self._mini_sector_size != 64:
            raise VbaReadError("unsupported compound-file sector geometry")
        self._mini_cutoff = int.from_bytes(data[56:60], "little")
        self._fat = self._read_fat()
        self._entries = self._read_directory(int.from_bytes(data[48:52], "little"))
        root = self._entries[0]
        self._mini_stream = self._read_chain(root.start, root.size, mini=False)
        self._mini_fat = self._read_mini_fat(
            int.from_bytes(data[60:64], "little"),
            int.from_bytes(data[64:68], "little"),
        )

    def _sector(self, number: int) -> bytes:
        offset = (number + 1) * self._sector_size
        if offset < 0 or offset + self._sector_size > len(self._data):
            raise VbaReadError("sector points outside the compound file")
        return self._data[offset : offset + self._sector_size]

    def _read_fat(self) -> list[int]:
        difat = [
            int.from_bytes(self._data[76 + index * 4 : 80 + index * 4], "little")
            for index in range(109)
        ]
        next_difat = int.from_bytes(self._data[68:72], "little")
        per_sector = self._sector_size // 4
        guard = 0
        while next_difat not in (_ENDOFCHAIN, _FREESECT):
            guard += 1
            if guard > _MAX_SECTOR_COUNT:
                raise VbaReadError("DIFAT chain does not terminate")
            sector = self._sector(next_difat)
            difat.extend(
                int.from_bytes(sector[index * 4 : index * 4 + 4], "little")
                for index in range(per_sector - 1)
            )
            next_difat = int.from_bytes(sector[-4:], "little")
        fat: list[int] = []
        for number in difat:
            if number in (_ENDOFCHAIN, _FREESECT) or len(fat) > _MAX_SECTOR_COUNT:
                continue
            sector = self._sector(number)
            fat.extend(
                int.from_bytes(sector[index * 4 : index * 4 + 4], "little")
                for index in range(per_sector)
            )
        if not fat:
            raise VbaReadError("compound file declares no allocation table")
        return fat

    def _read_mini_fat(self, start: int, count: int) -> list[int]:
        if count <= 0:
            return []
        raw = self._read_chain(start, count * self._sector_size, mini=False)
        return [
            int.from_bytes(raw[index : index + 4], "little")
            for index in range(0, len(raw) - 3, 4)
        ]

    def _chain(self, start: int, table: list[int]) -> list[int]:
        chain: list[int] = []
        seen: set[int] = set()
        current = start
        while current not in (_ENDOFCHAIN, _FREESECT):
            if current < 0 or current >= len(table) or current in seen:
                raise VbaReadError("allocation chain is cyclic or out of range")
            seen.add(current)
            chain.append(current)
            if len(chain) > _MAX_SECTOR_COUNT:
                raise VbaReadError("allocation chain is unreasonably long")
            current = table[current]
        return chain

    def _read_chain(self, start: int, size: int, *, mini: bool) -> bytes:
        if start in (_ENDOFCHAIN, _FREESECT) or size <= 0:
            return b""
        if mini:
            unit = self._mini_sector_size
            chain = self._chain(start, self._mini_fat)
            source = self._mini_stream
            blocks = [source[number * unit : number * unit + unit] for number in chain]
        else:
            blocks = [self._sector(number) for number in self._chain(start, self._fat)]
        return b"".join(blocks)[:size]

    def _read_directory(self, start: int) -> list[_DirectoryEntry]:
        raw = b"".join(self._sector(number) for number in self._chain(start, self._fat))
        entries: list[_DirectoryEntry] = []
        for offset in range(0, len(raw) - _DIRECTORY_ENTRY_SIZE + 1, _DIRECTORY_ENTRY_SIZE):
            block = raw[offset : offset + _DIRECTORY_ENTRY_SIZE]
            name_length = int.from_bytes(block[64:66], "little")
            if not 0 < name_length <= 64:
                name = ""
            else:
                name = block[: name_length - 2].decode("utf-16-le", errors="replace")
            entries.append(
                _DirectoryEntry(
                    name=name,
                    kind=block[66],
                    left=int.from_bytes(block[68:72], "little"),
                    right=int.from_bytes(block[72:76], "little"),
                    child=int.from_bytes(block[76:80], "little"),
                    start=int.from_bytes(block[116:120], "little"),
                    size=int.from_bytes(block[120:128], "little"),
                )
            )
            if len(entries) > _MAX_DIRECTORY_ENTRIES:
                raise VbaReadError("compound file declares too many directory entries")
        if not entries or entries[0].kind != _ROOT:
            raise VbaReadError("compound file has no root directory entry")
        return entries

    def children(self, index: int) -> list[_DirectoryEntry]:
        """Directory children of one entry, walked without trusting the tree shape."""
        if not 0 <= index < len(self._entries):
            return []
        pending = [self._entries[index].child]
        seen: set[int] = set()
        found: list[_DirectoryEntry] = []
        while pending:
            current = pending.pop()
            if current == _NOSTREAM or not 0 <= current < len(self._entries):
                continue
            if current in seen:
                continue
            seen.add(current)
            entry = self._entries[current]
            found.append(entry)
            pending.extend((entry.left, entry.right, entry.child))
        return found

    def find(self, name: str, kind: int) -> int | None:
        target = name.casefold()
        for index, entry in enumerate(self._entries):
            if entry.kind == kind and entry.name.casefold() == target:
                return index
        return None

    def read(self, entry: _DirectoryEntry) -> bytes:
        mini = entry.size < self._mini_cutoff and entry.kind != _ROOT
        return self._read_chain(entry.start, entry.size, mini=mini)


# --- module extraction ---------------------------------------------------


def _module_source(stream: bytes) -> str | None:
    """Find and expand the compressed source that follows the performance cache."""
    fallback: str | None = None
    attempts = 0
    for offset in range(len(stream) - 2):
        if stream[offset] != 0x01:
            continue
        header = int.from_bytes(stream[offset + 1 : offset + 3], "little")
        if (header >> 12) & 0x07 != 0b011:
            continue
        attempts += 1
        if attempts > _MAX_CONTAINER_ATTEMPTS:
            break
        try:
            expanded = decompress_ovba(stream, offset)
        except VbaReadError:
            continue
        # VBA source is stored in the project code page; cp1252 is a superset of
        # ASCII, so pure-ASCII modules decode exactly.
        text = expanded.decode("cp1252", errors="replace")
        if text.startswith("Attribute "):
            return text
        if fallback is None:
            fallback = text
    return fallback


def _module(name: str, text: str) -> VbaModule:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    return VbaModule(
        name=name,
        line_count=len(normalized.splitlines()),
        digest=hashlib.sha256(normalized.encode("utf-8", errors="replace")).hexdigest(),
        text=normalized,
    )


def scan_vba_project(data: bytes) -> VbaProjectScan:
    """Read every VBA module in an OOXML package without executing anything."""
    try:
        archive = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        return VbaProjectScan(detail="workbook is not an OOXML package")
    with archive:
        part = next(
            (name for name in archive.namelist() if name.casefold() == _VBA_PART),
            None,
        )
        if part is None:
            return VbaProjectScan(present=False, available=True)
        project = archive.read(part)

    try:
        compound = _CompoundFile(project)
        storage = compound.find("VBA", _STORAGE)
        if storage is None:
            raise VbaReadError("vbaProject.bin has no VBA storage")
        children = compound.children(storage)
    except VbaReadError as exc:
        return VbaProjectScan(present=True, available=False, detail=str(exc))

    protected = False
    for entry in compound.children(0):
        if entry.kind == _STREAM and entry.name.casefold() == "project":
            text = compound.read(entry).decode("cp1252", errors="replace")
            protected = any(marker in text for marker in ("CMG=", "DPB=", "GC="))
            break

    modules: list[VbaModule] = []
    unreadable = 0
    for entry in children:
        if entry.kind != _STREAM:
            continue
        lowered = entry.name.casefold()
        if lowered in _RESERVED_STREAMS or lowered.startswith(_RESERVED_PREFIX):
            continue
        try:
            stream = compound.read(entry)
        except VbaReadError:
            unreadable += 1
            continue
        source = _module_source(stream)
        if source is None:
            unreadable += 1
            continue
        modules.append(_module(entry.name, source))

    details: list[str] = []
    if unreadable:
        details.append(f"{unreadable} module streams could not be decoded")
    if protected:
        details.append("the VBA project is locked for viewing in Excel")
    return VbaProjectScan(
        modules=tuple(sorted(modules, key=lambda item: item.name)),
        present=True,
        available=True,
        protected=protected,
        unreadable_modules=unreadable,
        detail="; ".join(details),
    )
