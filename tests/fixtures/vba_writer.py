"""Minimal MS-CFB and MS-OVBA writer used to build synthetic VBA fixtures.

The production reader in ``qc_tool/io/vba.py`` was validated byte-for-byte
against an independent OLE implementation on a genuine Microsoft-produced
``vbaProject.bin`` before this writer existed, so a round trip through the
reader is meaningful evidence rather than a self-agreeing pair.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from itertools import pairwise

_SECTOR = 512
_MINI_SECTOR = 64
_MINI_CUTOFF = 4096
_ENDOFCHAIN = 0xFFFFFFFE
_FREESECT = 0xFFFFFFFF
_FATSECT = 0xFFFFFFFD
_NOSTREAM = 0xFFFFFFFF
_MAX_CHUNK = 4096


def compress_ovba(payload: bytes) -> bytes:
    """Produce an MS-OVBA container, using copy tokens wherever they apply."""
    out = bytearray(b"\x01")
    for start in range(0, max(len(payload), 1), _MAX_CHUNK):
        block = payload[start : start + _MAX_CHUNK]
        out += _compressed_chunk(block)
    return bytes(out)


def _compressed_chunk(block: bytes) -> bytes:
    body = bytearray()
    tokens: list[bytes] = []
    flags = 0
    position = 0
    while position < len(block):
        offset, length = _longest_match(block, position)
        bit_count = max((position - 1).bit_length() if position else 0, 4)
        # A match longer than the token can express is truncated, never dropped.
        length = min(length, (0xFFFF >> bit_count) + 3)
        if length >= 3 and offset <= 1 << bit_count:
            token = ((offset - 1) << (16 - bit_count)) | (length - 3)
            tokens.append(token.to_bytes(2, "little"))
            flags |= 1 << (len(tokens) - 1)
            position += length
        else:
            tokens.append(block[position : position + 1])
            position += 1
        if len(tokens) == 8:
            body.append(flags)
            body += b"".join(tokens)
            tokens, flags = [], 0
    if tokens:
        body.append(flags)
        body += b"".join(tokens)
    if len(body) > _MAX_CHUNK:
        if len(block) != _MAX_CHUNK:
            raise ValueError("fixture block is incompressible and not a full chunk")
        return (0x3000 | (_MAX_CHUNK - 1)).to_bytes(2, "little") + block
    header = 0x8000 | 0x3000 | (len(body) - 1)
    return header.to_bytes(2, "little") + bytes(body)


def _longest_match(block: bytes, position: int) -> tuple[int, int]:
    best_offset = best_length = 0
    window_start = max(0, position - 4096)
    for candidate in range(window_start, position):
        length = 0
        while (
            position + length < len(block)
            and block[candidate + length] == block[position + length]
            and length < 4096
        ):
            length += 1
        if length > best_length:
            best_offset, best_length = position - candidate, length
    return best_offset, best_length


@dataclass
class _Entry:
    name: str
    kind: int
    data: bytes = b""
    children: list[int] = field(default_factory=list)
    start: int = _ENDOFCHAIN
    size: int = 0
    child: int = _NOSTREAM
    right: int = _NOSTREAM


def _entry_record(entry: _Entry) -> bytes:
    name = entry.name.encode("utf-16-le") + b"\x00\x00"
    record = bytearray(128)
    record[: len(name)] = name
    record[64:66] = len(name).to_bytes(2, "little")
    record[66] = entry.kind
    record[67] = 1  # black
    record[68:72] = _NOSTREAM.to_bytes(4, "little")
    record[72:76] = entry.right.to_bytes(4, "little")
    record[76:80] = entry.child.to_bytes(4, "little")
    record[116:120] = entry.start.to_bytes(4, "little")
    record[120:128] = entry.size.to_bytes(8, "little")
    return bytes(record)


def build_vba_project(modules: dict[str, str], *, protected: bool = False) -> bytes:
    """Assemble a compound file holding one VBA module stream per entry."""
    project_text = "ID=\"{00000000-0000-0000-0000-000000000000}\"\r\n"
    if protected:
        project_text += "CMG=\"0000\"\r\nDPB=\"0000\"\r\nGC=\"0000\"\r\n"

    entries = [_Entry("Root Entry", 5)]
    vba_children: list[int] = []
    entries.append(_Entry("PROJECT", 2, project_text.encode("cp1252")))
    project_index = len(entries) - 1
    entries.append(_Entry("VBA", 1))
    vba_index = len(entries) - 1
    for name in ("dir", "_VBA_PROJECT"):
        entries.append(_Entry(name, 2, b"\x01\x03\x00\x00"))
        vba_children.append(len(entries) - 1)
    for name, source in modules.items():
        cache = b"\x00" * 64
        entries.append(_Entry(name, 2, cache + compress_ovba(source.encode("cp1252"))))
        vba_children.append(len(entries) - 1)

    entries[0].child = project_index
    entries[project_index].right = vba_index
    entries[vba_index].child = vba_children[0]
    for previous, following in pairwise(vba_children):
        entries[previous].right = following

    mini_stream = bytearray()
    for entry in entries:
        if entry.kind != 2 or not entry.data:
            continue
        entry.start = len(mini_stream) // _MINI_SECTOR
        entry.size = len(entry.data)
        mini_stream += entry.data
        while len(mini_stream) % _MINI_SECTOR:
            mini_stream.append(0)

    mini_sector_count = len(mini_stream) // _MINI_SECTOR
    mini_fat = bytearray()
    for index in range(mini_sector_count):
        following = _ENDOFCHAIN
        for entry in entries:
            if entry.kind == 2 and entry.data:
                last = entry.start + -(-entry.size // _MINI_SECTOR) - 1
                if entry.start <= index < last:
                    following = index + 1
        mini_fat += following.to_bytes(4, "little")

    entries[0].start = 0  # placeholder, fixed below
    entries[0].size = len(mini_stream)

    directory = bytearray(b"".join(_entry_record(entry) for entry in entries))
    while len(directory) % _SECTOR:
        directory += bytes(128)

    def sector_count(payload: bytes | bytearray) -> int:
        return -(-len(payload) // _SECTOR) if payload else 0

    dir_sectors = sector_count(directory)
    minifat_sectors = sector_count(mini_fat)
    ministream_sectors = sector_count(mini_stream)

    first_dir = 1
    first_minifat = first_dir + dir_sectors
    first_ministream = first_minifat + minifat_sectors
    total = 1 + dir_sectors + minifat_sectors + ministream_sectors

    entries[0].start = first_ministream if ministream_sectors else _ENDOFCHAIN
    directory = bytearray(b"".join(_entry_record(entry) for entry in entries))
    while len(directory) % _SECTOR:
        directory += bytes(128)

    fat = [_FREESECT] * (_SECTOR // 4)
    fat[0] = _FATSECT
    for base, count in (
        (first_dir, dir_sectors),
        (first_minifat, minifat_sectors),
        (first_ministream, ministream_sectors),
    ):
        for offset in range(count):
            fat[base + offset] = (
                base + offset + 1 if offset + 1 < count else _ENDOFCHAIN
            )

    header = bytearray(_SECTOR)
    header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    header[24:26] = (0x003E).to_bytes(2, "little")
    header[26:28] = (0x0003).to_bytes(2, "little")
    header[28:30] = (0xFFFE).to_bytes(2, "little")
    header[30:32] = (9).to_bytes(2, "little")
    header[32:34] = (6).to_bytes(2, "little")
    header[44:48] = (1).to_bytes(4, "little")
    header[48:52] = first_dir.to_bytes(4, "little")
    header[56:60] = _MINI_CUTOFF.to_bytes(4, "little")
    header[60:64] = (
        first_minifat if minifat_sectors else _ENDOFCHAIN
    ).to_bytes(4, "little")
    header[64:68] = minifat_sectors.to_bytes(4, "little")
    header[68:72] = _ENDOFCHAIN.to_bytes(4, "little")
    header[72:76] = (0).to_bytes(4, "little")
    header[76:80] = (0).to_bytes(4, "little")
    for index in range(1, 109):
        header[76 + index * 4 : 80 + index * 4] = _FREESECT.to_bytes(4, "little")

    body = bytearray()
    body += b"".join(value.to_bytes(4, "little") for value in fat)
    body += directory
    body += mini_fat + bytes(-len(mini_fat) % _SECTOR)
    body += mini_stream + bytes(-len(mini_stream) % _SECTOR)
    assert len(body) == total * _SECTOR
    return bytes(header) + bytes(body)
