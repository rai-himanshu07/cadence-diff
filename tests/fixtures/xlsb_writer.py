"""Minimal BIFF12 (.xlsb) writer for test fixtures.

Emits values-only workbooks — exactly the scope of the project's degraded
xlsb QC path. Record layout is grounded on the pyxlsb reader implementation
(`pyxlsb.reader`, `pyxlsb.handlers`), which is the production read path for
xlsb in this project. Compatibility target is pyxlsb readback, not Excel:
real deliverables are read, never written, so fixtures only need to exercise
the reader.

Style/date-system records (`BrtWbProp`, `BrtBeginFmts`/`BrtFmt`,
`BrtBeginCellXfs`/`BrtXF`) are grounded the same way: cross-checked against
calamine's independently tested `xlsb` reader (MIT-licensed,
https://github.com/tafia/calamine), not pyxlsb, because pyxlsb never reads
`xl/styles.bin` or the workbook date-system flag at all.
"""

import struct
import zipfile
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path

from pyxlsb import biff12

CellValue = str | float | int | None


@dataclass(frozen=True, slots=True)
class StyledCell:
    """A cell needing an explicit cell-XF index and/or a formula record.

    ``xf_index`` indexes into the ``cell_xfs`` list passed to ``write_xlsb``.
    ``is_formula`` selects BrtFmlaNum instead of BrtCellReal; pyxlsb's own
    ``CellHandler`` reads an identical col+style+double layout for both and
    the BIFF12Reader main loop always seeks to the next record boundary by
    declared length, so a minimal formula record with no trailing token
    stream is valid input — exactly what the project's own risk/formula
    scanner already assumes.
    """

    value: float | int
    xf_index: int = 0
    is_formula: bool = False


#: One populated cell: a bare value (xf_index 0, not a formula) or an
#: explicit ``StyledCell``.
RowCell = CellValue | StyledCell

_ZIP_EPOCH = (1980, 1, 1, 0, 0, 0)
#: Naive-decoded (not spec-documented) record ids -- see the matching
#: constants and comment in `qc_tool/io/xlsb_formula.py` for why these differ
#: from the officially-documented BrtWbProp/BrtBeginFmts/BrtBeginCellXfs
#: numbers, and how they were verified against an independent reader.
_WBPROP = 0x0199
_FMTS_BEGIN = 0x04E7
_FMT = 0x002C
_CELLXFS_BEGIN = 0x04E9
_XF = 0x002F

_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" '
    'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="bin" '
    'ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>'
    "</Types>"
)

_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/officeDocument" Target="xl/workbook.bin"/>'
    "</Relationships>"
)


def _rec_id(recid: int) -> bytes:
    """Encode a record id the way ``BIFF12Reader.read_id`` decodes it."""
    if recid < 0x80:
        return bytes([recid])
    low, high = recid & 0xFF, (recid >> 8) & 0xFF
    if not low & 0x80 or high & 0x80:
        raise ValueError(f"record id {recid:#x} is not encodable in two bytes")
    return bytes([low, high])


def _rec_len(length: int) -> bytes:
    """Encode a record length as the 7-bit little-endian varint of read_len."""
    out = bytearray()
    while True:
        chunk = length & 0x7F
        length >>= 7
        if length:
            out.append(chunk | 0x80)
        else:
            out.append(chunk)
            return bytes(out)


def _record(recid: int, payload: bytes = b"") -> bytes:
    return _rec_id(recid) + _rec_len(len(payload)) + payload


def _xl_string(text: str) -> bytes:
    encoded = text.encode("utf-16-le")
    return struct.pack("<I", len(encoded) // 2) + encoded


def _workbook_part(sheet_names: list[str], *, date1904: bool) -> bytes:
    out = bytearray()
    out += _record(biff12.WORKBOOK)
    out += _record(_WBPROP, struct.pack("<I", 1 if date1904 else 0))
    out += _record(biff12.SHEETS)
    for index, name in enumerate(sheet_names, start=1):
        payload = struct.pack("<II", 0, index) + _xl_string(f"rId{index}") + _xl_string(name)
        out += _record(biff12.SHEET, payload)
    out += _record(biff12.SHEETS_END)
    out += _record(biff12.WORKBOOK_END)
    return bytes(out)


def _workbook_rels(sheet_count: int) -> bytes:
    rels = [
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/'
        f'2006/relationships/worksheet" Target="worksheets/sheet{i}.bin"/>'
        for i in range(1, sheet_count + 1)
    ]
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels)
        + "</Relationships>"
    )
    return xml.encode("utf-8")


def _shared_strings_part(strings: list[str], total_count: int) -> bytes:
    out = bytearray()
    out += _record(biff12.SST, struct.pack("<II", total_count, len(strings)))
    for text in strings:
        out += _record(biff12.SI, b"\x00" + _xl_string(text))
    out += _record(biff12.SST_END)
    return bytes(out)


def _styles_part(custom_formats: dict[int, str], cell_xfs: list[int]) -> bytes:
    """Just enough of ``xl/styles.bin`` for `parse_xlsb_styles` to round-trip.

    No `BrtBeginStyleSheet`/end-marker wrapper or font/fill/border/cell-style
    groups: the project's reader is a flat count-driven scan that never looks
    for them, so omitting them keeps this writer using only verified IDs.
    """
    out = bytearray()
    out += _record(_FMTS_BEGIN, struct.pack("<I", len(custom_formats)))
    for ifmt, code in custom_formats.items():
        out += _record(_FMT, struct.pack("<H", ifmt) + _xl_string(code))
    out += _record(_CELLXFS_BEGIN, struct.pack("<I", len(cell_xfs)))
    for ifmt in cell_xfs:
        out += _record(_XF, struct.pack("<HH", 0, ifmt))  # ixfParent=0, iFmt
    return bytes(out)


def _sheet_part(rows: Sequence[Sequence[RowCell]], string_index: dict[str, int]) -> bytes:
    n_rows = len(rows)
    n_cols = max((len(r) for r in rows), default=1)
    out = bytearray()
    out += _record(biff12.WORKSHEET)
    out += _record(
        biff12.DIMENSION, struct.pack("<IIII", 0, max(n_rows - 1, 0), 0, max(n_cols - 1, 0))
    )
    out += _record(biff12.SHEETDATA)
    for row_idx, row in enumerate(rows):
        out += _record(biff12.ROW, struct.pack("<I", row_idx))
        for col_idx, cell in enumerate(row):
            if cell is None:
                continue
            if isinstance(cell, StyledCell):
                value, xf_index, is_formula = cell.value, cell.xf_index, cell.is_formula
            else:
                value, xf_index, is_formula = cell, 0, False
            if isinstance(value, str):
                payload = struct.pack("<III", col_idx, xf_index, string_index[value])
                out += _record(biff12.STRING, payload)
            else:
                payload = struct.pack("<IId", col_idx, xf_index, float(value))
                out += _record(biff12.FORMULA_FLOAT if is_formula else biff12.FLOAT, payload)
    out += _record(biff12.SHEETDATA_END)
    out += _record(biff12.WORKSHEET_END)
    return bytes(out)


def write_xlsb(
    path: Path,
    sheets: Mapping[str, Sequence[Sequence[RowCell]]],
    *,
    date1904: bool = False,
    custom_formats: dict[int, str] | None = None,
    cell_xfs: list[int] | None = None,
) -> None:
    """Write a deterministic, pyxlsb-readable xlsb workbook.

    ``custom_formats``/``cell_xfs`` are opt-in: omitting both keeps every
    existing caller's output byte-identical (no `xl/styles.bin` part at all).
    """
    strings: list[str] = []
    string_index: dict[str, int] = {}
    total_strings = 0
    for rows in sheets.values():
        for row in rows:
            for cell in row:
                value = cell.value if isinstance(cell, StyledCell) else cell
                if isinstance(value, str):
                    total_strings += 1
                    if value not in string_index:
                        string_index[value] = len(strings)
                        strings.append(value)

    sheet_names = list(sheets)
    parts: dict[str, bytes] = {
        "[Content_Types].xml": _CONTENT_TYPES.encode("utf-8"),
        "_rels/.rels": _ROOT_RELS.encode("utf-8"),
        "xl/workbook.bin": _workbook_part(sheet_names, date1904=date1904),
        "xl/_rels/workbook.bin.rels": _workbook_rels(len(sheet_names)),
        "xl/sharedStrings.bin": _shared_strings_part(strings, total_strings),
    }
    if custom_formats or cell_xfs:
        # Read directly by its fixed package path (like the existing workload
        # scan already does), so no workbook-relationship entry is needed.
        parts["xl/styles.bin"] = _styles_part(custom_formats or {}, cell_xfs or [0])
    for index, rows in enumerate(sheets.values(), start=1):
        parts[f"xl/worksheets/sheet{index}.bin"] = _sheet_part(rows, string_index)

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in sorted(parts):
            info = zipfile.ZipInfo(name, date_time=_ZIP_EPOCH)
            info.compress_type = zipfile.ZIP_DEFLATED
            zf.writestr(info, parts[name])
