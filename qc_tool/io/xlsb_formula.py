"""Bounds-checked structural scanning for XLSB formula enrichment.

This module deliberately does not decode BIFF12 formula tokens. It identifies
formula-cell records and package features that make opening a workbook in an
external Office process unsafe. Formula text from Excel or LibreOffice is
accepted only when its coordinates exactly match this independent scan.
"""

from __future__ import annotations

import io
import posixpath
import zipfile
from collections.abc import Iterator
from dataclasses import dataclass, field
from xml.etree import ElementTree

_ROW_RECORD = 0x0000
_RETAINED_CELL_RECORDS = frozenset(
    {
        0x0002,  # BrtCellRk
        0x0003,  # BrtCellError
        0x0004,  # BrtCellBool
        0x0005,  # BrtCellReal
        0x0007,  # BrtCellIsst
        0x0008,  # BrtFmlaString
        0x0009,  # BrtFmlaNum
        0x000A,  # BrtFmlaBool
        0x000B,  # BrtFmlaError
    }
)
_FORMULA_RECORDS = frozenset({0x0008, 0x0009, 0x000A, 0x000B})
_DIMENSION_RECORD = 0x0194
_SHEET_RECORD = 0x019C
_MAX_RECORD_BYTES = 64 * 1024 * 1024
_MAX_STRING_CHARS = 1_000_000
_MAX_ROW = 1_048_575
_MAX_COLUMN = 16_383

_RISKY_PARTS = {
    "xl/vbaproject.bin": "VBA project",
    "xl/connections.bin": "external data connections",
}
_RISKY_PATH_SEGMENTS = {
    "/activex/": "ActiveX controls",
    "/ctrlprops/": "control properties",
    "/dialogsheets/": "dialog sheets",
    "/embeddings/": "embedded OLE content",
    "/externalconnections/": "external data connections",
    "/externallinks/": "external workbook links",
    "/macrosheets/": "Excel 4.0 macro sheets",
    "/querytables/": "external query tables",
    "/customui/": "custom Office UI",
}
_RISKY_RELATIONSHIP_KINDS = {
    "activexcontrol": "ActiveX controls",
    "attachedtoolbars": "custom Office toolbars",
    "connections": "external data connections",
    "ctrlprop": "control properties",
    "customui": "custom Office UI",
    "dialogsheet": "dialog sheets",
    "externallink": "external workbook links",
    "externallinkpath": "external workbook links",
    "macrosheet": "Excel 4.0 macro sheets",
    "oleobject": "embedded OLE content",
    "querytable": "external query tables",
    "vbaproject": "VBA project",
}
#: Relationship kinds allowed to target a package part outside ``xl/``.
#: These are metadata parts opened for risk inspection only; they are never
#: treated as worksheets. Every other non-``xl/`` target fails closed.
_NON_XL_RELATIONSHIP_KINDS = {"customxml": "customXml/"}


class XlsbFormulaScanError(ValueError):
    """The XLSB package cannot be scanned without guessing."""


@dataclass(frozen=True, slots=True)
class XlsbWorksheetMetrics:
    """Bounded structural workload facts for one binary worksheet."""

    cell_count: int = 0
    max_row: int = 0
    max_column: int = 0
    binary_bytes: int = 0


@dataclass(frozen=True, slots=True)
class XlsbFormulaScan:
    """Formula coordinates and risky package features found in an XLSB."""

    formula_cells: dict[str, frozenset[tuple[int, int]]]
    risky_features: tuple[str, ...] = ()
    worksheet_metrics: dict[str, XlsbWorksheetMetrics] = field(default_factory=dict)
    package_worksheet_binary_bytes: int = 0
    shared_string_bytes: int = 0
    styles_bytes: int = 0
    sheet_count: int = 0

    @property
    def formula_count(self) -> int:
        return sum(len(cells) for cells in self.formula_cells.values())

    @property
    def cell_count(self) -> int:
        return sum(metrics.cell_count for metrics in self.worksheet_metrics.values())

    @property
    def worksheet_binary_bytes(self) -> int:
        return self.package_worksheet_binary_bytes or sum(
            metrics.binary_bytes for metrics in self.worksheet_metrics.values()
        )

    @property
    def largest_sheet_area(self) -> int:
        return max(
            (
                metrics.max_row * metrics.max_column
                for metrics in self.worksheet_metrics.values()
            ),
            default=0,
        )

    @property
    def largest_sheet_dimensions(self) -> tuple[int, int]:
        if not self.worksheet_metrics:
            return 0, 0
        metrics = max(
            self.worksheet_metrics.values(),
            key=lambda item: item.max_row * item.max_column,
        )
        return metrics.max_row, metrics.max_column

    @property
    def workload_metrics_available(self) -> bool:
        return self.sheet_count > 0 and len(self.worksheet_metrics) == self.sheet_count

    @property
    def safe_for_external_engine(self) -> bool:
        return not self.risky_features


@dataclass(frozen=True, slots=True)
class _Relationship:
    target: str
    kind: str


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _read_record_id(data: bytes, position: int) -> tuple[int, int]:
    value = 0
    for index in range(4):
        if position >= len(data):
            raise XlsbFormulaScanError("truncated BIFF12 record id")
        byte = data[position]
        position += 1
        value += byte << (8 * index)
        if byte & 0x80 == 0:
            return value, position
    raise XlsbFormulaScanError("BIFF12 record id exceeds four bytes")


def _read_record_length(data: bytes, position: int) -> tuple[int, int]:
    value = 0
    for index in range(4):
        if position >= len(data):
            raise XlsbFormulaScanError("truncated BIFF12 record length")
        byte = data[position]
        position += 1
        value += (byte & 0x7F) << (7 * index)
        if byte & 0x80 == 0:
            if value > _MAX_RECORD_BYTES:
                raise XlsbFormulaScanError(
                    f"BIFF12 record length {value} exceeds the safety limit"
                )
            return value, position
    raise XlsbFormulaScanError("BIFF12 record length exceeds four bytes")


def _records(data: bytes) -> Iterator[tuple[int, bytes]]:
    position = 0
    while position < len(data):
        record_id, position = _read_record_id(data, position)
        length, position = _read_record_length(data, position)
        end = position + length
        if end > len(data):
            raise XlsbFormulaScanError(
                f"BIFF12 record {record_id:#x} is truncated: needs {length} bytes"
            )
        yield record_id, data[position:end]
        position = end


def _read_wide_string(payload: bytes, position: int) -> tuple[str, int]:
    if position + 4 > len(payload):
        raise XlsbFormulaScanError("truncated BIFF12 string length")
    length = int.from_bytes(payload[position : position + 4], "little")
    if length > _MAX_STRING_CHARS:
        raise XlsbFormulaScanError(f"BIFF12 string has unsafe length {length}")
    position += 4
    end = position + length * 2
    if end > len(payload):
        raise XlsbFormulaScanError("truncated BIFF12 UTF-16 string")
    try:
        value = payload[position:end].decode("utf-16-le")
    except UnicodeDecodeError as exc:
        raise XlsbFormulaScanError("invalid BIFF12 UTF-16 string") from exc
    return value, end


def _resolve_target(source_part: str, target: str, *, kind: str = "") -> str:
    normalized_target = target.replace("\\", "/")
    if normalized_target.startswith("/"):
        resolved = posixpath.normpath(normalized_target.lstrip("/"))
    else:
        resolved = posixpath.normpath(
            posixpath.join(posixpath.dirname(source_part), normalized_target)
        )
    if resolved == ".." or resolved.startswith("../"):
        raise XlsbFormulaScanError(f"unsafe workbook relationship target {target!r}")
    if resolved.startswith("xl/"):
        return resolved
    prefix = _NON_XL_RELATIONSHIP_KINDS.get(kind)
    if prefix is not None and resolved.startswith(prefix):
        return resolved
    raise XlsbFormulaScanError(f"unsafe workbook relationship target {target!r}")


def _workbook_relationships(archive: zipfile.ZipFile) -> dict[str, _Relationship]:
    part = "xl/_rels/workbook.bin.rels"
    try:
        root = ElementTree.fromstring(archive.read(part))
    except KeyError as exc:
        raise XlsbFormulaScanError(f"missing required XLSB part {part}") from exc
    except ElementTree.ParseError as exc:
        raise XlsbFormulaScanError(f"invalid relationship XML in {part}") from exc

    relationships: dict[str, _Relationship] = {}
    for element in root:
        if _local_name(element.tag) != "Relationship":
            continue
        relationship_id = element.get("Id")
        target = element.get("Target")
        relationship_type = element.get("Type", "")
        if not relationship_id or not target:
            raise XlsbFormulaScanError("workbook relationship is missing Id or Target")
        if element.get("TargetMode") == "External":
            continue
        if relationship_id in relationships:
            raise XlsbFormulaScanError(
                f"duplicate workbook relationship id {relationship_id!r}"
            )
        kind = relationship_type.rsplit("/", 1)[-1].lower()
        relationships[relationship_id] = _Relationship(
            target=_resolve_target("xl/workbook.bin", target, kind=kind),
            kind=kind,
        )
    return relationships


def _workbook_sheets(data: bytes) -> list[tuple[str, str]]:
    sheets: list[tuple[str, str]] = []
    names: set[str] = set()
    for record_id, payload in _records(data):
        if record_id != _SHEET_RECORD:
            continue
        if len(payload) < 12:
            raise XlsbFormulaScanError("truncated BrtBundleSh record")
        relationship_id, position = _read_wide_string(payload, 8)
        name, _ = _read_wide_string(payload, position)
        if not name or not relationship_id:
            raise XlsbFormulaScanError("BrtBundleSh has an empty name or relationship id")
        if name in names:
            raise XlsbFormulaScanError(f"duplicate workbook sheet name {name!r}")
        names.add(name)
        sheets.append((name, relationship_id))
    if not sheets:
        raise XlsbFormulaScanError("workbook contains no BrtBundleSh records")
    return sheets


def _worksheet_scan(
    data: bytes, sheet_name: str
) -> tuple[frozenset[tuple[int, int]], XlsbWorksheetMetrics]:
    row: int | None = None
    cells: set[tuple[int, int]] = set()
    cell_count = 0
    max_row = 0
    max_column = 0
    for record_id, payload in _records(data):
        if record_id == _ROW_RECORD:
            if len(payload) < 4:
                raise XlsbFormulaScanError(f"{sheet_name}: truncated BrtRowHdr record")
            row = int.from_bytes(payload[:4], "little")
            if row > _MAX_ROW:
                raise XlsbFormulaScanError(f"{sheet_name}: row {row} is out of bounds")
        elif record_id == _DIMENSION_RECORD:
            if len(payload) < 16:
                raise XlsbFormulaScanError(
                    f"{sheet_name}: truncated BrtWsDim record"
                )
            first_row = int.from_bytes(payload[0:4], "little")
            last_row = int.from_bytes(payload[4:8], "little")
            first_column = int.from_bytes(payload[8:12], "little")
            last_column = int.from_bytes(payload[12:16], "little")
            if first_row > last_row or last_row > _MAX_ROW:
                raise XlsbFormulaScanError(
                    f"{sheet_name}: worksheet row dimensions are out of bounds"
                )
            if first_column > last_column or last_column > _MAX_COLUMN:
                raise XlsbFormulaScanError(
                    f"{sheet_name}: worksheet column dimensions are out of bounds"
                )
            max_row = max(max_row, last_row + 1)
            max_column = max(max_column, last_column + 1)
        elif record_id in _RETAINED_CELL_RECORDS:
            if row is None:
                raise XlsbFormulaScanError(
                    f"{sheet_name}: cell record appears before a row header"
                )
            if len(payload) < 4:
                raise XlsbFormulaScanError(f"{sheet_name}: truncated cell record")
            column = int.from_bytes(payload[:4], "little")
            if column > _MAX_COLUMN:
                raise XlsbFormulaScanError(
                    f"{sheet_name}: column {column} is out of bounds"
                )
            cell_count += 1
            max_row = max(max_row, row + 1)
            max_column = max(max_column, column + 1)
            if record_id not in _FORMULA_RECORDS:
                continue
            coordinate = (row + 1, column + 1)
            if coordinate in cells:
                raise XlsbFormulaScanError(
                    f"{sheet_name}: duplicate formula cell at {coordinate}"
                )
            cells.add(coordinate)
    return frozenset(cells), XlsbWorksheetMetrics(
        cell_count=cell_count,
        max_row=max_row,
        max_column=max_column,
        binary_bytes=len(data),
    )


def _package_part_bytes(archive: zipfile.ZipFile, part: str) -> int:
    normalized = part.casefold()
    return sum(
        info.file_size
        for info in archive.infolist()
        if info.filename.replace("\\", "/").casefold() == normalized
    )


def _package_worksheet_bytes(archive: zipfile.ZipFile) -> int:
    return sum(
        info.file_size
        for info in archive.infolist()
        if (name := info.filename.replace("\\", "/").casefold()).startswith(
            "xl/worksheets/"
        )
        and name.endswith(".bin")
    )


def _risky_features(
    archive: zipfile.ZipFile, relationships: dict[str, _Relationship]
) -> tuple[str, ...]:
    features: set[str] = set()
    for name in archive.namelist():
        lowered = "/" + name.replace("\\", "/").lower().lstrip("/")
        exact = lowered.lstrip("/")
        if exact in _RISKY_PARTS:
            features.add(_RISKY_PARTS[exact])
        for segment, label in _RISKY_PATH_SEGMENTS.items():
            if segment in lowered:
                features.add(label)
    for relationship in relationships.values():
        label = _RISKY_RELATIONSHIP_KINDS.get(relationship.kind)
        if label:
            features.add(label)
    for part in archive.namelist():
        if not part.lower().endswith(".rels"):
            continue
        try:
            root = ElementTree.fromstring(archive.read(part))
        except ElementTree.ParseError:
            features.add("unreadable relationship metadata")
            continue
        for element in root:
            if _local_name(element.tag) != "Relationship":
                continue
            kind = element.get("Type", "").rsplit("/", 1)[-1].lower()
            label = _RISKY_RELATIONSHIP_KINDS.get(kind)
            if label:
                features.add(label)
            if element.get("TargetMode") == "External" and kind != "hyperlink":
                features.add("external relationships")
    return tuple(sorted(features))


def scan_xlsb_formulas(data: bytes) -> XlsbFormulaScan:
    """Return formula coordinates and risky features without decoding formulas."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            relationships = _workbook_relationships(archive)
            try:
                workbook_data = archive.read("xl/workbook.bin")
            except KeyError as exc:
                raise XlsbFormulaScanError("missing required XLSB part xl/workbook.bin") from exc

            sheets = _workbook_sheets(workbook_data)
            formula_cells: dict[str, frozenset[tuple[int, int]]] = {}
            worksheet_metrics: dict[str, XlsbWorksheetMetrics] = {}
            for sheet_name, relationship_id in sheets:
                relationship = relationships.get(relationship_id)
                if relationship is None:
                    raise XlsbFormulaScanError(
                        f"sheet {sheet_name!r} has unresolved relationship {relationship_id!r}"
                    )
                if relationship.kind != "worksheet":
                    formula_cells[sheet_name] = frozenset()
                    worksheet_metrics[sheet_name] = XlsbWorksheetMetrics()
                    continue
                try:
                    sheet_data = archive.read(relationship.target)
                except KeyError as exc:
                    raise XlsbFormulaScanError(
                        f"sheet {sheet_name!r} is missing part {relationship.target!r}"
                    ) from exc
                formulas, metrics = _worksheet_scan(sheet_data, sheet_name)
                formula_cells[sheet_name] = formulas
                worksheet_metrics[sheet_name] = metrics

            return XlsbFormulaScan(
                formula_cells=formula_cells,
                risky_features=_risky_features(archive, relationships),
                worksheet_metrics=worksheet_metrics,
                package_worksheet_binary_bytes=_package_worksheet_bytes(archive),
                shared_string_bytes=_package_part_bytes(
                    archive, "xl/sharedStrings.bin"
                ),
                styles_bytes=_package_part_bytes(archive, "xl/styles.bin"),
                sheet_count=len(sheets),
            )
    except zipfile.BadZipFile as exc:
        raise XlsbFormulaScanError("invalid XLSB ZIP package") from exc
