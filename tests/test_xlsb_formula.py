"""Structural XLSB formula scanning and safety classification."""

import struct
import zipfile
from pathlib import Path

import pytest
from pyxlsb import biff12

from qc_tool.io.xlsb_formula import XlsbFormulaScanError, scan_xlsb_formulas
from tests.fixtures.xlsb_writer import _record, write_xlsb


def _replace_part(path: Path, part: str, content: bytes) -> None:
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts[part] = content
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)


def _formula_sheet(*, malformed: bool = False) -> bytes:
    records = bytearray()
    records += _record(biff12.WORKSHEET)
    records += _record(biff12.SHEETDATA)
    records += _record(biff12.ROW, struct.pack("<I", 2))
    records += _record(biff12.FORMULA_FLOAT, struct.pack("<I", 4))
    records += _record(biff12.SHEETDATA_END)
    records += _record(biff12.WORKSHEET_END)
    if malformed:
        return bytes(records[:-1])
    return bytes(records)


def test_values_only_workbook_has_formula_presence_capability(tmp_path: Path) -> None:
    path = tmp_path / "values.xlsb"
    write_xlsb(path, {"Data": [["Label", "Value"], ["A", 1.0]]})

    scan = scan_xlsb_formulas(path.read_bytes())

    assert scan.formula_cells == {"Data": frozenset()}
    assert scan.formula_count == 0
    assert scan.safe_for_external_engine


def test_formula_records_are_reported_at_one_based_coordinates(tmp_path: Path) -> None:
    path = tmp_path / "formula.xlsb"
    write_xlsb(path, {"Data": [["Label"]]})
    _replace_part(path, "xl/worksheets/sheet1.bin", _formula_sheet())

    scan = scan_xlsb_formulas(path.read_bytes())

    assert scan.formula_cells == {"Data": frozenset({(3, 5)})}
    assert scan.formula_count == 1


def test_active_content_and_connections_block_external_engine(tmp_path: Path) -> None:
    path = tmp_path / "active.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("xl/vbaProject.bin", b"not executable in this fixture")
        archive.writestr("xl/queryTables/queryTable1.bin", b"query")

    scan = scan_xlsb_formulas(path.read_bytes())

    assert not scan.safe_for_external_engine
    assert scan.risky_features == ("VBA project", "external query tables")


def test_truncated_biff_record_fails_closed(tmp_path: Path) -> None:
    path = tmp_path / "malformed.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    _replace_part(path, "xl/worksheets/sheet1.bin", _formula_sheet(malformed=True))

    with pytest.raises(XlsbFormulaScanError, match="truncated"):
        scan_xlsb_formulas(path.read_bytes())


def test_relationship_target_cannot_escape_package(tmp_path: Path) -> None:
    path = tmp_path / "escape.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    relationships = b"""<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Target="../../outside.bin"/>
    </Relationships>"""
    _replace_part(path, "xl/_rels/workbook.bin.rels", relationships)

    with pytest.raises(XlsbFormulaScanError, match="unsafe workbook relationship"):
        scan_xlsb_formulas(path.read_bytes())


def test_external_non_hyperlink_relationship_blocks_external_engine(tmp_path: Path) -> None:
    path = tmp_path / "external.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    relationships = b"""<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdExternal"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        Target="https://example.invalid/image.png" TargetMode="External"/>
    </Relationships>"""
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("xl/worksheets/_rels/sheet1.bin.rels", relationships)

    scan = scan_xlsb_formulas(path.read_bytes())

    assert "external relationships" in scan.risky_features
    assert not scan.safe_for_external_engine
