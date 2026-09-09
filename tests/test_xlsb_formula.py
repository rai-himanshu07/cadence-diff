"""Structural XLSB formula scanning and safety classification."""

import hashlib
import struct
import zipfile
from pathlib import Path

import pytest
from pyxlsb import biff12

from qc_tool.io.xlsb_formula import (
    XlsbFormulaScanError,
    parse_worksheet_cell_styles,
    parse_xlsb_date_system,
    parse_xlsb_styles,
    resolve_worksheet_targets,
    scan_xlsb_formulas,
)
from tests.fixtures.xlsb_writer import StyledCell, _record, write_xlsb


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
    assert scan.cell_count == 4
    assert scan.sheet_count == 1
    assert scan.worksheet_binary_bytes > 0
    assert scan.shared_string_bytes > 0
    assert scan.styles_bytes == 0
    assert scan.largest_sheet_area == 4
    assert scan.largest_sheet_dimensions == (2, 2)
    assert scan.worksheet_metrics["Data"].max_row == 2
    assert scan.worksheet_metrics["Data"].max_column == 2


def test_formula_records_are_reported_at_one_based_coordinates(tmp_path: Path) -> None:
    path = tmp_path / "formula.xlsb"
    write_xlsb(path, {"Data": [["Label"]]})
    _replace_part(path, "xl/worksheets/sheet1.bin", _formula_sheet())

    scan = scan_xlsb_formulas(path.read_bytes())

    assert scan.formula_cells == {"Data": frozenset({(3, 5)})}
    assert scan.formula_count == 1
    assert scan.cell_count == 1
    assert scan.largest_sheet_area == 15
    assert scan.largest_sheet_dimensions == (3, 5)


def test_workload_bytes_include_unreferenced_worksheet_parts(tmp_path: Path) -> None:
    path = tmp_path / "orphan.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    scan_before = scan_xlsb_formulas(path.read_bytes())
    orphan = b"unreferenced worksheet bytes"
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("xl/worksheets/orphan.bin", orphan)

    scan = scan_xlsb_formulas(path.read_bytes())

    assert scan.cell_count == scan_before.cell_count
    assert scan.worksheet_binary_bytes == scan_before.worksheet_binary_bytes + len(orphan)


def test_active_content_and_connections_block_external_engine(tmp_path: Path) -> None:
    path = tmp_path / "active.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("xl/vbaProject.bin", b"not executable in this fixture")
        archive.writestr("xl/queryTables/queryTable1.bin", b"query")

    scan = scan_xlsb_formulas(path.read_bytes())

    assert not scan.safe_for_external_engine
    assert scan.risky_features == ("VBA project", "external query tables")


def test_legacy_risky_features_only_constructor_remains_fail_closed() -> None:
    from qc_tool.io.xlsb_formula import XlsbFormulaScan

    scan = XlsbFormulaScan(
        formula_cells={},
        risky_features=("VBA project",),
    )

    assert not scan.safe_for_external_engine


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


def test_customxml_relationship_is_scanned_without_mutating_the_source(
    tmp_path: Path,
) -> None:
    path = tmp_path / "customxml.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    relationships = b"""<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Target="worksheets/sheet1.bin"/>
      <Relationship Id="rIdCustom"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
        Target="../customXml/item1.xml"/>
    </Relationships>"""
    _replace_part(path, "xl/_rels/workbook.bin.rels", relationships)
    before = path.read_bytes()

    scan = scan_xlsb_formulas(before)

    assert "Data" in scan.formula_cells
    assert hashlib.sha256(path.read_bytes()).digest() == hashlib.sha256(before).digest()


def test_non_customxml_targets_outside_xl_still_fail_closed(tmp_path: Path) -> None:
    path = tmp_path / "outside.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    relationships = b"""<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Target="worksheets/sheet1.bin"/>
      <Relationship Id="rIdOther"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        Target="../media/image1.png"/>
    </Relationships>"""
    _replace_part(path, "xl/_rels/workbook.bin.rels", relationships)

    with pytest.raises(XlsbFormulaScanError, match="unsafe workbook relationship"):
        scan_xlsb_formulas(path.read_bytes())


def test_customxml_relationship_cannot_escape_its_allowed_prefix(tmp_path: Path) -> None:
    path = tmp_path / "traversal.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    relationships = b"""<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Target="worksheets/sheet1.bin"/>
      <Relationship Id="rIdCustom"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
        Target="../../customXml/item1.xml"/>
    </Relationships>"""
    _replace_part(path, "xl/_rels/workbook.bin.rels", relationships)

    with pytest.raises(XlsbFormulaScanError, match="unsafe workbook relationship"):
        scan_xlsb_formulas(path.read_bytes())


def test_a_customxml_relationship_is_never_treated_as_a_worksheet(
    tmp_path: Path,
) -> None:
    path = tmp_path / "sheet-customxml.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    relationships = b"""<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
        Target="../customXml/item1.xml"/>
    </Relationships>"""
    _replace_part(path, "xl/_rels/workbook.bin.rels", relationships)

    scan = scan_xlsb_formulas(path.read_bytes())

    assert scan.formula_cells == {"Data": frozenset()}


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


# --- Step 4a: passive/blocking/unknown-external classification -----------


def test_external_link_path_segment_is_classified_passive(tmp_path: Path) -> None:
    path = tmp_path / "passive-part.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("xl/externalLinks/externalLink1.xml", b"<externalLink/>")

    scan = scan_xlsb_formulas(path.read_bytes())

    assert scan.passive_features == ("external workbook links",)
    assert scan.blocking_features == ()
    assert scan.unknown_external_features == ()
    assert scan.risky_features == ("external workbook links",)
    assert scan.safe_for_external_engine


def test_externallinkpath_targetmode_external_is_classified_passive(
    tmp_path: Path,
) -> None:
    """The real production shape: an externalLinkPath relationship with
    TargetMode=External is passive, not unknown-external."""
    path = tmp_path / "passive-rel.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    relationships = b"""<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdExternal"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
        Target="file:///C:/other.xlsx" TargetMode="External"/>
    </Relationships>"""
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("xl/externalLinks/_rels/externalLink1.xml.rels", relationships)

    scan = scan_xlsb_formulas(path.read_bytes())

    assert scan.passive_features == ("external relationships", "external workbook links")
    assert scan.blocking_features == ()
    assert scan.unknown_external_features == ()
    assert scan.safe_for_external_engine


def test_externallinklongpath_targetmode_external_is_classified_passive(
    tmp_path: Path,
) -> None:
    path = tmp_path / "passive-long-rel.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    relationships = b"""<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdExternal"
        Type="http://schemas.microsoft.com/office/2006/relationships/externalLinkLongPath"
        Target="file:///C:/a/long/path/other.xlsx" TargetMode="External"/>
    </Relationships>"""
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("xl/externalLinks/_rels/externalLink1.xml.rels", relationships)

    scan = scan_xlsb_formulas(path.read_bytes())

    assert scan.passive_features == ("external relationships", "external workbook links")
    assert scan.blocking_features == ()
    assert scan.unknown_external_features == ()
    assert scan.safe_for_external_engine


def test_unknown_external_targetmode_kind_refuses_but_is_not_blocking(
    tmp_path: Path,
) -> None:
    """A non-hyperlink external relationship of an unrecognized kind must
    refuse the adapter, but its own classification is unknown_external, not
    blocking -- Step 4b's reachability evidence is what can later resolve it,
    not a permanent "active content" label."""
    path = tmp_path / "unknown-external.xlsb"
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

    assert scan.unknown_external_features == ("external relationships",)
    assert scan.blocking_features == ()
    assert scan.passive_features == ()
    assert not scan.safe_for_external_engine


def test_active_content_blocks_even_alongside_passive_metadata(tmp_path: Path) -> None:
    """Blocking always wins: a passive external-link part next to a VBA
    project must still refuse the adapter."""
    path = tmp_path / "mixed.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    with zipfile.ZipFile(path, "a") as archive:
        archive.writestr("xl/externalLinks/externalLink1.xml", b"<externalLink/>")
        archive.writestr("xl/vbaProject.bin", b"not executable in this fixture")

    scan = scan_xlsb_formulas(path.read_bytes())

    assert "external workbook links" in scan.passive_features
    assert scan.blocking_features == ("VBA project",)
    assert not scan.safe_for_external_engine


# --- Step 3: number-format/date-system reconnaissance ---------------------


def _workbook_bin(path: Path) -> bytes:
    with zipfile.ZipFile(path) as archive:
        return archive.read("xl/workbook.bin")


def test_date_system_defaults_to_1900_and_reads_1904_flag(tmp_path: Path) -> None:
    path_1900 = tmp_path / "epoch1900.xlsb"
    write_xlsb(path_1900, {"Data": [[1.0]]})
    assert parse_xlsb_date_system(_workbook_bin(path_1900)) is False

    path_1904 = tmp_path / "epoch1904.xlsb"
    write_xlsb(path_1904, {"Data": [[1.0]]}, date1904=True)
    assert parse_xlsb_date_system(_workbook_bin(path_1904)) is True


def test_missing_workbook_prop_record_degrades_to_1900_not_a_crash(tmp_path: Path) -> None:
    path = tmp_path / "no-wbprop.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    with zipfile.ZipFile(path) as archive:
        workbook_bin = archive.read("xl/workbook.bin")
    # Strip every BrtWbProp (0x0199 naive-decoded) record out of the raw stream.
    stripped = workbook_bin.replace(_record(0x0199, struct.pack("<I", 0)), b"")
    assert stripped != workbook_bin
    assert parse_xlsb_date_system(stripped) is False


def test_parse_xlsb_styles_resolves_custom_and_builtin_formats(tmp_path: Path) -> None:
    path = tmp_path / "styled.xlsb"
    write_xlsb(
        path,
        {"Data": [[1.0]]},
        custom_formats={165: "mmm-yy", 166: "0.0%"},
        cell_xfs=[0, 14, 165, 166],
    )
    with zipfile.ZipFile(path) as archive:
        styles_data = archive.read("xl/styles.bin")

    table = parse_xlsb_styles(styles_data)

    assert table.custom_formats == {165: "mmm-yy", 166: "0.0%"}
    assert table.cell_xf_format_ids == (0, 14, 165, 166)
    assert table.number_format(0) == "General"
    assert table.number_format(1) == "mm-dd-yy"  # built-in id 14
    assert table.number_format(2) == "mmm-yy"  # custom
    assert table.number_format(3) == "0.0%"  # custom
    assert table.number_format(99) is None  # out of range: unresolvable, not a crash


def test_styles_bin_count_mismatch_raises_instead_of_misreading(tmp_path: Path) -> None:
    # A BrtBeginFmts declaring 2 entries but only 1 BrtFmt record follows.
    payload = _record(0x04E7, struct.pack("<I", 2)) + _record(
        0x002C, struct.pack("<H", 200) + b"\x03\x00\x00\x00" + "abc".encode("utf-16-le")
    )
    with pytest.raises(XlsbFormulaScanError):
        parse_xlsb_styles(payload)


def test_worksheet_cell_styles_are_bounded_to_one_sheet(tmp_path: Path) -> None:
    path = tmp_path / "per-cell.xlsb"
    write_xlsb(
        path,
        {
            "Data": [
                [StyledCell(1.0, xf_index=0), StyledCell(2.0, xf_index=2)],
                [StyledCell(3.0, xf_index=1)],
            ]
        },
        cell_xfs=[0, 14, 165],
        custom_formats={165: "mmm-yy"},
    )
    targets = resolve_worksheet_targets(path.read_bytes())
    assert targets == {"Data": "xl/worksheets/sheet1.bin"}
    with zipfile.ZipFile(path) as archive:
        sheet_data = archive.read(targets["Data"])

    styles = parse_worksheet_cell_styles(sheet_data, "Data")

    assert styles == {(1, 1): 0, (1, 2): 2, (2, 1): 1}


def test_formula_backed_numeric_cache_carries_its_xf_index(tmp_path: Path) -> None:
    """Mirrors the client-cited scenario: a formula-cached numeric serial
    with a custom date-formatted XF, distinct from an adjacent percent row."""
    path = tmp_path / "formula-date.xlsb"
    write_xlsb(
        path,
        {
            "Data": [
                [StyledCell(0.4557, xf_index=1)],  # percent row
                [StyledCell(46023.0, xf_index=2, is_formula=True)],  # date-formula row
            ]
        },
        cell_xfs=[0, 9, 165],
        custom_formats={165: "mmm-yy"},
    )
    scan = scan_xlsb_formulas(path.read_bytes())
    assert scan.formula_cells["Data"] == frozenset({(2, 1)})

    targets = resolve_worksheet_targets(path.read_bytes())
    with zipfile.ZipFile(path) as archive:
        sheet_data = archive.read(targets["Data"])
    styles = parse_worksheet_cell_styles(sheet_data, "Data")
    assert styles == {(1, 1): 1, (2, 1): 2}
