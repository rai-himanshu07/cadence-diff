"""IO layer tests: snapshot loading across formats, encryption, read-only guarantee."""

import datetime as dt
import hashlib
import io
import shutil
import struct
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.utils.datetime import to_excel
from openpyxl.worksheet.table import Table
from pyxlsb import biff12

import qc_tool.io.loader as loader_module
from qc_tool.io.decrypt import (
    InvalidPasswordError,
    PasswordRequiredError,
    is_encrypted,
)
from qc_tool.io.formula_enrichment import FormulaExtraction
from qc_tool.io.loader import UnsupportedFormatError, _parse_pivots, load_workbook_snapshot
from qc_tool.io.model import display_cell_value, serialize_cell_value
from tests.fixtures.manifest_schema import FixtureManifest
from tests.fixtures.xlsb_writer import _record, write_xlsb


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_sources_unmodified_by_loading(fixture_dir: Path, manifest: FixtureManifest) -> None:
    targets = {
        "baseline.xlsx": None,
        "current.xlsx": None,
        "baseline.xlsb": None,
        "current.xlsb": None,
        "current_encrypted.xlsx": manifest.password,
    }
    before = {name: _sha256(fixture_dir / name) for name in targets}
    for name, password in targets.items():
        load_workbook_snapshot(fixture_dir / name, password=password)
    after = {name: _sha256(fixture_dir / name) for name in targets}
    assert before == after


def test_xlsx_snapshot_content(fixture_dir: Path, manifest: FixtureManifest) -> None:
    snap = load_workbook_snapshot(fixture_dir / "current.xlsx")
    assert snap.file_format == "xlsx"
    assert snap.formulas_available and snap.styles_available
    assert snap.sheet_names == [
        "Long_Monthly",
        "Wide_Weekly",
        "Dashboard",
        "Summary",
        "New_Analysis",
        "Params",
    ]

    long_monthly = snap.sheet("Long_Monthly")
    assert long_monthly.max_row == 25

    e01 = manifest.defect("E01")
    cell = long_monthly.cell("C7")
    assert cell is not None and cell.value == float(e01.current or "")

    hardcoded = long_monthly.cell("E10")  # E02
    assert hardcoded is not None
    assert hardcoded.formula is None and isinstance(hardcoded.value, float)

    logic_changed = long_monthly.cell("E14")  # E03
    assert logic_changed is not None and logic_changed.formula == "=C14-D14*1.1"
    assert logic_changed.has_formula
    assert logic_changed.value is None  # openpyxl-authored file has no cached results

    assert long_monthly.cell("E25") is None  # E04: not extended

    summary = snap.sheet("Summary")
    error_const = summary.cell("C5")  # E05
    assert error_const is not None and error_const.is_error and error_const.value == "#REF!"
    error_formula = summary.cell("D5")  # E13
    assert error_formula is not None and error_formula.formula == "=SUM(#REF!)"

    assert snap.sheet("Params").visibility == "hidden"  # E11


def test_xlsx_snapshot_preserves_typed_dates(tmp_path: Path) -> None:
    path = tmp_path / "typed-date.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = dt.date(2026, 4, 1)
    sheet["A1"].number_format = "mmm-yy"
    workbook.save(path)

    snapshot = load_workbook_snapshot(path)

    cell = snapshot.sheet(sheet.title).cell("A1")
    assert cell is not None
    assert cell.value == dt.datetime(2026, 4, 1)
    assert display_cell_value(cell.value) == "2026-04-01T00:00:00"
    assert serialize_cell_value(cell.value) == "2026-04-01T00:00:00"


def test_xlsx_snapshot_preserves_formula_cached_dates(tmp_path: Path) -> None:
    path = tmp_path / "formula-date.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = "=DATE(2026,4,1)"
    sheet["A1"].number_format = "mmm-yy"
    workbook.save(path)
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    worksheet_part = "xl/worksheets/sheet1.xml"
    xml = parts[worksheet_part].decode("utf-8")
    parts[worksheet_part] = xml.replace(
        "<v></v>",
        f"<v>{to_excel(dt.datetime(2026, 4, 1))}</v>",
    ).encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)

    snapshot = load_workbook_snapshot(path)

    cell = snapshot.sheet(sheet.title).cell("A1")
    assert cell is not None
    assert cell.formula == "=DATE(2026,4,1)"
    assert cell.value == dt.datetime(2026, 4, 1)


def test_xlsx_snapshot_structure(fixture_dir: Path, manifest: FixtureManifest) -> None:
    base = load_workbook_snapshot(fixture_dir / "baseline.xlsx")
    curr = load_workbook_snapshot(fixture_dir / "current.xlsx")

    named_base = {n.name: n.target for n in base.named_ranges}
    named_curr = {n.name: n.target for n in curr.named_ranges}
    assert named_base["KPI_Margin"] == "Dashboard!$B$4"
    assert named_curr["KPI_Margin"] == "Dashboard!$B$5"  # E12

    e16 = manifest.defect("E16")
    chart_base = next(c for c in base.charts if c.sheet == "Dashboard")
    chart_curr = next(c for c in curr.charts if c.sheet == "Dashboard")
    assert chart_base.title == "Revenue trend (workbook)"
    assert chart_base.series[0].values_ref == e16.baseline
    assert chart_curr.series[0].values_ref == e16.current

    e17 = manifest.defect("E17")
    assert base.pivots[0].name == "RevenuePivot"
    assert base.pivots[0].source_sheet == "Long_Monthly"
    assert base.pivots[0].source_ref == e17.baseline
    assert curr.pivots[0].source_ref == e17.current

    dash_base = base.sheet("Dashboard")
    dash_curr = curr.sheet("Dashboard")
    b4_base, b4_curr = dash_base.cell("B4"), dash_curr.cell("B4")
    assert b4_base is not None and b4_base.number_format == "0.0%"
    assert b4_curr is not None and b4_curr.number_format == "0.00"  # E09
    a2_base, a2_curr = dash_base.cell("A2"), dash_curr.cell("A2")
    assert a2_base is not None and a2_curr is not None
    assert a2_base.style_key != a2_curr.style_key  # E18


def test_xlsx_snapshot_captures_table_schema(tmp_path: Path) -> None:
    path = tmp_path / "table.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet.append(["Key", "Amount", "Cost"])
    sheet.append(["A", 10, 4])
    sheet.add_table(Table(displayName="MetricsTable", ref="A1:C2"))
    workbook.save(path)

    snapshot = load_workbook_snapshot(path)

    assert len(snapshot.tables) == 1
    table = snapshot.tables[0]
    assert (table.sheet, table.name, table.display_name) == (
        "Data",
        "MetricsTable",
        "MetricsTable",
    )
    assert table.cell_range == "A1:C2"
    assert table.columns == ["Key", "Amount", "Cost"]
    assert table.header_row_count == 1
    assert table.totals_row_count == 0


def test_pivots_resolve_shared_and_non_positional_caches() -> None:
    workbook = b"""<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <pivotCaches><pivotCache cacheId="1" r:id="rId10"/>
        <pivotCache cacheId="7" r:id="rId11"/></pivotCaches></workbook>"""
    relationships = b"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId10" Target="pivotCache/pivotCacheDefinition2.xml"/>
        <Relationship Id="rId11" Target="pivotCache/pivotCacheDefinition10.xml"/>
        </Relationships>"""

    def table(name: str, cache_id: int) -> bytes:
        return (
            '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/'
            f'spreadsheetml/2006/main" name="{name}" cacheId="{cache_id}">'
            '<location ref="A1:B2"/></pivotTableDefinition>'
        ).encode()

    def cache(sheet: str, ref: str) -> bytes:
        return (
            '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/'
            'spreadsheetml/2006/main"><cacheSource><worksheetSource '
            f'sheet="{sheet}" ref="{ref}"/></cacheSource></pivotCacheDefinition>'
        ).encode()

    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w") as archive:
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", relationships)
        archive.writestr("xl/pivotTables/pivotTable1.xml", table("Other", 7))
        archive.writestr("xl/pivotTables/pivotTable2.xml", table("Shared A", 1))
        archive.writestr("xl/pivotTables/pivotTable10.xml", table("Shared B", 1))
        archive.writestr(
            "xl/pivotCache/pivotCacheDefinition2.xml", cache("Data", "A1:B20")
        )
        archive.writestr(
            "xl/pivotCache/pivotCacheDefinition10.xml", cache("OtherData", "D1:E8")
        )

    pivots = {pivot.name: pivot for pivot in _parse_pivots(stream.getvalue())}

    assert (pivots["Shared A"].source_sheet, pivots["Shared A"].source_ref) == (
        "Data",
        "A1:B20",
    )
    assert (pivots["Shared B"].source_sheet, pivots["Shared B"].source_ref) == (
        "Data",
        "A1:B20",
    )
    assert (pivots["Other"].source_sheet, pivots["Other"].source_ref) == (
        "OtherData",
        "D1:E8",
    )


def test_xlsm_loads_identically(fixture_dir: Path, tmp_path: Path) -> None:
    source = fixture_dir / "current.xlsx"
    xlsm = tmp_path / "current.xlsm"
    shutil.copyfile(source, xlsm)

    snap_xlsx = load_workbook_snapshot(source)
    snap_xlsm = load_workbook_snapshot(xlsm)
    assert snap_xlsm.file_format == "xlsm"
    assert snap_xlsm.sheet_names == snap_xlsx.sheet_names
    assert snap_xlsm.sheet("Long_Monthly").cells == snap_xlsx.sheet("Long_Monthly").cells


def test_xlsb_snapshot(fixture_dir: Path, manifest: FixtureManifest) -> None:
    snap = load_workbook_snapshot(fixture_dir / "current.xlsb")
    assert snap.file_format == "xlsb"
    assert not snap.formulas_available and not snap.styles_available
    assert snap.formula_presence_available
    assert snap.sheet_names == ["Long_Monthly"]

    sheet = snap.sheet("Long_Monthly")
    assert sheet.max_row == 25
    header = sheet.cell("A1")
    assert header is not None and header.value == "Period"

    xb01 = manifest.defect("XB01")
    cell = sheet.cell("C7")
    assert cell is not None and cell.value == float(xb01.current or "")
    assert cell.formula is None and cell.style_key is None
    assert cell.is_formula is False and not cell.has_formula


def test_xlsb_enrichment_preserves_original_cached_value(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    path = tmp_path / "formula.xlsb"
    write_xlsb(path, {"Data": [[1.0]]})
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    sheet = bytearray()
    sheet += _record(biff12.WORKSHEET)
    sheet += _record(biff12.DIMENSION, struct.pack("<IIII", 0, 0, 0, 0))
    sheet += _record(biff12.SHEETDATA)
    sheet += _record(biff12.ROW, (0).to_bytes(4, "little"))
    formula_payload = (
        (0).to_bytes(4, "little")
        + (0).to_bytes(4, "little")
        + struct.pack("<d", 42.0)
        + b"\x00\x00"
        + (0).to_bytes(4, "little")
    )
    sheet += _record(biff12.FORMULA_FLOAT, formula_payload)
    sheet += _record(biff12.SHEETDATA_END)
    sheet += _record(biff12.WORKSHEET_END)
    parts["xl/worksheets/sheet1.bin"] = bytes(sheet)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)

    def fake_extract(data: bytes, scan: object) -> FormulaExtraction:
        return FormulaExtraction(
            formulas={"Data": {(1, 1): "=40+2"}},
            engine="libreoffice:test",
            detail="test formula adapter",
        )

    monkeypatch.setattr(loader_module, "_extract_xlsb_formulas", fake_extract)

    snapshot = load_workbook_snapshot(path)

    cell = snapshot.sheet("Data").cells[(1, 1)]
    assert cell.value == 42.0
    assert cell.formula == "=40+2" and cell.has_formula
    assert snapshot.formulas_available
    assert snapshot.formula_source == "libreoffice:test"


def test_encrypted_workbook(fixture_dir: Path, manifest: FixtureManifest) -> None:
    encrypted = fixture_dir / "current_encrypted.xlsx"
    assert is_encrypted(encrypted)
    assert not is_encrypted(fixture_dir / "current.xlsx")

    snap = load_workbook_snapshot(encrypted, password=manifest.password)
    plain = load_workbook_snapshot(fixture_dir / "current.xlsx")
    assert snap.sheet_names == plain.sheet_names
    assert snap.sheet("Long_Monthly").cells == plain.sheet("Long_Monthly").cells

    with pytest.raises(PasswordRequiredError):
        load_workbook_snapshot(encrypted)
    with pytest.raises(InvalidPasswordError):
        load_workbook_snapshot(encrypted, password="not-the-password")


def test_unsupported_format(tmp_path: Path) -> None:
    bogus = tmp_path / "file.xls"
    bogus.write_bytes(b"PK\x03\x04junk")
    with pytest.raises(UnsupportedFormatError):
        load_workbook_snapshot(bogus)
