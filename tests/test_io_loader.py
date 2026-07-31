"""IO layer tests: snapshot loading across formats, encryption, read-only guarantee."""

import datetime as dt
import hashlib
import io
import shutil
import struct
import zipfile
from copy import copy
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

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
from qc_tool.io.loader import (
    OOXMLWorkloadError,
    UnsupportedFormatError,
    _assess_ooxml_workload,
    _load_ooxml_oracle,
    _load_ooxml_streaming,
    _parse_pivots,
    load_workbook_snapshot,
)
from qc_tool.io.model import display_cell_value, serialize_cell_value
from qc_tool.io.ooxml_worksheet import (
    OOXMLMetadataError,
    WorkbookMetadata,
    WorksheetMetadata,
    parse_ooxml_worksheet_metadata,
)
from tests.fixtures.manifest_schema import FixtureManifest
from tests.fixtures.xlsb_writer import _record, write_xlsb


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _rewrite_package(path: Path, replacements: dict[str, bytes]) -> None:
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts.update(replacements)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)


def _add_shared_strings(path: Path) -> None:
    with zipfile.ZipFile(path) as archive:
        content_types = ElementTree.fromstring(archive.read("[Content_Types].xml"))
        relationships = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    content_type_namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
    relationship_namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
    ElementTree.SubElement(
        content_types,
        f"{{{content_type_namespace}}}Override",
        PartName="/xl/sharedStrings.xml",
        ContentType=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
        ),
    )
    ElementTree.SubElement(
        relationships,
        f"{{{relationship_namespace}}}Relationship",
        Id="rIdSharedStrings",
        Type=("http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"),
        Target="sharedStrings.xml",
    )
    worksheet = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A5"/>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="A2" t="s"><v>1</v></c></row>
    <row r="3"><c r="A3" t="s"><v>2</v></c></row>
    <row r="4"><c r="A4" t="s"><v>0</v></c></row>
    <row r="5"><c r="A5" t="s"><v>3</v></c></row>
  </sheetData>
</worksheet>"""
    shared_strings = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
     count="5" uniqueCount="4">
  <si><t>plain</t></si>
  <si><t/></si>
  <si><t xml:space="preserve"> spaced </t></si>
  <si><r><t>Rich</t></r><r><t xml:space="preserve"> text</t></r></si>
</sst>"""
    _rewrite_package(
        path,
        {
            "[Content_Types].xml": ElementTree.tostring(
                content_types,
                encoding="utf-8",
                xml_declaration=True,
            ),
            "xl/_rels/workbook.xml.rels": ElementTree.tostring(
                relationships,
                encoding="utf-8",
                xml_declaration=True,
            ),
            "xl/worksheets/sheet1.xml": worksheet,
            "xl/sharedStrings.xml": shared_strings,
        },
    )


def _streaming_snapshot(path: Path):
    return _load_ooxml_streaming(
        path.read_bytes(),
        source_name=path.name,
        file_format=path.suffix.lstrip("."),
    )


def _oracle_snapshot(path: Path):
    return _load_ooxml_oracle(
        path.read_bytes(),
        source_name=path.name,
        file_format=path.suffix.lstrip("."),
    )


def _workload_metadata(
    *,
    cell_count: int,
    max_row: int = 1,
    max_column: int = 1,
) -> WorkbookMetadata:
    return WorkbookMetadata(
        sheets=[
            WorksheetMetadata(
                name="Data",
                visibility="visible",
                part="xl/worksheets/sheet1.xml",
                max_row=max_row,
                max_column=max_column,
                cell_count=cell_count,
                xml_bytes=1_000,
            )
        ],
        shared_string_bytes=0,
        styles_bytes=1_000,
        style_count=2,
    )


def test_ooxml_workload_preflight_warns_without_truncating() -> None:
    workload = _assess_ooxml_workload(
        _workload_metadata(cell_count=1_000_000),
        source_name="large.xlsx",
        allow_large_workbook=False,
    )

    assert workload.degraded
    assert not workload.override_used
    assert "warning limit" in workload.detail


def test_ooxml_workload_preflight_refuses_and_requires_explicit_override() -> None:
    metadata = _workload_metadata(cell_count=5_000_000)

    with pytest.raises(OOXMLWorkloadError, match="allow_large_workbook"):
        _assess_ooxml_workload(
            metadata,
            source_name="too-large.xlsx",
            allow_large_workbook=False,
        )

    workload = _assess_ooxml_workload(
        metadata,
        source_name="too-large.xlsx",
        allow_large_workbook=True,
    )
    assert workload.degraded
    assert workload.override_used
    assert "override accepted" in workload.detail


def test_ooxml_workload_preflight_rejects_sparse_pathological_area() -> None:
    metadata = _workload_metadata(
        cell_count=2,
        max_row=1_048_576,
        max_column=16_384,
    )

    with pytest.raises(OOXMLWorkloadError, match="largest sheet area"):
        _assess_ooxml_workload(
            metadata,
            source_name="sparse.xlsx",
            allow_large_workbook=False,
        )


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


def test_xlsx_snapshot_ignores_false_small_dimension(tmp_path: Path) -> None:
    path = tmp_path / "false-dimension.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = "head"
    sheet["Z40"] = 42
    workbook.save(path)
    with zipfile.ZipFile(path) as archive:
        worksheet = archive.read("xl/worksheets/sheet1.xml")
    _rewrite_package(
        path,
        {
            "xl/worksheets/sheet1.xml": worksheet.replace(
                b'ref="A1:Z40"',
                b'ref="A1:A1"',
            )
        },
    )

    snapshot = load_workbook_snapshot(path)

    loaded = snapshot.sheet(sheet.title)
    assert (loaded.max_row, loaded.max_column) == (40, 26)
    tail = loaded.cell("Z40")
    assert tail is not None
    assert tail.value == 42
    raw = parse_ooxml_worksheet_metadata(path.read_bytes()).sheets[0]
    assert (raw.declared_max_row, raw.declared_max_column) == (1, 1)
    assert (raw.max_row, raw.max_column, raw.cell_count) == (40, 26, 2)
    assert _streaming_snapshot(path) == snapshot


def test_streaming_empty_sheet_bounds_match_oracle(tmp_path: Path) -> None:
    path = tmp_path / "empty.xlsx"
    Workbook().save(path)

    oracle = _oracle_snapshot(path)
    streaming = _streaming_snapshot(path)

    assert streaming == oracle
    assert (streaming.sheets[0].max_row, streaming.sheets[0].max_column) == (1, 1)


def test_xlsx_snapshot_derives_missing_dimension_and_hidden_state(tmp_path: Path) -> None:
    path = tmp_path / "missing-dimension.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["D8"] = "tail"
    sheet.row_dimensions[8].hidden = True
    sheet.column_dimensions["C"].hidden = True
    workbook.save(path)
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    for child in list(root):
        if child.tag.rsplit("}", 1)[-1] == "dimension":
            root.remove(child)
    _rewrite_package(
        path,
        {
            "xl/worksheets/sheet1.xml": ElementTree.tostring(
                root,
                encoding="utf-8",
                xml_declaration=True,
            )
        },
    )

    snapshot = load_workbook_snapshot(path)

    loaded = snapshot.sheet(sheet.title)
    assert (loaded.max_row, loaded.max_column) == (8, 4)
    assert loaded.hidden_rows == frozenset({8})
    assert loaded.hidden_columns == frozenset({3})
    assert loaded.cell("D8") is not None
    raw = parse_ooxml_worksheet_metadata(path.read_bytes()).sheets[0]
    assert (raw.declared_max_row, raw.declared_max_column) == (None, None)
    assert (raw.max_row, raw.max_column) == (8, 4)
    assert raw.hidden_rows == frozenset({8})
    assert raw.hidden_columns == frozenset({3})
    assert _streaming_snapshot(path) == snapshot


def test_xlsx_snapshot_reads_genuine_shared_strings(tmp_path: Path) -> None:
    path = tmp_path / "shared-strings.xlsx"
    workbook = Workbook()
    workbook.save(path)
    _add_shared_strings(path)

    snapshot = load_workbook_snapshot(path)

    sheet = snapshot.sheets[0]
    values = []
    for row in range(1, 6):
        cell = sheet.cell(f"A{row}")
        assert cell is not None
        values.append(cell.value)
    assert values == [
        "plain",
        "",
        " spaced ",
        "plain",
        "Rich text",
    ]
    assert _streaming_snapshot(path) == snapshot


def test_xlsx_snapshot_preserves_shared_formulas_and_cached_values(
    tmp_path: Path,
) -> None:
    path = tmp_path / "shared-formulas.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    for row in range(1, 4):
        sheet.cell(row, 1, f"=ROW()+{row}")
    workbook.save(path)
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    cells = [element for element in root.iter() if element.tag.endswith("}c")]
    for index, cell in enumerate(cells):
        formula = next(element for element in cell if element.tag.endswith("}f"))
        value = next(element for element in cell if element.tag.endswith("}v"))
        formula.set("t", "shared")
        formula.set("si", "0")
        if index == 0:
            formula.set("ref", "A1:A3")
            formula.text = "ROW()+1"
        else:
            formula.text = None
        value.text = str(index + 2)
    _rewrite_package(
        path,
        {
            "xl/worksheets/sheet1.xml": ElementTree.tostring(
                root,
                encoding="utf-8",
                xml_declaration=True,
            )
        },
    )

    oracle = _oracle_snapshot(path)
    streaming = _streaming_snapshot(path)

    assert streaming == oracle
    assert [streaming.sheets[0].cells[(row, 1)].formula for row in range(1, 4)] == [
        "=ROW()+1",
        "=ROW()+1",
        "=ROW()+1",
    ]
    assert [streaming.sheets[0].cells[(row, 1)].value for row in range(1, 4)] == [
        2,
        3,
        4,
    ]


def test_xlsx_snapshot_preserves_array_formula_text_and_declared_range(
    tmp_path: Path,
) -> None:
    path = tmp_path / "array-formula.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = 1
    sheet["A2"] = 2
    sheet["B1"] = "=SUM(A1:A2)"
    workbook.save(path)
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    formula = next(element for element in root.iter() if element.tag.endswith("}f"))
    formula.set("t", "array")
    formula.set("ref", "B1:B2")
    formula.set("aca", "1")
    _rewrite_package(
        path,
        {
            "xl/worksheets/sheet1.xml": ElementTree.tostring(
                root,
                encoding="utf-8",
                xml_declaration=True,
            )
        },
    )

    oracle = _oracle_snapshot(path)
    streaming = _streaming_snapshot(path)

    assert streaming == oracle
    anchor = streaming.sheets[0].cell("B1")
    assert anchor is not None and anchor.formula == "=SUM(A1:A2)"
    assert len(streaming.formula_ranges) == 1
    descriptor = streaming.formula_ranges[0]
    assert descriptor.sheet == "Sheet"
    assert (descriptor.anchor_row, descriptor.anchor_column) == (1, 2)
    assert descriptor.cell_range == "B1:B2"
    assert descriptor.formula_type == "array"
    assert descriptor.always_calculate is True


def test_xlsx_snapshot_rejects_formula_range_with_wrong_anchor(tmp_path: Path) -> None:
    path = tmp_path / "bad-array-anchor.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["B2"] = "=SUM(A1:A2)"
    workbook.save(path)
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    formula = next(element for element in root.iter() if element.tag.endswith("}f"))
    formula.set("t", "array")
    formula.set("ref", "A1:B2")
    _rewrite_package(
        path,
        {
            "xl/worksheets/sheet1.xml": ElementTree.tostring(
                root,
                encoding="utf-8",
                xml_declaration=True,
            )
        },
    )

    with pytest.raises(OOXMLMetadataError, match="not the top-left"):
        load_workbook_snapshot(path)


def test_xlsx_snapshot_preserves_typed_formula_cached_values(tmp_path: Path) -> None:
    path = tmp_path / "typed-formula-cache.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    formulas = {
        "A1": ("=1+1", "n", "2"),
        "B1": ("=TRUE()", "b", "1"),
        "C1": ('="ready"', "str", "ready"),
        "D1": ("=1/0", "e", "#DIV/0!"),
    }
    for coordinate, (formula, _data_type, _cached) in formulas.items():
        sheet[coordinate] = formula
    workbook.save(path)
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    for cell in root.iter():
        if not cell.tag.endswith("}c") or cell.get("r") not in formulas:
            continue
        _formula, data_type, cached = formulas[str(cell.get("r"))]
        cell.set("t", data_type)
        value = next(element for element in cell if element.tag.endswith("}v"))
        value.text = cached
    _rewrite_package(
        path,
        {
            "xl/worksheets/sheet1.xml": ElementTree.tostring(
                root,
                encoding="utf-8",
                xml_declaration=True,
            )
        },
    )

    oracle = _oracle_snapshot(path)
    streaming = _streaming_snapshot(path)

    assert streaming == oracle
    assert [streaming.sheets[0].cells[(1, column)].value for column in range(1, 5)] == [
        2,
        True,
        "ready",
        "#DIV/0!",
    ]


def test_xlsx_snapshot_preserves_shared_string_formula_cache(tmp_path: Path) -> None:
    path = tmp_path / "shared-string-formula-cache.xlsx"
    Workbook().save(path)
    _add_shared_strings(path)
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    cell = next(
        element
        for element in root.iter()
        if element.tag.endswith("}c") and element.get("r") == "A1"
    )
    formula = ElementTree.Element(cell.tag.rsplit("}", 1)[0] + "}f")
    formula.text = '"plain"'
    cell.insert(0, formula)
    _rewrite_package(
        path,
        {
            "xl/worksheets/sheet1.xml": ElementTree.tostring(
                root,
                encoding="utf-8",
                xml_declaration=True,
            )
        },
    )

    oracle = _oracle_snapshot(path)
    streaming = _streaming_snapshot(path)

    assert streaming == oracle
    cached = streaming.sheets[0].cells[(1, 1)]
    assert cached.formula == '="plain"'
    assert cached.value == "plain"


def test_streaming_interns_equal_cell_styles(tmp_path: Path) -> None:
    path = tmp_path / "styles.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = 1
    sheet["B1"] = 2
    for cell in (sheet["A1"], sheet["B1"]):
        cell.number_format = "0.00"
        font = copy(cell.font)
        font.bold = True
        cell.font = font
    workbook.save(path)

    snapshot = _streaming_snapshot(path)

    first = snapshot.sheets[0].cells[(1, 1)]
    second = snapshot.sheets[0].cells[(1, 2)]
    assert first.style_key is second.style_key
    assert first.number_format is second.number_format


def test_xlsx_snapshot_expands_hidden_column_spans(tmp_path: Path) -> None:
    path = tmp_path / "hidden-span.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = "head"
    sheet.column_dimensions.group("C", "E", hidden=True)
    workbook.save(path)

    snapshot = load_workbook_snapshot(path)

    assert snapshot.sheets[0].hidden_columns == frozenset({3, 4, 5})
    assert _streaming_snapshot(path) == snapshot


@pytest.mark.parametrize("metadata_kind", ["row", "column"])
def test_raw_metadata_errors_include_sheet_context(
    tmp_path: Path,
    metadata_kind: str,
) -> None:
    path = tmp_path / f"malformed-{metadata_kind}.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet["C8"] = 1
    sheet.row_dimensions[8].hidden = True
    sheet.column_dimensions["C"].hidden = True
    workbook.save(path)
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    if metadata_kind == "row":
        row = next(
            element
            for element in root.iter()
            if element.tag.endswith("}row") and element.get("hidden") == "1"
        )
        row.set("r", "not-a-row")
    else:
        column = next(
            element
            for element in root.iter()
            if element.tag.endswith("}col") and element.get("hidden") == "1"
        )
        column.set("min", "not-a-column")
    _rewrite_package(
        path,
        {
            "xl/worksheets/sheet1.xml": ElementTree.tostring(
                root,
                encoding="utf-8",
                xml_declaration=True,
            )
        },
    )

    with pytest.raises(OOXMLMetadataError, match="Data: invalid hidden"):
        parse_ooxml_worksheet_metadata(path.read_bytes())


@pytest.mark.parametrize("name", ["baseline.xlsx", "current.xlsx"])
def test_streaming_xlsx_snapshot_matches_oracle(
    fixture_dir: Path,
    name: str,
) -> None:
    path = fixture_dir / name

    assert _streaming_snapshot(path) == _oracle_snapshot(path)


def test_streaming_closes_formula_workbook_before_cached_value_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    package = io.BytesIO()
    Workbook().save(package)
    original_load_workbook = loader_module.load_workbook

    class TrackingWorkbook:
        def __init__(self, workbook: Any) -> None:
            self.workbook = workbook
            self.closed = False

        def __getattr__(self, name: str) -> Any:
            return getattr(self.workbook, name)

        def __getitem__(self, name: str) -> Any:
            return self.workbook[name]

        def close(self) -> None:
            self.closed = True
            self.workbook.close()

    formula_workbook: TrackingWorkbook | None = None
    def fake_load_workbook(*args: Any, **kwargs: Any) -> Any:
        nonlocal formula_workbook
        formula_workbook = TrackingWorkbook(original_load_workbook(*args, **kwargs))
        return formula_workbook

    def fail_cached_values(*args: Any, **kwargs: Any) -> None:
        assert formula_workbook is not None
        assert formula_workbook.closed
        raise ValueError("cached-value stream failed")

    monkeypatch.setattr(loader_module, "load_workbook", fake_load_workbook)
    monkeypatch.setattr(
        loader_module,
        "_merge_cached_formula_values",
        fail_cached_values,
    )

    with pytest.raises(ValueError, match="cached-value stream"):
        loader_module._load_ooxml_streaming(
            package.getvalue(),
            source_name="broken.xlsx",
            file_format="xlsx",
        )

    assert formula_workbook is not None
    assert formula_workbook.closed


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
    raw = parse_ooxml_worksheet_metadata(path.read_bytes())
    assert raw.sheets[0].tables == snapshot.tables
    assert _streaming_snapshot(path) == snapshot


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
        archive.writestr("xl/pivotCache/pivotCacheDefinition2.xml", cache("Data", "A1:B20"))
        archive.writestr("xl/pivotCache/pivotCacheDefinition10.xml", cache("OtherData", "D1:E8"))

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
