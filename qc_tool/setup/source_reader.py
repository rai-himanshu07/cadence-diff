"""Setup-only workbook readers with no full-QC or Office-adapter work."""

from __future__ import annotations

import contextlib
import datetime as dt
import io
import zipfile
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Protocol
from xml.etree import ElementTree

from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format, is_timedelta_format
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
from openpyxl.utils.datetime import MAC_EPOCH, WINDOWS_EPOCH, from_excel, from_ISO8601
from pyxlsb import open_workbook as open_xlsb

from qc_tool.io.decrypt import open_decrypted
from qc_tool.io.model import CellRecord, CellValue, SheetSnapshot, is_cell_value
from qc_tool.io.opc import resolve_internal_relationship_target
from qc_tool.io.xlsb_formula import (
    XlsbFormulaScanError,
    parse_worksheet_cell_styles,
    parse_xlsb_date_system,
    parse_xlsb_styles,
    resolve_worksheet_targets,
    scan_xlsb_formulas,
)
from qc_tool.setup.models import XlsbRiskProfile

MAX_FORMULA_TEXT_CHARS = 256

_XLSB_ERRORS = {
    "0x0": "#NULL!",
    "0x7": "#DIV/0!",
    "0xf": "#VALUE!",
    "0x17": "#REF!",
    "0x1d": "#NAME?",
    "0x24": "#NUM!",
    "0x2a": "#N/A",
}


class UnsupportedSetupSourceError(ValueError):
    """The source format is not supported by setup inspection."""


class SetupSourceReader(Protocol):
    """A one-shot, per-sheet setup source."""

    file_format: str
    sheet_count: int
    xlsb_risk: XlsbRiskProfile | None

    def iter_sheets(self) -> Iterator[SheetSnapshot]: ...


@dataclass(slots=True)
class _SetupSource:
    file_format: str
    sheet_count: int
    xlsb_risk: XlsbRiskProfile | None
    _iterator: Iterator[SheetSnapshot]
    _consumed: bool = False

    def iter_sheets(self) -> Iterator[SheetSnapshot]:
        if self._consumed:
            raise RuntimeError("setup source sheets may only be consumed once")
        self._consumed = True
        return self._iterator

    def close(self) -> None:
        close = getattr(self._iterator, "close", None)
        if callable(close):
            close()


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _relationship_id(element: ElementTree.Element) -> str | None:
    return next(
        (value for name, value in element.attrib.items() if _local_name(name) == "id"),
        None,
    )


def _ooxml_sheet_specs(
    archive: zipfile.ZipFile,
) -> tuple[list[tuple[str, str, str]], dt.datetime]:
    workbook_part = "xl/workbook.xml"
    workbook = ElementTree.fromstring(archive.read(workbook_part))
    relationships = ElementTree.fromstring(
        archive.read("xl/_rels/workbook.xml.rels")
    )
    targets: dict[str, str] = {}
    for relationship in relationships:
        relationship_id = relationship.get("Id")
        target = relationship.get("Target")
        if (
            relationship_id
            and target
            and relationship.get("TargetMode") != "External"
        ):
            targets[relationship_id] = resolve_internal_relationship_target(
                workbook_part, target
            )
    epoch = WINDOWS_EPOCH
    for element in workbook.iter():
        if _local_name(element.tag) == "workbookPr" and str(
            element.get("date1904", "")
        ).casefold() in {"1", "true", "on"}:
            epoch = MAC_EPOCH
            break
    sheets: list[tuple[str, str, str]] = []
    for element in workbook.iter():
        if _local_name(element.tag) != "sheet":
            continue
        name = element.get("name")
        relationship_id = _relationship_id(element)
        if not name or not relationship_id or relationship_id not in targets:
            raise ValueError("workbook contains an unresolved worksheet")
        sheets.append((name, element.get("state", "visible"), targets[relationship_id]))
    return sheets, epoch


def _ooxml_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    part = "xl/sharedStrings.xml"
    if part not in archive.namelist():
        return []
    root = ElementTree.fromstring(archive.read(part))
    return [
        "".join(
            child.text or ""
            for child in item.iter()
            if _local_name(child.tag) == "t"
        )
        for item in root
        if _local_name(item.tag) == "si"
    ]


def _ooxml_number_formats(archive: zipfile.ZipFile) -> list[str | None]:
    part = "xl/styles.xml"
    if part not in archive.namelist():
        return []
    root = ElementTree.fromstring(archive.read(part))
    custom = {
        int(element.get("numFmtId", "0")): element.get("formatCode", "General")
        for element in root.iter()
        if _local_name(element.tag) == "numFmt"
    }
    cell_xfs = next(
        (element for element in root.iter() if _local_name(element.tag) == "cellXfs"),
        None,
    )
    if cell_xfs is None:
        return []
    return [
        custom.get(
            int(element.get("numFmtId", "0")),
            BUILTIN_FORMATS.get(int(element.get("numFmtId", "0")), "General"),
        )
        for element in cell_xfs
        if _local_name(element.tag) == "xf"
    ]


def _numeric_value(raw: str) -> int | float:
    return float(raw) if "." in raw or "E" in raw.upper() else int(raw)


def _typed_numeric(
    raw: str,
    *,
    number_format: str | None,
    epoch: dt.datetime,
) -> CellValue:
    numeric = _numeric_value(raw)
    if number_format and is_date_format(number_format):
        converted = from_excel(
            numeric,
            epoch,
            timedelta=is_timedelta_format(number_format),
        )
        if is_cell_value(converted):
            return converted
    return numeric


def _ooxml_cell_value(
    element: ElementTree.Element,
    *,
    shared_strings: list[str],
    number_format: str | None,
    epoch: dt.datetime,
) -> CellValue:
    data_type = element.get("t", "n")
    if data_type == "inlineStr":
        return "".join(
            child.text or ""
            for child in element.iter()
            if _local_name(child.tag) == "t"
        )
    raw = next(
        (child.text for child in element if _local_name(child.tag) == "v"),
        None,
    )
    if raw is None:
        return None
    if data_type == "s":
        try:
            return shared_strings[int(raw)]
        except (IndexError, ValueError) as exc:
            raise ValueError("invalid shared-string index") from exc
    if data_type == "b":
        return bool(int(raw))
    if data_type in {"e", "str"}:
        return raw
    if data_type == "d":
        converted = from_ISO8601(raw)
        return converted if is_cell_value(converted) else raw
    return _typed_numeric(raw, number_format=number_format, epoch=epoch)


def _ooxml_sheet(
    archive: zipfile.ZipFile,
    *,
    name: str,
    visibility: str,
    part: str,
    shared_strings: list[str],
    number_formats: list[str | None],
    epoch: dt.datetime,
) -> SheetSnapshot:
    cells: dict[tuple[int, int], CellRecord] = {}
    hidden_rows: set[int] = set()
    hidden_columns: set[int] = set()
    max_row = max_column = 0
    declared_max_row = declared_max_column = 0
    cell_depth = 0
    try:
        with archive.open(part) as stream:
            for event, element in ElementTree.iterparse(stream, events=("start", "end")):
                local_name = _local_name(element.tag)
                if event == "start":
                    if cell_depth:
                        cell_depth += 1
                    elif local_name == "c":
                        cell_depth = 1
                    continue
                if cell_depth:
                    cell_depth -= 1
                    if cell_depth:
                        continue
                    reference = element.get("r")
                    if not reference:
                        raise ValueError("worksheet cell is missing its coordinate")
                    row, column = coordinate_to_tuple(reference)
                    style_index = int(element.get("s", "0"))
                    number_format = (
                        number_formats[style_index]
                        if 0 <= style_index < len(number_formats)
                        else None
                    )
                    formula_element = next(
                        (child for child in element if _local_name(child.tag) == "f"),
                        None,
                    )
                    is_formula = formula_element is not None
                    formula = None
                    if formula_element is not None and formula_element.text:
                        formula = formula_element.text
                        if not formula.startswith("="):
                            formula = f"={formula}"
                        formula = formula[:MAX_FORMULA_TEXT_CHARS]
                    value = _ooxml_cell_value(
                        element,
                        shared_strings=shared_strings,
                        number_format=number_format,
                        epoch=epoch,
                    )
                    if value is not None or is_formula:
                        cells[(row, column)] = CellRecord(
                            row=row,
                            column=column,
                            value=value,
                            formula=formula,
                            is_formula=is_formula,
                            number_format=number_format,
                        )
                    max_row = max(max_row, row)
                    max_column = max(max_column, column)
                    element.clear()
                    continue
                if local_name == "dimension":
                    reference = element.get("ref")
                    if reference:
                        try:
                            _, _, declared_max_column, declared_max_row = range_boundaries(
                                reference
                            )
                        except ValueError:
                            declared_max_row = declared_max_column = 0
                elif local_name == "row" and str(
                    element.get("hidden", "")
                ).casefold() in {"1", "true", "on"}:
                    raw_row = element.get("r")
                    if raw_row:
                        hidden_rows.add(int(raw_row))
                elif local_name == "col" and str(
                    element.get("hidden", "")
                ).casefold() in {"1", "true", "on"}:
                    minimum = int(element.get("min", "0"))
                    maximum = int(element.get("max", "0"))
                    if minimum > 0 and maximum >= minimum:
                        hidden_columns.update(range(minimum, maximum + 1))
                element.clear()
    except ElementTree.ParseError as exc:
        raise ValueError("worksheet XML is malformed") from exc
    return SheetSnapshot(
        name=name,
        visibility=visibility,
        max_row=max(max_row, declared_max_row or 0, 1),
        max_column=max(max_column, declared_max_column or 0, 1),
        cells=cells,
        hidden_rows=frozenset(hidden_rows),
        hidden_columns=frozenset(hidden_columns),
    )


def _iter_ooxml_sheets(
    data: bytes,
    sheet_specs: list[tuple[str, str, str]],
    epoch: dt.datetime,
) -> Iterator[SheetSnapshot]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        shared_strings = _ooxml_shared_strings(archive)
        number_formats = _ooxml_number_formats(archive)
        for name, visibility, part in sheet_specs:
            yield _ooxml_sheet(
                archive,
                name=name,
                visibility=visibility,
                part=part,
                shared_strings=shared_strings,
                number_formats=number_formats,
                epoch=epoch,
            )


def _xlsb_value(raw: object) -> CellValue:
    if isinstance(raw, str):
        return _XLSB_ERRORS.get(raw.casefold(), raw)
    if isinstance(raw, bool | int | float) or raw is None:
        return raw
    return str(raw)


def _kernel_xlsb_value(
    number: float | None,
    boolean: bool | None,
    text: str | None,
) -> CellValue:
    if boolean is not None:
        return boolean
    if number is not None:
        return number
    if text is not None:
        return _XLSB_ERRORS.get(text.casefold(), text)
    return None


def _pyxlsb_sheet_cells(
    workbook: Any, sheet_name: str
) -> Iterator[tuple[int, int, CellValue]]:
    with workbook.get_sheet(sheet_name) as worksheet:
        for row in worksheet.rows(sparse=True):
            for raw_cell in row:
                if raw_cell.v is not None:
                    yield raw_cell.r, raw_cell.c, _xlsb_value(raw_cell.v)


def _native_xlsb_cells(
    data: bytes,
) -> dict[
    str,
    list[tuple[int, int, float | None, bool | None, str | None]],
] | None:
    from qc_tool.io.native_kernel import native_kernel_available, raw_values_report

    if not native_kernel_available():
        return None
    try:
        return dict(raw_values_report(data))
    except (SystemExit, KeyboardInterrupt, GeneratorExit):
        raise
    except BaseException:
        return None


def _iter_xlsb_sheets(
    data: bytes,
    formula_cells: dict[str, frozenset[tuple[int, int]]],
    sheet_visibility: dict[str, str],
) -> Iterator[SheetSnapshot]:
    date1904 = False
    style_table = None
    targets: dict[str, str] = {}
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            date1904 = parse_xlsb_date_system(archive.read("xl/workbook.bin"))
            if "xl/styles.bin" in archive.namelist():
                style_table = parse_xlsb_styles(archive.read("xl/styles.bin"))
            targets = resolve_worksheet_targets(data)
    except (KeyError, XlsbFormulaScanError, zipfile.BadZipFile):
        style_table = None
        targets = {}
    epoch = MAC_EPOCH if date1904 else WINDOWS_EPOCH
    with open_xlsb(io.BytesIO(data)) as workbook, zipfile.ZipFile(
        io.BytesIO(data)
    ) as archive:
        native_cells = _native_xlsb_cells(data)
        if native_cells is not None and tuple(native_cells) != tuple(workbook.sheets):
            native_cells = None
        for sheet_name in workbook.sheets:
            cells: dict[tuple[int, int], CellRecord] = {}
            max_row = max_column = 0
            styles: dict[tuple[int, int], int] = {}
            target = targets.get(sheet_name)
            if style_table is not None and target:
                try:
                    styles = parse_worksheet_cell_styles(
                        archive.read(target), sheet_name
                    )
                except (KeyError, XlsbFormulaScanError):
                    styles = {}
            sheet_cells = (
                (
                    (row, column, _kernel_xlsb_value(number, boolean, text))
                    for row, column, number, boolean, text in native_cells.pop(
                        sheet_name
                    )
                )
                if native_cells is not None
                else _pyxlsb_sheet_cells(workbook, sheet_name)
            )
            for row, column, value in sheet_cells:
                row_number = row + 1
                column_number = column + 1
                number_format = None
                style_index = styles.get((row_number, column_number))
                if style_table is not None and style_index is not None:
                    number_format = style_table.number_format(style_index)
                    if (
                        number_format
                        and isinstance(value, int | float)
                        and not isinstance(value, bool)
                        and is_date_format(number_format)
                    ):
                        converted = from_excel(
                            value,
                            epoch,
                            timedelta=is_timedelta_format(number_format),
                        )
                        if is_cell_value(converted):
                            value = converted
                cells[(row_number, column_number)] = CellRecord(
                    row=row_number,
                    column=column_number,
                    value=value,
                    is_formula=(row_number, column_number)
                    in formula_cells.get(sheet_name, frozenset()),
                    number_format=number_format,
                )
                max_row = max(max_row, row_number)
                max_column = max(max_column, column_number)
            for row_number, column_number in formula_cells.get(
                sheet_name, frozenset()
            ).difference(cells):
                cells[(row_number, column_number)] = CellRecord(
                    row=row_number,
                    column=column_number,
                    value=None,
                    is_formula=True,
                )
                max_row = max(max_row, row_number)
                max_column = max(max_column, column_number)
            yield SheetSnapshot(
                name=sheet_name,
                visibility=sheet_visibility.get(sheet_name, "visible"),
                max_row=max(max_row, 1),
                max_column=max(max_column, 1),
                cells=cells,
            )


def _iter_xlsb_source(data: bytes, source: _SetupSource) -> Iterator[SheetSnapshot]:
    try:
        scan = scan_xlsb_formulas(data)
    except XlsbFormulaScanError:
        formula_cells: dict[str, frozenset[tuple[int, int]]] = {}
        sheet_visibility: dict[str, str] = {}
    else:
        formula_cells = scan.formula_cells
        sheet_visibility = scan.sheet_visibility
        source.xlsb_risk = XlsbRiskProfile(
            risky_features=scan.risky_features,
            passive_features=scan.passive_features,
            blocking_features=scan.blocking_features,
            unknown_external_features=scan.unknown_external_features,
            safe_for_external_engine=scan.safe_for_external_engine,
        )
    yield from _iter_xlsb_sheets(data, formula_cells, sheet_visibility)


@contextlib.contextmanager
def open_setup_source(
    path: Path, password: str | None = None
) -> Iterator[SetupSourceReader]:
    """Open one source once and yield a one-shot setup-only sheet reader."""
    data = open_decrypted(path, password).getvalue()
    suffix = path.suffix.casefold()
    if suffix in {".xlsx", ".xlsm"}:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            sheet_specs, epoch = _ooxml_sheet_specs(archive)
        source = _SetupSource(
            suffix.removeprefix("."),
            len(sheet_specs),
            None,
            _iter_ooxml_sheets(data, sheet_specs, epoch),
        )
    elif suffix == ".xlsb":
        try:
            sheet_count = len(resolve_worksheet_targets(data))
        except XlsbFormulaScanError:
            sheet_count = 0
        source = _SetupSource(
            "xlsb",
            sheet_count,
            None,
            iter(()),
        )
        source._iterator = _iter_xlsb_source(data, source)
    else:
        raise UnsupportedSetupSourceError(
            "source format is not supported for setup"
        )
    try:
        yield source
    finally:
        source.close()
