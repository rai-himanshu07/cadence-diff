"""Defined-name scope extraction read straight from the OOXML workbook part.

openpyxl exposes workbook-scoped and sheet-scoped names through two different
collections and never records which sheet index the package actually declared.
Scope is only authoritative in ``xl/workbook.xml``, where ``definedName`` carries
``@localSheetId`` as a position in the package's own ``<sheets>`` order, so the
scan below reads that part directly instead of reconstructing scope afterwards.
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from xml.etree import ElementTree

from qc_tool.io.model import NamedRange

_WORKBOOK_PART = "xl/workbook.xml"
_BUILTIN_PREFIX = "_xlnm"
_TRUE = {"1", "true", "on"}


@dataclass(frozen=True, slots=True)
class DefinedNameScan:
    """What the workbook part declares about defined-name scope."""

    workbook_scoped: tuple[NamedRange, ...] = ()
    sheet_scoped: tuple[NamedRange, ...] = ()
    available: bool = False
    builtin_skipped: int = 0
    macro_skipped: int = 0
    unresolved_scopes: int = 0
    detail: str = ""


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _attribute(element: ElementTree.Element, name: str) -> str | None:
    for key, value in element.attrib.items():
        if _local_name(key) == name:
            return value
    return None


def _unavailable(detail: str) -> DefinedNameScan:
    return DefinedNameScan(available=False, detail=detail)


def scan_defined_names(data: bytes) -> DefinedNameScan:
    """Read every declared defined name, with its scope, from a workbook package."""
    try:
        archive = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        return _unavailable("workbook is not an OOXML package")
    with archive:
        if _WORKBOOK_PART not in archive.namelist():
            return _unavailable(f"{_WORKBOOK_PART} is not present in the package")
        try:
            root = ElementTree.fromstring(archive.read(_WORKBOOK_PART))
        except ElementTree.ParseError as exc:
            return _unavailable(f"{_WORKBOOK_PART} is not readable XML: {exc}")

    sheet_order: list[str] = []
    declarations: list[ElementTree.Element] = []
    for child in root:
        tag = _local_name(child.tag)
        if tag == "sheets":
            sheet_order = [
                _attribute(sheet, "name") or ""
                for sheet in child
                if _local_name(sheet.tag) == "sheet"
            ]
        elif tag == "definedNames":
            declarations = [
                item for item in child if _local_name(item.tag) == "definedName"
            ]

    workbook_scoped: list[NamedRange] = []
    sheet_scoped: list[NamedRange] = []
    builtin_skipped = 0
    macro_skipped = 0
    unresolved_scopes = 0

    for declaration in declarations:
        name = _attribute(declaration, "name")
        if not name:
            unresolved_scopes += 1
            continue
        if name.startswith(_BUILTIN_PREFIX):
            builtin_skipped += 1
            continue
        # A macro-bound name resolves to a procedure, not a range, so a range
        # comparison cannot represent it without inventing an invalid target.
        if (_attribute(declaration, "function") or "").lower() in _TRUE or (
            _attribute(declaration, "vbProcedure") or ""
        ).lower() in _TRUE:
            macro_skipped += 1
            continue
        hidden = (_attribute(declaration, "hidden") or "").lower() in _TRUE
        target = (declaration.text or "").strip()
        local_sheet_id = _attribute(declaration, "localSheetId")
        if local_sheet_id is None:
            workbook_scoped.append(NamedRange(name=name, target=target, hidden=hidden))
            continue
        try:
            index = int(local_sheet_id)
        except ValueError:
            unresolved_scopes += 1
            continue
        if not 0 <= index < len(sheet_order) or not sheet_order[index]:
            unresolved_scopes += 1
            continue
        sheet_scoped.append(
            NamedRange(
                name=name,
                target=target,
                sheet=sheet_order[index],
                hidden=hidden,
            )
        )

    details: list[str] = []
    if unresolved_scopes:
        details.append(f"{unresolved_scopes} defined names declare an unreadable scope")
    if macro_skipped:
        details.append(f"{macro_skipped} macro-bound names are not compared as ranges")
    return DefinedNameScan(
        workbook_scoped=tuple(workbook_scoped),
        sheet_scoped=tuple(sheet_scoped),
        available=True,
        builtin_skipped=builtin_skipped,
        macro_skipped=macro_skipped,
        unresolved_scopes=unresolved_scopes,
        detail="; ".join(details),
    )
