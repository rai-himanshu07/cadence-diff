"""Streaming worksheet metadata extraction from OOXML package parts."""

from __future__ import annotations

import io
import posixpath
import zipfile
from dataclasses import dataclass, field
from typing import Any, cast
from xml.etree import ElementTree

from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

from qc_tool.io.model import TableDescriptor
from qc_tool.io.ooxml_interaction import (
    InteractionExtraction,
    extract_conditional_formatting_element,
    extract_data_validation_element,
)
from qc_tool.progress import CancellationToken, check_cancelled


class OOXMLMetadataError(ValueError):
    """The workbook package metadata cannot be interpreted without guessing."""


@dataclass(slots=True)
class WorksheetMetadata:
    name: str
    visibility: str
    part: str
    max_row: int
    max_column: int
    cell_count: int
    xml_bytes: int
    declared_max_row: int | None = None
    declared_max_column: int | None = None
    hidden_rows: frozenset[int] = frozenset()
    hidden_columns: frozenset[int] = frozenset()
    tables: list[TableDescriptor] = field(default_factory=list)
    interactions: InteractionExtraction = field(default_factory=InteractionExtraction)


@dataclass(slots=True)
class WorkbookMetadata:
    sheets: list[WorksheetMetadata]
    shared_string_bytes: int
    styles_bytes: int
    style_count: int


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _relationship_id(element: ElementTree.Element) -> str | None:
    return next(
        (value for attribute, value in element.attrib.items() if _local_name(attribute) == "id"),
        None,
    )


def _relationships_part(source_part: str) -> str:
    return posixpath.join(
        posixpath.dirname(source_part),
        "_rels",
        f"{posixpath.basename(source_part)}.rels",
    )


def _relationship_targets(
    archive: zipfile.ZipFile,
    source_part: str,
) -> dict[str, str]:
    relationship_part = _relationships_part(source_part)
    if relationship_part not in archive.namelist():
        return {}
    try:
        root = ElementTree.fromstring(archive.read(relationship_part))
    except ElementTree.ParseError as exc:
        raise OOXMLMetadataError(f"malformed relationship part {relationship_part}: {exc}") from exc
    targets: dict[str, str] = {}
    for relationship in root:
        if _local_name(relationship.tag) != "Relationship":
            continue
        relationship_id = relationship.get("Id")
        target = relationship.get("Target")
        if not relationship_id or not target or relationship.get("TargetMode") == "External":
            continue
        if target.startswith("/"):
            resolved = target.lstrip("/")
        else:
            resolved = posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))
        targets[relationship_id] = resolved
    return targets


def _xml_bool(value: str | None) -> bool | None:
    if value is None:
        return None
    return value.casefold() in {"1", "true", "on"}


def _parse_dimension(reference: str | None) -> tuple[int | None, int | None]:
    if not reference:
        return None, None
    try:
        _min_column, _min_row, max_column, max_row = range_boundaries(reference)
    except ValueError:
        return None, None
    return max_row, max_column


def _parse_differential_styles(
    archive: zipfile.ZipFile,
    cancellation_token: CancellationToken | None = None,
) -> tuple[list[DifferentialStyle], int]:
    part = "xl/styles.xml"
    if part not in archive.namelist():
        return [], 0
    styles: list[DifferentialStyle] = []
    style_count = 0
    element_stack: list[str] = []
    dxf_depth = 0
    try:
        with archive.open(part) as stream:
            for event, element in ElementTree.iterparse(stream, events=("start", "end")):
                local_name = _local_name(element.tag)
                if event == "start":
                    element_stack.append(local_name)
                    if (
                        local_name == "xf"
                        and len(element_stack) >= 2
                        and element_stack[-2] == "cellXfs"
                    ):
                        style_count += 1
                        if style_count % 1_000 == 0:
                            check_cancelled(cancellation_token)
                    if dxf_depth:
                        dxf_depth += 1
                    elif local_name == "dxf":
                        dxf_depth = 1
                    continue
                if dxf_depth:
                    dxf_depth -= 1
                    if dxf_depth == 0:
                        style = cast(
                            DifferentialStyle | None,
                            DifferentialStyle.from_tree(cast(Any, element)),
                        )
                        if style is None:
                            raise OOXMLMetadataError(
                                "differential style element could not be parsed"
                            )
                        styles.append(style)
                        element.clear()
                else:
                    element.clear()
                element_stack.pop()
    except ElementTree.ParseError as exc:
        raise OOXMLMetadataError(f"malformed styles part {part}: {exc}") from exc
    return styles, style_count


def _parse_table(
    archive: zipfile.ZipFile,
    *,
    part: str,
    sheet: str,
) -> TableDescriptor:
    if part not in archive.namelist():
        raise OOXMLMetadataError(f"{sheet}: linked table part is missing: {part}")
    try:
        root = ElementTree.fromstring(archive.read(part))
    except ElementTree.ParseError as exc:
        raise OOXMLMetadataError(f"{sheet}: malformed table part {part}: {exc}") from exc
    name = root.get("name")
    display_name = root.get("displayName")
    cell_range = root.get("ref")
    if not name or not display_name or not cell_range:
        raise OOXMLMetadataError(f"{sheet}: incomplete table metadata in {part}")
    header_row_count = int(root.get("headerRowCount", "1"))
    totals_count = root.get("totalsRowCount")
    totals_row_count = (
        int(totals_count)
        if totals_count is not None
        else int(bool(_xml_bool(root.get("totalsRowShown"))))
    )
    source_id = root.get("id")
    columns = [
        str(element.get("name"))
        for element in root.iter()
        if _local_name(element.tag) == "tableColumn" and element.get("name") is not None
    ]
    return TableDescriptor(
        sheet=sheet,
        name=name,
        display_name=display_name,
        cell_range=cell_range,
        columns=columns,
        header_row_count=header_row_count,
        totals_row_count=totals_row_count,
        source_id=int(source_id) if source_id is not None else None,
    )


def _merge_interactions(
    target: InteractionExtraction,
    source: InteractionExtraction,
) -> None:
    target.data_validations.extend(source.data_validations)
    target.conditional_formats.extend(source.conditional_formats)
    target.rules_supported = target.rules_supported and source.rules_supported
    target.rule_details.extend(source.rule_details)
    target.styles_supported = target.styles_supported and source.styles_supported
    target.style_details.extend(source.style_details)


def _parse_worksheet(
    archive: zipfile.ZipFile,
    *,
    name: str,
    visibility: str,
    part: str,
    differential_styles: list[DifferentialStyle],
    cancellation_token: CancellationToken | None = None,
) -> WorksheetMetadata:
    check_cancelled(cancellation_token)
    if part not in archive.namelist():
        raise OOXMLMetadataError(f"{name}: linked worksheet part is missing: {part}")
    max_row = max_column = cell_count = 0
    declared_max_row = declared_max_column = None
    hidden_rows: set[int] = set()
    hidden_columns: set[int] = set()
    table_relationship_ids: list[str] = []
    interactions = InteractionExtraction()
    validation_index = conditional_index = 0
    retained_depth = 0
    try:
        with archive.open(part) as stream:
            for event, element in ElementTree.iterparse(stream, events=("start", "end")):
                local_name = _local_name(element.tag)
                if event == "start":
                    if retained_depth:
                        retained_depth += 1
                    elif local_name in {"dataValidation", "conditionalFormatting"}:
                        retained_depth = 1
                    continue

                if retained_depth:
                    retained_depth -= 1
                    if retained_depth:
                        continue
                    if local_name == "dataValidation":
                        interactions.data_validations.append(
                            extract_data_validation_element(
                                element,
                                sheet=name,
                                source_index=validation_index,
                            )
                        )
                        validation_index += 1
                    elif local_name == "conditionalFormatting":
                        group = extract_conditional_formatting_element(
                            element,
                            sheet=name,
                            source_index=conditional_index,
                            differential_styles=differential_styles,
                        )
                        conditional_index += len(group.conditional_formats)
                        _merge_interactions(interactions, group)
                    element.clear()
                    continue

                if local_name == "dimension":
                    declared_max_row, declared_max_column = _parse_dimension(element.get("ref"))
                elif local_name == "c":
                    reference = element.get("r")
                    if not reference:
                        raise OOXMLMetadataError(
                            f"{name}: worksheet cell is missing its coordinate"
                        )
                    try:
                        row, column = coordinate_to_tuple(reference)
                    except ValueError as exc:
                        raise OOXMLMetadataError(
                            f"{name}: invalid worksheet cell coordinate {reference!r}"
                        ) from exc
                    max_row = max(max_row, row)
                    max_column = max(max_column, column)
                    cell_count += 1
                    if cell_count % 10_000 == 0:
                        check_cancelled(cancellation_token)
                elif local_name == "row" and _xml_bool(element.get("hidden")):
                    row_index = element.get("r")
                    if row_index is not None:
                        try:
                            hidden_rows.add(int(row_index))
                        except ValueError as exc:
                            raise OOXMLMetadataError(
                                f"{name}: invalid hidden row index {row_index!r}"
                            ) from exc
                elif local_name == "col" and _xml_bool(element.get("hidden")):
                    raw_minimum = element.get("min", "0")
                    raw_maximum = element.get("max", "0")
                    try:
                        minimum = int(raw_minimum)
                        maximum = int(raw_maximum)
                    except ValueError as exc:
                        raise OOXMLMetadataError(
                            f"{name}: invalid hidden column span "
                            f"{raw_minimum!r}:{raw_maximum!r}"
                        ) from exc
                    if minimum < 1 or maximum < minimum:
                        raise OOXMLMetadataError(
                            f"{name}: invalid hidden column span {minimum}:{maximum}"
                        )
                    hidden_columns.update(range(minimum, maximum + 1))
                elif local_name == "tablePart":
                    relationship_id = _relationship_id(element)
                    if relationship_id:
                        table_relationship_ids.append(relationship_id)
                element.clear()
    except ElementTree.ParseError as exc:
        raise OOXMLMetadataError(f"{name}: malformed worksheet part {part}: {exc}") from exc

    relationships = _relationship_targets(archive, part)
    tables = [
        _parse_table(
            archive,
            part=relationships[relationship_id],
            sheet=name,
        )
        for relationship_id in table_relationship_ids
        if relationship_id in relationships
    ]
    return WorksheetMetadata(
        name=name,
        visibility=visibility,
        part=part,
        max_row=max_row or 1,
        max_column=max_column or 1,
        cell_count=cell_count,
        xml_bytes=archive.getinfo(part).file_size,
        declared_max_row=declared_max_row,
        declared_max_column=declared_max_column,
        hidden_rows=frozenset(hidden_rows),
        hidden_columns=frozenset(hidden_columns),
        tables=tables,
        interactions=interactions,
    )


def parse_ooxml_worksheet_metadata(
    data: bytes,
    *,
    cancellation_token: CancellationToken | None = None,
) -> WorkbookMetadata:
    """Parse workbook and worksheet metadata without constructing cell object graphs."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        workbook_part = "xl/workbook.xml"
        if workbook_part not in archive.namelist():
            raise OOXMLMetadataError("OOXML package is missing xl/workbook.xml")
        try:
            workbook = ElementTree.fromstring(archive.read(workbook_part))
        except ElementTree.ParseError as exc:
            raise OOXMLMetadataError(f"malformed workbook part: {exc}") from exc
        relationships = _relationship_targets(archive, workbook_part)
        differential_styles, style_count = _parse_differential_styles(
            archive,
            cancellation_token,
        )
        sheets: list[WorksheetMetadata] = []
        for element in workbook.iter():
            if _local_name(element.tag) != "sheet":
                continue
            check_cancelled(cancellation_token)
            name = element.get("name")
            relationship_id = _relationship_id(element)
            if not name or not relationship_id or relationship_id not in relationships:
                raise OOXMLMetadataError("workbook contains an unresolved worksheet")
            sheets.append(
                _parse_worksheet(
                    archive,
                    name=name,
                    visibility=element.get("state", "visible"),
                    part=relationships[relationship_id],
                    differential_styles=differential_styles,
                    cancellation_token=cancellation_token,
                )
            )
        names = set(archive.namelist())
        return WorkbookMetadata(
            sheets=sheets,
            shared_string_bytes=(
                archive.getinfo("xl/sharedStrings.xml").file_size
                if "xl/sharedStrings.xml" in names
                else 0
            ),
            styles_bytes=(
                archive.getinfo("xl/styles.xml").file_size if "xl/styles.xml" in names else 0
            ),
            style_count=style_count,
        )
