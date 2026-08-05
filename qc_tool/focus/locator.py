"""Structured locator grammar for private focus targets.

Only these parsers may produce a focus address. Message prose, ``element``
labels, ``baseline_value``, and ``current_value`` are never parsed: a target
that cannot be proved by a structured locator simply does not exist.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from openpyxl.utils import get_column_letter

#: Worksheet grid limits of the modern Excel file formats.
MAX_COLUMN = 16_384
MAX_ROW = 1_048_576

#: Largest selection a focus action may request, i.e. one whole column.
MAX_FOCUS_CELLS = MAX_ROW

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$")
_ROW_ONLY_RE = re.compile(r"^\$?([1-9][0-9]*)$")
_COLUMN_ONLY_RE = re.compile(r"^\$?([A-Za-z]{1,3})$")
_ROW_SPAN_RE = re.compile(r"^row ([1-9][0-9]*)$")
_COLUMN_SPAN_RE = re.compile(r"^column ([A-Za-z]{1,3})$")


@dataclass(frozen=True, slots=True)
class CellAddress:
    row: int
    column: int


def column_index(letters: str) -> int | None:
    """1-based column index for ``A``..``XFD``, or ``None`` when out of grid."""
    index = 0
    for character in letters.upper():
        if not "A" <= character <= "Z":
            return None
        index = index * 26 + (ord(character) - ord("A") + 1)
    if not 1 <= index <= MAX_COLUMN:
        return None
    return index


def parse_cell(text: str | None) -> CellAddress | None:
    """One A1 cell reference, absolute markers allowed."""
    if not text:
        return None
    match = _CELL_RE.match(text.strip())
    if match is None:
        return None
    column = column_index(match.group(1))
    row = int(match.group(2))
    if column is None or row > MAX_ROW:
        return None
    return CellAddress(row=row, column=column)


def _endpoint(text: str) -> tuple[int | None, int | None] | None:
    """``(row, column)`` for one range endpoint; ``None`` marks a whole span."""
    cell = parse_cell(text)
    if cell is not None:
        return cell.row, cell.column
    row_only = _ROW_ONLY_RE.match(text)
    if row_only is not None:
        row = int(row_only.group(1))
        return (row, None) if row <= MAX_ROW else None
    column_only = _COLUMN_ONLY_RE.match(text)
    if column_only is not None:
        column = column_index(column_only.group(1))
        return (None, column) if column is not None else None
    return None


def _render(row: int | None, column: int | None) -> str:
    letter = get_column_letter(column) if column is not None else ""
    return f"{letter}{row if row is not None else ''}"


def normalize_address(text: str | None) -> str | None:
    """Canonical bounded A1 address: ``B5``, ``A1:D6``, ``5:5``, or ``C:C``.

    Returns ``None`` for anything that is not a structurally valid, bounded
    reference. Endpoints are ordered so a reversed range is still deterministic.
    """
    if not text:
        return None
    candidate = text.strip()
    if not candidate or candidate.count(":") > 1:
        return None
    if ":" not in candidate:
        cell = parse_cell(candidate)
        return None if cell is None else _render(cell.row, cell.column)
    left, right = candidate.split(":")
    start = _endpoint(left.strip())
    end = _endpoint(right.strip())
    if start is None or end is None:
        return None
    start_row, start_column = start
    end_row, end_column = end
    if (start_row is None) != (end_row is None):
        return None
    if (start_column is None) != (end_column is None):
        return None
    if start_row is not None and end_row is not None:
        start_row, end_row = min(start_row, end_row), max(start_row, end_row)
    if start_column is not None and end_column is not None:
        start_column, end_column = (
            min(start_column, end_column),
            max(start_column, end_column),
        )
    rows = (end_row - start_row + 1) if start_row is not None and end_row is not None else MAX_ROW
    columns = (
        (end_column - start_column + 1)
        if start_column is not None and end_column is not None
        else MAX_COLUMN
    )
    if rows * columns > MAX_FOCUS_CELLS:
        return None
    return f"{_render(start_row, start_column)}:{_render(end_row, end_column)}"


def parse_axis_span(text: str | None) -> str | None:
    """The row/column axis wording emitted by region diffs, as an A1 span."""
    if not text:
        return None
    candidate = text.strip()
    row_span = _ROW_SPAN_RE.match(candidate)
    if row_span is not None:
        return normalize_address(f"{row_span.group(1)}:{row_span.group(1)}")
    column_span = _COLUMN_SPAN_RE.match(candidate)
    if column_span is not None:
        letter = column_span.group(1).upper()
        return normalize_address(f"{letter}:{letter}")
    return None


def parse_excel_location(text: str | None) -> str | None:
    """A structured Excel locator from a producer-supplied location field."""
    return normalize_address(text) or parse_axis_span(text)


def parse_qualified_location(text: str | None) -> tuple[str, str] | None:
    """``Sheet!A1`` used by crosscheck findings, split on the final separator.

    A worksheet name may itself contain ``!``, while an address never can, so
    the split is anchored on the last separator.
    """
    if not text:
        return None
    candidate = text.strip()
    if "!" not in candidate:
        return None
    sheet, _, address = candidate.rpartition("!")
    sheet = sheet.strip().strip("'")
    normalized = normalize_address(address)
    if not sheet or normalized is None:
        return None
    return sheet, normalized
