"""Synthetic twin of the whole-column-aggregate workbook shape (in memory).

Reproduces the measured 2026-08-13 pathology drivers: a tall data sheet whose
columns are referenced with `$X:$X` whole-column aggregates from many formula
cells (identical rectangles re-walked per cell), mixed with large-but-concrete
literal ranges and single-cell references.
"""

from __future__ import annotations

from openpyxl.utils import get_column_letter

from qc_tool.io.model import CellRecord, SheetSnapshot, WorkbookSnapshot

DATA_SHEET = "Data"
CALC_SHEET = "Calc"


def whole_column_twin(
    *,
    data_rows: int,
    formula_cells: int,
    whole_columns: tuple[int, ...] = (1, 2, 3),
    literal_range_rows: int = 0,
) -> WorkbookSnapshot:
    """Build the twin: ``formula_cells`` cells each aggregating the data sheet.

    Every formula references one `$X:$X` whole column (rotating through
    ``whole_columns``) plus one direct cell; when ``literal_range_rows`` > 0 a
    concrete literal range of that height is added, exercising the descriptor
    path alongside the whole-column path.
    """
    data_cells: dict[tuple[int, int], CellRecord] = {}
    for column in whole_columns:
        for row in range(1, data_rows + 1):
            data_cells[(row, column)] = CellRecord(row, column, float(row))
    data = SheetSnapshot(
        DATA_SHEET, "visible", data_rows, max(whole_columns), data_cells
    )

    calc_cells: dict[tuple[int, int], CellRecord] = {}
    per_row = 8
    for index in range(formula_cells):
        row, column = divmod(index, per_row)
        row += 1
        column += 1
        letter = get_column_letter(whole_columns[index % len(whole_columns)])
        parts = [f"SUMIF('{DATA_SHEET}'!${letter}:${letter},$A$1)"]
        parts.append(f"'{DATA_SHEET}'!$A$1")
        if literal_range_rows:
            parts.append(f"SUM('{DATA_SHEET}'!$B$1:$B${literal_range_rows})")
        calc_cells[(row, column)] = CellRecord(
            row, column, 0.0, formula="=" + "+".join(parts)
        )
    rows = (formula_cells + per_row - 1) // per_row
    calc = SheetSnapshot(CALC_SHEET, "visible", max(rows, 1), per_row, calc_cells)

    return WorkbookSnapshot(
        "whole-column-twin.xlsx",
        "xlsx",
        True,
        True,
        formula_presence_available=True,
        formula_source="openpyxl",
        sheets=[data, calc],
    )


def equality_twin() -> WorkbookSnapshot:
    """Small shape for exact oracle parity (force symbolic via max_range_cells)."""
    return whole_column_twin(
        data_rows=120, formula_cells=48, literal_range_rows=40
    )


def pathology_twin() -> WorkbookSnapshot:
    """CI-sized pathology: whole columns exceed the real 50k symbolic threshold."""
    return whole_column_twin(
        data_rows=55_000,
        formula_cells=2_000,
        whole_columns=(1, 2, 3, 4, 5),
        literal_range_rows=30_000,
    )
