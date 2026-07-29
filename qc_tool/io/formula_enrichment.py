"""Fail-closed contracts for enriching XLSB snapshots with formula text."""

from __future__ import annotations

from dataclasses import dataclass

from qc_tool.io.model import CellRecord, WorkbookSnapshot
from qc_tool.io.xlsb_formula import XlsbFormulaScan

FormulaCoordinate = tuple[int, int]
FormulaMap = dict[str, dict[FormulaCoordinate, str]]


class FormulaEnrichmentError(RuntimeError):
    """Formula text cannot be trusted for this workbook."""


@dataclass(frozen=True, slots=True)
class FormulaExtraction:
    """Formula text produced by one external spreadsheet engine."""

    formulas: FormulaMap
    engine: str
    detail: str

    @property
    def coordinates(self) -> frozenset[tuple[str, int, int]]:
        return frozenset(
            (sheet, row, column)
            for sheet, cells in self.formulas.items()
            for row, column in cells
        )


def scan_coordinates(scan: XlsbFormulaScan) -> frozenset[tuple[str, int, int]]:
    """Flatten scan coordinates for exact two-way parity checks."""
    return frozenset(
        (sheet, row, column)
        for sheet, cells in scan.formula_cells.items()
        for row, column in cells
    )


def validate_formula_extraction(
    scan: XlsbFormulaScan, extraction: FormulaExtraction
) -> None:
    """Require valid formula text and exact source/extractor coordinate parity."""
    expected = scan_coordinates(scan)
    actual = extraction.coordinates
    missing = expected - actual
    unexpected = actual - expected
    if missing or unexpected:
        raise FormulaEnrichmentError(
            "formula-coordinate mismatch: "
            f"{len(missing)} missing and {len(unexpected)} unexpected"
        )
    for sheet, cells in extraction.formulas.items():
        for coordinate, formula in cells.items():
            if not isinstance(formula, str) or not formula.startswith("="):
                raise FormulaEnrichmentError(
                    f"{sheet}!{coordinate}: external engine returned invalid formula text"
                )


def merge_formula_extraction(
    snapshot: WorkbookSnapshot,
    scan: XlsbFormulaScan,
    extraction: FormulaExtraction,
) -> None:
    """Merge trusted formula text without replacing original cached values."""
    if snapshot.file_format != "xlsb":
        raise FormulaEnrichmentError("formula enrichment is only valid for XLSB snapshots")
    validate_formula_extraction(scan, extraction)

    for sheet_name, formulas in extraction.formulas.items():
        try:
            sheet = snapshot.sheet(sheet_name)
        except KeyError as exc:
            raise FormulaEnrichmentError(
                f"external engine returned unknown sheet {sheet_name!r}"
            ) from exc
        for (row, column), formula in formulas.items():
            cell = sheet.cells.get((row, column))
            if cell is None:
                cell = CellRecord(row=row, column=column, value=None, is_formula=True)
                sheet.cells[(row, column)] = cell
                sheet.max_row = max(sheet.max_row, row)
                sheet.max_column = max(sheet.max_column, column)
            cell.formula = formula
            cell.is_formula = True

    snapshot.formula_presence_available = True
    snapshot.formulas_available = True
    snapshot.formula_source = extraction.engine
    snapshot.formula_detail = extraction.detail
