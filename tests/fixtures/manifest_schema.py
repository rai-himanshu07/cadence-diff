"""Ground-truth manifest schema for generated QC fixtures.

The generator seeds known defects into the "current" artifacts and records
every seed here. Engine tests assert their findings against this manifest.
"""

from enum import StrEnum

from pydantic import BaseModel, Field


class Artifact(StrEnum):
    EXCEL = "excel"
    XLSB = "xlsb"
    PPT = "ppt"
    CROSSCHECK = "crosscheck"


class DefectClass(StrEnum):
    VALUE_CHANGED = "value_changed"
    FORMULA_ERROR = "formula_error"
    FORMULA_HARDCODED = "formula_hardcoded"
    FORMULA_NOT_EXTENDED = "formula_not_extended"
    FORMULA_LOGIC_CHANGED = "formula_logic_changed"
    FORMULA_INCONSISTENT = "formula_inconsistent"
    ROW_DELETED = "row_deleted"
    COLUMN_DELETED = "column_deleted"
    SHEET_ADDED = "sheet_added"
    SHEET_REMOVED = "sheet_removed"
    HIDDEN_CHANGED = "hidden_changed"
    NAMED_RANGE_CHANGED = "named_range_changed"
    NUMBER_FORMAT_CHANGED = "number_format_changed"
    STYLE_CHANGED = "style_changed"
    CHART_SERIES_CHANGED = "chart_series_changed"
    PIVOT_SOURCE_CHANGED = "pivot_source_changed"
    SLIDE_ADDED = "slide_added"
    SLIDE_REMOVED = "slide_removed"
    SLIDE_TEXT_CHANGED = "slide_text_changed"
    TABLE_VALUE_CHANGED = "table_value_changed"
    CHART_VALUE_CHANGED = "chart_value_changed"
    CROSSCHECK_MISMATCH = "crosscheck_mismatch"


class SeededDefect(BaseModel):
    """One deliberately seeded problem the QC engines must detect."""

    defect_id: str
    artifact: Artifact
    classes: list[DefectClass] = Field(min_length=1)
    sheet: str | None = None
    cell: str | None = None
    baseline_cell: str | None = None
    slide_title: str | None = None
    element: str | None = None
    baseline: str | None = None
    current: str | None = None
    impacts: list[str] = Field(default_factory=list)
    note: str = ""


class ExpectedChange(BaseModel):
    """A cadence-growth change that must NOT be reported as an error."""

    change_id: str
    artifact: Artifact
    kind: str
    sheet: str | None = None
    slide_title: str | None = None
    detail: str


class CrosscheckEntry(BaseModel):
    """Ground truth linking a PPT figure to its workbook source cell."""

    slide_title: str
    figure_label: str
    figure_text: str
    source_cell: str
    matches: bool
    defect_id: str | None = None


class FixtureManifest(BaseModel):
    schema_version: int = 1
    seed: int
    generated_at: str
    password: str
    files: dict[str, str]
    defects: list[SeededDefect]
    expected: list[ExpectedChange]
    crosscheck_map: list[CrosscheckEntry]

    def defect(self, defect_id: str) -> SeededDefect:
        for d in self.defects:
            if d.defect_id == defect_id:
                return d
        raise KeyError(f"no seeded defect with id {defect_id!r}")
