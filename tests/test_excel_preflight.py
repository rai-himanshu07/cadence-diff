"""Standalone current-workbook preflight behavior."""

from pathlib import Path

from qc_tool.config.profile import CadenceBand, SheetProfile, default_profile
from qc_tool.coverage import CoverageState, QCRunMode
from qc_tool.engine import run_qc
from qc_tool.excel.preflight import preflight_workbook
from qc_tool.findings import FindingClass
from qc_tool.history.store import sha256_file
from qc_tool.io.model import (
    CellRecord,
    ChartDescriptor,
    ChartSeries,
    NamedRange,
    PivotDescriptor,
    SheetSnapshot,
    WorkbookSnapshot,
)


def test_current_excel_preflight_runs_without_baseline(fixture_dir: Path) -> None:
    source = fixture_dir / "current.xlsx"
    before = sha256_file(source)

    result = run_qc(current_excel=source, mode=QCRunMode.CURRENT_FILE_PREFLIGHT)

    assert result.mode is QCRunMode.CURRENT_FILE_PREFLIGHT
    assert result.files == {"current_excel": "current.xlsx"}
    classes = {finding.finding_class for finding in result.findings}
    assert FindingClass.FORMULA_ERROR in classes
    assert FindingClass.FORMULA_MISSING in classes
    assert FindingClass.FORMULA_CACHE_MISSING in classes
    assert FindingClass.HIDDEN_CONTENT in classes
    formula_gap = next(
        finding
        for finding in result.findings
        if finding.finding_class is FindingClass.FORMULA_MISSING
    )
    assert formula_gap.current_excerpt is not None
    comparison = next(
        item for item in result.coverage if item.check_id == "excel-cycle-comparison"
    )
    assert comparison.state is CoverageState.UNAVAILABLE
    assert sha256_file(source) == before


def test_preflight_detects_intrinsic_structure_and_period_failures() -> None:
    sheet = SheetSnapshot(
        name="Data",
        visibility="visible",
        max_row=3,
        max_column=5,
        cells={
            (1, 1): CellRecord(1, 1, "Metric"),
            (1, 2): CellRecord(1, 2, "Jan-26"),
            (1, 3): CellRecord(1, 3, "Mar-26"),
            (1, 4): CellRecord(1, 4, "Feb-26"),
            (1, 5): CellRecord(1, 5, "Feb-26"),
            (2, 1): CellRecord(2, 1, "Revenue"),
            (2, 2): CellRecord(2, 2, 10),
            (2, 3): CellRecord(2, 3, 20),
            (2, 4): CellRecord(2, 4, 15),
            (2, 5): CellRecord(2, 5, 15),
        },
    )
    workbook = WorkbookSnapshot(
        source_name="current.xlsx",
        file_format="xlsx",
        formulas_available=True,
        styles_available=True,
        sheets=[sheet],
        named_ranges=[NamedRange("Broken", "Missing!$A$1")],
        charts=[
            ChartDescriptor(
                sheet="Data",
                title="Trend",
                chart_type="LineChart",
                series=[
                    ChartSeries(
                        index=0,
                        values_ref="Data!$B$2:$E$2",
                        categories_ref="Data!$B$1:$C$1",
                    )
                ],
            )
        ],
        pivots=[PivotDescriptor("Pivot", "A1:B2", "Missing", "A1:B2")],
        calculation_mode="manual",
        external_links=["file:///old/source.xlsx"],
    )

    result = preflight_workbook(workbook, default_profile())
    classes = {finding.finding_class for finding in result.findings}

    assert {
        FindingClass.PERIOD_DUPLICATE,
        FindingClass.PERIOD_OUT_OF_ORDER,
        FindingClass.PERIOD_GAP,
        FindingClass.CALCULATION_MODE,
        FindingClass.EXTERNAL_LINK,
        FindingClass.NAMED_RANGE_INVALID,
        FindingClass.CHART_LENGTH_MISMATCH,
        FindingClass.PIVOT_SOURCE_INVALID,
    } <= classes


def test_preflight_validates_mixed_cadence_bands_independently() -> None:
    labels = ["Apr-26", "May-26", "Jun-26", "Q4 25", "Q1 26", "Q2 26"]
    cells = {(1, 1): CellRecord(1, 1, "Metric")}
    for column, label in enumerate(labels, start=2):
        cells[(1, column)] = CellRecord(1, column, label)
        cells[(2, column)] = CellRecord(2, column, column * 10)
    cells[(2, 1)] = CellRecord(2, 1, "Value")
    workbook = WorkbookSnapshot(
        source_name="mixed-cadence.xlsx",
        file_format="xlsx",
        formulas_available=True,
        styles_available=True,
        formula_presence_available=True,
        sheets=[
            SheetSnapshot(
                name="Data",
                visibility="visible",
                max_row=2,
                max_column=len(labels) + 1,
                cells=cells,
            )
        ],
    )

    result = preflight_workbook(workbook, default_profile())
    period_findings = {
        finding.finding_class
        for finding in result.findings
        if finding.finding_class
        in {FindingClass.PERIOD_OUT_OF_ORDER, FindingClass.PERIOD_GAP}
    }

    assert period_findings == set()


def test_preflight_duplicate_labels_are_scoped_to_configured_bands() -> None:
    labels = ["Jan-26", "Feb-26", "Jan-26", "Feb-26"]
    cells = {(1, 1): CellRecord(1, 1, "Metric"), (2, 1): CellRecord(2, 1, "Value")}
    for column, label in enumerate(labels, start=2):
        cells[(1, column)] = CellRecord(1, column, label)
        cells[(2, column)] = CellRecord(2, column, column)
    workbook = WorkbookSnapshot(
        source_name="banded.xlsx",
        file_format="xlsx",
        formulas_available=True,
        styles_available=True,
        formula_presence_available=True,
        sheets=[SheetSnapshot("Data", "visible", 2, 5, cells)],
    )
    profile = default_profile()
    profile.excel.sheets["Data"] = SheetProfile(
        cadence_bands=[
            CadenceBand.model_validate({"range": "B1:C1", "kind": "month"}),
            CadenceBand.model_validate({"range": "D1:E1", "kind": "month"}),
        ]
    )

    result = preflight_workbook(workbook, profile)

    assert not any(
        finding.finding_class is FindingClass.PERIOD_DUPLICATE
        for finding in result.findings
    )
