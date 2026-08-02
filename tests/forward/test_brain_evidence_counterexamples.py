"""Forward-failing behavior contracts for the Brain evidence plan Step 1."""

from __future__ import annotations

from qc_tool.config.profile import NumericTolerance
from qc_tool.excel.align import WorkbookAlignment, align_regions
from qc_tool.excel.diff_values import _RangeSet, diff_region_values
from qc_tool.excel.formulas import diff_workbook_formulas
from qc_tool.excel.regions import TableRegion
from qc_tool.findings import (
    Finding,
    FindingClass,
    Materiality,
    Severity,
)
from qc_tool.io.model import CellRecord, CellValue, SheetSnapshot, WorkbookSnapshot
from qc_tool.review import build_pattern_groups
from qc_tool.story import StoryKind, build_stories
from qc_tool.triage.rules import assign_severity, triage


def _sheet(values: dict[tuple[int, int], CellValue]) -> SheetSnapshot:
    cells = {
        (row, col): CellRecord(row, col, value)
        for (row, col), value in values.items()
    }
    return SheetSnapshot("Panel", "visible", 2, 2, cells)


def test_catastrophic_numeric_block_refresh_remains_critical() -> None:
    baseline = _sheet(
        {
            (1, 1): "Cycle",
            (1, 2): "Jun-26",
            (2, 1): "Revenue",
            (2, 2): 100.0,
        }
    )
    current = _sheet(
        {
            (1, 1): "Cycle",
            (1, 2): "Jul-26",
            (2, 1): "Revenue",
            (2, 2): 1_000_000_000.0,
        }
    )
    region = TableRegion("Panel", 1, 1, 2, 2, "block", None, 1, "none")
    alignment = align_regions(baseline, current, region, region)

    findings = diff_region_values(
        baseline,
        current,
        alignment,
        NumericTolerance(),
        ignore=_RangeSet([]),
        refresh=_RangeSet([]),
        sheet_profile=None,
    )
    catastrophic = next(
        finding
        for finding in findings
        if finding.finding_class is FindingClass.VALUE_CHANGED
        and finding.location == "B2"
    )

    assert assign_severity(catastrophic) is Severity.CRITICAL


def _workbook(cells: dict[tuple[int, int], CellRecord]) -> WorkbookSnapshot:
    return WorkbookSnapshot(
        source_name="anonymous.xlsx",
        file_format="xlsx",
        formulas_available=True,
        styles_available=True,
        formula_presence_available=True,
        sheets=[
            SheetSnapshot(
                "Data",
                "visible",
                max((row for row, _ in cells), default=1),
                max((col for _, col in cells), default=1),
                cells,
            )
        ],
    )


def _standalone_errors(
    cells: dict[tuple[int, int], CellRecord],
) -> list[Finding]:
    findings = diff_workbook_formulas(
        _workbook({}),
        _workbook(cells),
        WorkbookAlignment(),
        cycle=False,
    )
    return triage(
        [
            finding
            for finding in findings
            if finding.finding_class is FindingClass.FORMULA_ERROR
        ]
    )


def _sparse_error_population(
    column: int, literal: str, *, count: int = 250
) -> dict[tuple[int, int], CellRecord]:
    cells = {
        (row, column): CellRecord(row, column, literal)
        for row in range(2, count + 2)
    }
    cells.update(
        {
            (row, column): CellRecord(row, column, float(row))
            for row in range(300, 5000)
        }
    )
    return cells


def test_sparse_value_only_mass_is_one_critical_group() -> None:
    errors = _standalone_errors(_sparse_error_population(2, "#N/A"))
    groups = [
        group
        for group in build_pattern_groups(errors)
        if group.finding_class is FindingClass.FORMULA_ERROR
    ]

    assert {finding.severity for finding in errors} == {Severity.CRITICAL}
    assert len(groups) == 1


def test_distinct_error_columns_and_literals_form_distinct_groups() -> None:
    cells = _sparse_error_population(2, "#N/A", count=200)
    cells.update(_sparse_error_population(3, "#VALUE!", count=200))
    errors = _standalone_errors(cells)
    groups = [
        group
        for group in build_pattern_groups(errors)
        if group.finding_class is FindingClass.FORMULA_ERROR
    ]

    assert len(groups) == 2


def test_unrelated_same_sheet_drivers_form_separate_stories() -> None:
    findings = triage(
        [
            Finding(
                artifact="excel",
                finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
                sheet="Setup",
                element="RevenueTable",
                current_value="RevenueAmount",
                message="revenue table columns changed",
            ),
            Finding(
                artifact="excel",
                finding_class=FindingClass.DATA_VALIDATION_CHANGED,
                sheet="Setup",
                element="RegionSelector",
                current_value="RegionCode",
                message="region selector criteria changed",
            ),
        ]
    )
    stories = [
        story
        for story in build_stories(findings)
        if story.kind is StoryKind.STRUCTURE_DRIVER
    ]

    assert len(stories) == 2
    assert sorted(story.member_count for story in stories) == [1, 1]


def test_accepted_difference_is_not_a_data_refresh_story() -> None:
    accepted = Finding(
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        materiality=Materiality.WITHIN_TOLERANCE,
        sheet="Data",
        location="B2",
        baseline_location="B2",
        baseline_value="100",
        current_value="101",
        message="accepted numeric difference",
    )
    findings = triage([accepted])
    stories = build_stories(findings)
    refresh_ids = {
        finding_id
        for story in stories
        if story.kind is StoryKind.DATA_REFRESH
        for finding_id in story.finding_ids
    }

    assert findings[0].finding_id not in refresh_ids
