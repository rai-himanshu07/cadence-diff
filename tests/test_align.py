"""Alignment engine tests against the fixture pair's known ground truth."""

from pathlib import Path

import pytest
from openpyxl import Workbook

from qc_tool.config.profile import DeliverableProfile
from qc_tool.coverage import CoverageState
from qc_tool.engine import run_qc
from qc_tool.excel.align import (
    AxisEntry,
    RegionAlignment,
    WorkbookAlignment,
    _align_axis,
    _pair_regions,
    align_workbooks,
)
from qc_tool.excel.periods import parse_period
from qc_tool.excel.regions import detect_regions
from qc_tool.findings import FindingClass
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.io.model import CellRecord, SheetSnapshot, WorkbookSnapshot


@pytest.fixture(scope="module")
def baseline(fixture_dir: Path) -> WorkbookSnapshot:
    return load_workbook_snapshot(fixture_dir / "baseline.xlsx")


@pytest.fixture(scope="module")
def current(fixture_dir: Path) -> WorkbookSnapshot:
    return load_workbook_snapshot(fixture_dir / "current.xlsx")


@pytest.fixture(scope="module")
def alignment(baseline: WorkbookSnapshot, current: WorkbookSnapshot) -> WorkbookAlignment:
    return align_workbooks(baseline, current)


def _single_region(alignment: WorkbookAlignment, sheet: str) -> RegionAlignment:
    regions = alignment.regions[sheet]
    assert len(regions) == 1, f"expected one region pair on {sheet}"
    return regions[0]


def test_sheet_pairing(alignment: WorkbookAlignment) -> None:
    assert alignment.added_sheets == ["New_Analysis"]  # E15
    assert alignment.removed_sheets == ["Old_Sheet"]  # E14
    assert set(alignment.common_sheets) == {
        "Long_Monthly",
        "Wide_Weekly",
        "Dashboard",
        "Summary",
        "Params",
    }


def test_long_monthly_growth_isolated(alignment: WorkbookAlignment) -> None:
    region = _single_region(alignment, "Long_Monthly")
    assert region.rows.method == "keys"
    # 20 historical rows pair exactly onto themselves; zero false movement.
    assert region.rows.pairs == [(row, row) for row in range(2, 22)]
    assert region.rows.growth == [22, 23, 24, 25]  # EX01: Jun-26 block
    assert region.rows.inserted == []
    assert region.rows.deleted == []
    # Header columns are stable.
    assert region.columns.pairs == [(col, col) for col in range(1, 6)]
    assert region.columns.growth == []


def test_wide_weekly_deletion_and_growth(alignment: WorkbookAlignment) -> None:
    region = _single_region(alignment, "Wide_Weekly")
    columns = region.columns
    assert columns.method == "keys"
    assert columns.deleted == [4]  # E07: W03 lived in column D
    assert columns.growth == [21]  # EX02: W21 appended in column U
    assert columns.inserted == []
    # W04 onward shifted one column left; key alignment must absorb the shift.
    pair_map = dict(columns.pairs)
    assert pair_map[1] == 1  # label column
    assert pair_map[2] == 2 and pair_map[3] == 3  # W01, W02
    assert pair_map[5] == 4  # W04: E -> D
    assert pair_map[21] == 20  # W20: U -> T
    assert region.rows.pairs == [(row, row) for row in range(1, 5)]


def test_yearless_week_rollover_is_growth() -> None:
    baseline = [
        AxisEntry(index=1, key=("W51", 0), periods=(parse_period("W51"),)),
        AxisEntry(index=2, key=("W52", 0), periods=(parse_period("W52"),)),
    ]
    current = [
        *baseline,
        AxisEntry(index=3, key=("W01", 0), periods=(parse_period("W01"),)),
    ]

    aligned = _align_axis(baseline, current)

    assert aligned.growth == [3]
    assert aligned.inserted == []


def test_key_alignment_marks_low_confidence_positional_fallback() -> None:
    baseline = [
        AxisEntry(index=index, key=(label,), periods=(None,))
        for index, label in enumerate(("A", "B", "C", "D"), start=1)
    ]
    current = [
        AxisEntry(index=index, key=(label,), periods=(None,))
        for index, label in enumerate(("W", "X", "Y", "Z"), start=1)
    ]

    aligned = _align_axis(baseline, current)

    assert aligned.method == "positional"
    assert aligned.low_confidence_fallback


def test_low_confidence_alignment_degrades_and_skips_cell_comparison(
    tmp_path: Path,
) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"

    def write(path: Path, labels: list[str], offset: int) -> None:
        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "Data"
        sheet.append(["Key", "Value"])
        for index, label in enumerate(labels, start=1):
            sheet.append([label, index + offset])
        workbook.save(path)

    write(baseline, ["A", "B", "C", "D"], 0)
    write(current, ["W", "X", "Y", "Z"], 100)

    result = run_qc(baseline_excel=baseline, current_excel=current)

    value_coverage = next(
        item for item in result.coverage if item.check_id == "excel-values"
    )
    formula_coverage = next(
        item for item in result.coverage if item.check_id == "excel-formulas"
    )
    assert value_coverage.state is CoverageState.DEGRADED
    assert formula_coverage.state is CoverageState.DEGRADED
    assert "Low-confidence" in value_coverage.detail
    assert sum(
        finding.finding_class is FindingClass.ALIGNMENT_LOW_CONFIDENCE
        for finding in result.findings
    ) == 1
    assert not any(
        finding.finding_class is FindingClass.VALUE_CHANGED
        for finding in result.findings
    )


def test_dashboard_blocks_align(alignment: WorkbookAlignment) -> None:
    regions = alignment.regions["Dashboard"]
    assert len(regions) == 2
    kpi, headcount = regions

    kpi_rows = dict(kpi.rows.pairs)
    assert kpi_rows == {row: row for row in range(1, 6)}
    assert kpi.rows.deleted == [] and kpi.rows.inserted == []

    assert headcount.columns.growth == [11]  # EX05: Jun-26 headcount column (K)
    assert headcount.columns.deleted == []
    assert dict(headcount.rows.pairs) == {row: row for row in range(2, 5)}


def test_region_pairing_survives_inserted_block() -> None:
    def sheet(name: str, blocks: list[tuple[int, tuple[str, str]]]) -> SheetSnapshot:
        cells: dict[tuple[int, int], CellRecord] = {}
        for start_row, block_labels in blocks:
            for offset, label in enumerate(block_labels):
                row = start_row + offset
                cells[(row, 1)] = CellRecord(row=row, column=1, value=label)
                cells[(row, 2)] = CellRecord(row=row, column=2, value=100 + row)
        return SheetSnapshot(
            name=name,
            visibility="visible",
            max_row=max(row for row, _ in cells),
            max_column=2,
            cells=cells,
        )

    baseline = sheet("Dashboard", [(1, ("Revenue", "Margin")), (5, ("Headcount", "Attrition"))])
    current = sheet(
        "Dashboard",
        [
            (1, ("New metric", "Other")),
            (5, ("Revenue", "Margin")),
            (9, ("Headcount", "Attrition")),
        ],
    )

    pairs, unpaired_base, unpaired_current = _pair_regions(
        baseline,
        current,
        detect_regions(baseline),
        detect_regions(current),
    )

    assert [(base.cell_range, curr.cell_range) for base, curr in pairs] == [
        ("A1:B2", "A5:B6"),
        ("A5:B6", "A9:B10"),
    ]
    assert unpaired_base == []
    assert [region.cell_range for region in unpaired_current] == ["A1:B2"]


def test_summary_block_alignment(alignment: WorkbookAlignment) -> None:
    region = _single_region(alignment, "Summary")
    assert dict(region.rows.pairs) == {row: row for row in range(2, 7)}
    # C/D columns exist only in current (seeded error cells E05/E13): they are
    # unexpected insertions, not cadence growth.
    assert region.columns.inserted == [3, 4]
    assert region.columns.growth == []


def test_profile_sheet_ignore(baseline: WorkbookSnapshot, current: WorkbookSnapshot) -> None:
    profile = DeliverableProfile.model_validate(
        {"name": "test", "excel": {"ignore_sheets": ["Params", "Old_Sheet", "New_Analysis"]}}
    )
    aligned = align_workbooks(baseline, current, profile)
    assert aligned.added_sheets == []
    assert aligned.removed_sheets == []
    assert "Params" not in aligned.common_sheets


def test_cell_pairs_identity_on_unchanged_history(alignment: WorkbookAlignment) -> None:
    region = _single_region(alignment, "Long_Monthly")
    pairs = list(region.cell_pairs())
    assert ((7, 3), (7, 3)) in pairs  # E01's cell maps onto itself
    assert len(pairs) == 20 * 5
    assert all(base == curr for base, curr in pairs)
