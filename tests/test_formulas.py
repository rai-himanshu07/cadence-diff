"""Formula QC tests (acceptance criterion 3) against the fixture manifest."""

from pathlib import Path

import pytest

from qc_tool.excel.align import WorkbookAlignment, align_workbooks
from qc_tool.excel.formulas import (
    _differs_only_by_extension,
    diff_workbook_formulas,
    to_r1c1,
)
from qc_tool.findings import Finding, FindingClass
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.io.model import WorkbookSnapshot
from tests.fixtures.manifest_schema import FixtureManifest


@pytest.mark.parametrize(
    ("formula", "row", "col", "expected"),
    [
        ("=C10-D10", 10, 5, "=RC[-2]-RC[-1]"),
        ("=C11-D11", 11, 5, "=RC[-2]-RC[-1]"),  # same pattern one row down
        ("=$B$2-B3", 4, 2, "=R2C2-R[-1]C"),
        ("=SUM(Long_Monthly!C2:C25)", 2, 2, "=SUM(Long_Monthly!RC[1]:R[23]C[1])"),
        ("=B4/B2", 6, 2, "=R[-2]C/R[-4]C"),
        ("=SUM(RevenueData)", 3, 3, "=SUM(RevenueData)"),  # named range untouched
    ],
)
def test_to_r1c1(formula: str, row: int, col: int, expected: str) -> None:
    assert to_r1c1(formula, row, col) == expected


@pytest.mark.parametrize(
    ("base", "curr", "expected"),
    [
        ("=SUM(Long_Monthly!C2:C21)", "=SUM(Long_Monthly!C2:C25)", True),
        ("=SUM(C2:C21)", "=SUM(C2:C25)", True),
        ("=C14-D14", "=C14-D14*1.1", False),  # extra tokens
        ("=SUM(C2:C21)", "=SUM(D2:D25)", False),  # anchor moved
        ("=SUM(C2:C21)", "=SUM(C2:C21)", False),  # identical, nothing extended
        ("=SUM(C2:C25)", "=SUM(C2:C21)", False),  # shrunk, not extended
    ],
)
def test_differs_only_by_extension(base: str, curr: str, expected: bool) -> None:
    assert _differs_only_by_extension(base, curr) is expected


@pytest.fixture(scope="module")
def baseline(fixture_dir: Path) -> WorkbookSnapshot:
    return load_workbook_snapshot(fixture_dir / "baseline.xlsx")


@pytest.fixture(scope="module")
def current(fixture_dir: Path) -> WorkbookSnapshot:
    return load_workbook_snapshot(fixture_dir / "current.xlsx")


@pytest.fixture(scope="module")
def alignment(baseline: WorkbookSnapshot, current: WorkbookSnapshot) -> WorkbookAlignment:
    return align_workbooks(baseline, current)


@pytest.fixture(scope="module")
def findings(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot, alignment: WorkbookAlignment
) -> list[Finding]:
    return diff_workbook_formulas(baseline, current, alignment)


def _locations(findings: list[Finding], cls: FindingClass) -> set[tuple[str | None, str | None]]:
    return {(f.sheet, f.location) for f in findings if f.finding_class is cls}


def test_error_values_detected(findings: list[Finding], manifest: FixtureManifest) -> None:
    e05, e13 = manifest.defect("E05"), manifest.defect("E13")
    assert _locations(findings, FindingClass.FORMULA_ERROR) == {
        ("Summary", e05.cell),  # error constant
        ("Summary", e13.cell),  # error inside formula text
    }


def test_hardcoded_formula_detected(findings: list[Finding], manifest: FixtureManifest) -> None:
    e02 = manifest.defect("E02")
    hardcoded = [f for f in findings if f.finding_class is FindingClass.FORMULA_HARDCODED]
    assert [(f.sheet, f.location) for f in hardcoded] == [("Long_Monthly", e02.cell)]
    assert hardcoded[0].baseline_value == e02.baseline  # "=C10-D10"


def test_cleared_historical_formula_detected(fixture_dir: Path) -> None:
    base = load_workbook_snapshot(fixture_dir / "baseline.xlsx")
    curr = load_workbook_snapshot(fixture_dir / "current.xlsx")
    curr.sheet("Long_Monthly").cells.pop((10, 5))

    results = diff_workbook_formulas(base, curr, align_workbooks(base, curr))

    removed = [f for f in results if f.finding_class is FindingClass.FORMULA_REMOVED]
    assert [(f.sheet, f.location) for f in removed] == [("Long_Monthly", "E10")]
    assert removed[0].baseline_value == "=C10-D10"


def test_logic_changes_and_expected_extensions(
    findings: list[Finding], manifest: FixtureManifest
) -> None:
    logic = [f for f in findings if f.finding_class is FindingClass.FORMULA_LOGIC_CHANGED]
    unexpected = {(f.sheet, f.location) for f in logic if not f.expected_growth}
    e03, e08 = manifest.defect("E03"), manifest.defect("E08")
    # E08's deviant formula also changed vs baseline — both classes are truthful.
    assert unexpected == {("Long_Monthly", e03.cell), ("Wide_Weekly", e08.cell)}

    # EX06: Summary SUM ranges extended row 21 -> 25 are expected, never errors.
    expected = {(f.sheet, f.location) for f in logic if f.expected_growth}
    assert expected == {("Summary", "B2"), ("Summary", "B3"), ("Summary", "B5")}


def test_inconsistent_formulas_detected(
    findings: list[Finding], manifest: FixtureManifest
) -> None:
    e03, e08 = manifest.defect("E03"), manifest.defect("E08")
    assert _locations(findings, FindingClass.FORMULA_INCONSISTENT) == {
        ("Long_Monthly", e03.cell),  # deviates from its column pattern too
        ("Wide_Weekly", e08.cell),
    }
    outlier = next(
        f
        for f in findings
        if f.finding_class is FindingClass.FORMULA_INCONSISTENT and f.sheet == "Wide_Weekly"
    )
    assert outlier.current_value == e08.current


def test_formula_not_extended_detected(
    findings: list[Finding], manifest: FixtureManifest
) -> None:
    e04 = manifest.defect("E04")
    assert _locations(findings, FindingClass.FORMULA_NOT_EXTENDED) == {
        ("Long_Monthly", e04.cell)  # E25 only; E22-E24 carry the formula
    }


def test_xlsb_degrades_to_error_scan(fixture_dir: Path) -> None:
    base = load_workbook_snapshot(fixture_dir / "baseline.xlsb")
    curr = load_workbook_snapshot(fixture_dir / "current.xlsb")
    aligned = align_workbooks(base, curr)
    results = diff_workbook_formulas(base, curr, aligned)
    assert results == []  # no seeded xlsb errors; formula checks skipped


def test_presence_only_formula_checks_do_not_invent_semantic_changes(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    alignment: WorkbookAlignment,
) -> None:
    for workbook in (baseline, current):
        workbook.formulas_available = False
        workbook.formula_presence_available = True
        workbook.formula_source = None
        for sheet in workbook.sheets:
            for cell in sheet.cells.values():
                cell.formula = None

    results = diff_workbook_formulas(baseline, current, alignment)

    assert _locations(results, FindingClass.FORMULA_HARDCODED) == {
        ("Long_Monthly", "E10")
    }
    assert _locations(results, FindingClass.FORMULA_NOT_EXTENDED) == {
        ("Long_Monthly", "E25")
    }
    assert not _locations(results, FindingClass.FORMULA_LOGIC_CHANGED)
    assert not _locations(results, FindingClass.FORMULA_INCONSISTENT)
