"""Value + structure diff tests against the fixture manifest (criteria 1, 2, 4)."""

from pathlib import Path

import pytest

from qc_tool.config.profile import DeliverableProfile
from qc_tool.excel.align import WorkbookAlignment, align_workbooks
from qc_tool.excel.diff_structure import diff_workbook_structure
from qc_tool.excel.diff_values import diff_workbook_values
from qc_tool.findings import Finding, FindingClass
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.io.model import WorkbookSnapshot
from tests.fixtures.manifest_schema import FixtureManifest

REFRESH_PROFILE = DeliverableProfile.model_validate(
    {
        "name": "fixture",
        "excel": {"sheets": {"Dashboard": {"refresh_ranges": ["B2:B5"]}}},
    }
)


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
def value_findings(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot, alignment: WorkbookAlignment
) -> list[Finding]:
    return diff_workbook_values(baseline, current, alignment, REFRESH_PROFILE)


@pytest.fixture(scope="module")
def structure_findings(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot, alignment: WorkbookAlignment
) -> list[Finding]:
    return diff_workbook_structure(baseline, current, alignment)


def _by_class(findings: list[Finding], cls: FindingClass) -> list[Finding]:
    return [f for f in findings if f.finding_class is cls]


def test_seeded_value_changes_detected(
    value_findings: list[Finding], manifest: FixtureManifest
) -> None:
    unexpected = [
        f
        for f in _by_class(value_findings, FindingClass.VALUE_CHANGED)
        if not f.expected_growth
    ]
    locations = {(f.sheet, f.location) for f in unexpected}
    e01, e06 = manifest.defect("E01"), manifest.defect("E06")
    # Exactly the two seeded historical edits — zero false positives (criterion 1+2).
    assert locations == {("Long_Monthly", e01.cell), ("Wide_Weekly", e06.cell)}

    e01_finding = next(f for f in unexpected if f.sheet == "Long_Monthly")
    assert float(e01_finding.baseline_value or "") == float(e01.baseline or "")
    assert float(e01_finding.current_value or "") == float(e01.current or "")

    e06_finding = next(f for f in unexpected if f.sheet == "Wide_Weekly")
    assert e06_finding.baseline_location == e06.baseline_cell  # shift absorbed


def test_kpi_refresh_classified_expected(value_findings: list[Finding]) -> None:
    dashboard_changes = [
        f
        for f in _by_class(value_findings, FindingClass.VALUE_CHANGED)
        if f.sheet == "Dashboard"
    ]
    assert dashboard_changes, "KPI aggregate refresh should still be reported"
    assert all(f.expected_growth for f in dashboard_changes)  # EX03


def test_tolerance_suppresses_small_deltas(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot, alignment: WorkbookAlignment
) -> None:
    tolerant = DeliverableProfile.model_validate(
        {
            "name": "tolerant",
            "tolerance": {"absolute": 2000.0},
            "excel": {"sheets": {"Dashboard": {"refresh_ranges": ["B2:B5"]}}},
        }
    )
    findings = diff_workbook_values(baseline, current, alignment, tolerant)
    unexpected = [
        f
        for f in _by_class(findings, FindingClass.VALUE_CHANGED)
        if not f.expected_growth and f.sheet in {"Long_Monthly", "Wide_Weekly"}
    ]
    assert unexpected == []  # deltas 1234 and 777 are inside tolerance


def test_axis_events(value_findings: list[Finding]) -> None:
    growth_rows = _by_class(value_findings, FindingClass.ROW_GROWTH)
    assert [f.location for f in growth_rows if f.sheet == "Long_Monthly"] == [
        "row 22",
        "row 23",
        "row 24",
        "row 25",
    ]  # EX01
    assert all(f.expected_growth for f in growth_rows)

    deleted_cols = _by_class(value_findings, FindingClass.COLUMN_DELETED)
    assert [(f.sheet, f.baseline_location) for f in deleted_cols] == [
        ("Wide_Weekly", "column D")
    ]  # E07

    growth_cols = _by_class(value_findings, FindingClass.COLUMN_GROWTH)
    assert ("Wide_Weekly", "column U") in {(f.sheet, f.location) for f in growth_cols}  # EX02
    assert ("Dashboard", "column K") in {(f.sheet, f.location) for f in growth_cols}  # EX05

    inserted_cols = _by_class(value_findings, FindingClass.COLUMN_INSERTED)
    assert {(f.sheet, f.location) for f in inserted_cols} == {
        ("Summary", "column C"),
        ("Summary", "column D"),
    }  # seeded error cells arrive as unexpected insertions


def test_number_format_and_style_changes(
    value_findings: list[Finding], manifest: FixtureManifest
) -> None:
    formats = _by_class(value_findings, FindingClass.NUMBER_FORMAT_CHANGED)
    e09 = manifest.defect("E09")
    assert [(f.sheet, f.location, f.baseline_value, f.current_value) for f in formats] == [
        ("Dashboard", "B4", e09.baseline, e09.current)
    ]

    styles = _by_class(value_findings, FindingClass.STYLE_CHANGED)
    assert [(f.sheet, f.location) for f in styles] == [("Dashboard", "A2")]  # E18


def test_sheet_level_structure(structure_findings: list[Finding]) -> None:
    added = _by_class(structure_findings, FindingClass.SHEET_ADDED)
    removed = _by_class(structure_findings, FindingClass.SHEET_REMOVED)
    hidden = _by_class(structure_findings, FindingClass.HIDDEN_CHANGED)
    assert [f.sheet for f in added] == ["New_Analysis"]  # E15
    assert [f.sheet for f in removed] == ["Old_Sheet"]  # E14
    assert [(f.sheet, f.baseline_value, f.current_value) for f in hidden] == [
        ("Params", "visible", "hidden")
    ]  # E11


def test_named_range_findings(structure_findings: list[Finding]) -> None:
    named = _by_class(structure_findings, FindingClass.NAMED_RANGE_CHANGED)
    by_element = {f.element: f for f in named}
    assert set(by_element) == {"KPI_Margin", "RevenueData"}
    assert not by_element["KPI_Margin"].expected_growth  # E12: repointed
    assert by_element["RevenueData"].expected_growth  # EX04: pure extension


def test_chart_findings(structure_findings: list[Finding], manifest: FixtureManifest) -> None:
    charts = _by_class(structure_findings, FindingClass.CHART_SERIES_CHANGED)
    e16 = manifest.defect("E16")
    unexpected = [f for f in charts if not f.expected_growth]
    assert len(unexpected) == 1  # E16: series repointed C -> D
    assert unexpected[0].baseline_value == e16.baseline
    assert unexpected[0].current_value == e16.current
    expected = [f for f in charts if f.expected_growth]
    assert len(expected) == 1  # categories A2:A21 -> A2:A25 is growth


def test_pivot_findings(structure_findings: list[Finding], manifest: FixtureManifest) -> None:
    pivots = _by_class(structure_findings, FindingClass.PIVOT_SOURCE_CHANGED)
    e17 = manifest.defect("E17")
    assert len(pivots) == 1
    finding = pivots[0]
    assert not finding.expected_growth  # shrunk source is an error
    assert finding.baseline_value == f"Long_Monthly!{e17.baseline}"
    assert finding.current_value == f"Long_Monthly!{e17.current}"


def test_no_unpaired_regions(structure_findings: list[Finding]) -> None:
    assert _by_class(structure_findings, FindingClass.REGION_UNPAIRED) == []


@pytest.mark.parametrize(
    ("base", "curr", "expected"),
    [
        ("May-26", "Jun-26", True),  # cadence advanced
        ("Jun-26", "May-26", False),  # regression is a real finding
        ("W20", "W21", True),
        ("W52", "W01", True),  # bounded yearless rollover
        ("W20", "W01", False),  # arbitrary backwards movement is not rollover
        ("North", "South", False),  # not periods
        (100.0, 200.0, False),
    ],
)
def test_period_advanced_rule(base: object, curr: object, expected: bool) -> None:
    from qc_tool.excel.diff_values import _period_advanced

    assert _period_advanced(base, curr) is expected


def test_kpi_refresh_expected_without_profile(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot, alignment: WorkbookAlignment
) -> None:
    """No profile needed: the KPI block contains an advancing period label
    (Cycle May-26 -> Jun-26), so its constant refreshes classify expected."""
    findings = diff_workbook_values(baseline, current, alignment, None)
    changed = _by_class(findings, FindingClass.VALUE_CHANGED)
    dashboard = [f for f in changed if f.sheet == "Dashboard"]
    assert dashboard, "KPI refresh should still be reported"
    assert all(f.expected_growth for f in dashboard)
    unexpected = {(f.sheet, f.location) for f in changed if not f.expected_growth}
    assert unexpected == {("Long_Monthly", "C7"), ("Wide_Weekly", "E2")}


def test_period_regression_is_not_refresh() -> None:
    """A block whose period label moves backward gets no refresh treatment."""
    from qc_tool.io.model import CellRecord, SheetSnapshot, WorkbookSnapshot

    def sheet(period: str, value: float) -> SheetSnapshot:
        return SheetSnapshot(
            name="KPI",
            visibility="visible",
            max_row=2,
            max_column=2,
            cells={
                (1, 1): CellRecord(row=1, column=1, value="Cycle"),
                (1, 2): CellRecord(row=1, column=2, value=period),
                (2, 1): CellRecord(row=2, column=1, value="Total"),
                (2, 2): CellRecord(row=2, column=2, value=value),
            },
        )

    def workbook(period: str, value: float) -> WorkbookSnapshot:
        return WorkbookSnapshot(
            source_name="synthetic",
            file_format="xlsx",
            formulas_available=True,
            styles_available=True,
            sheets=[sheet(period, value)],
        )

    base = workbook("Jun-26", 100.0)
    curr = workbook("May-26", 250.0)  # period went BACKWARD + value changed
    aligned = align_workbooks(base, curr)
    findings = diff_workbook_values(base, curr, aligned, None)
    changed = _by_class(findings, FindingClass.VALUE_CHANGED)
    assert changed, "both changes must be reported"
    assert all(not f.expected_growth for f in changed)  # nothing excused
