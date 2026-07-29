"""Dependency graph and impact annotation tests (acceptance criterion 5)."""

from pathlib import Path

import pytest

from qc_tool.excel.align import align_workbooks
from qc_tool.excel.dependency import (
    annotate_impacts,
    build_dependency_graph,
    dependents_of,
)
from qc_tool.excel.diff_values import diff_workbook_values
from qc_tool.excel.formulas import diff_workbook_formulas
from qc_tool.findings import FindingClass
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.io.model import (
    CellRecord,
    NamedRange,
    SheetSnapshot,
    WorkbookSnapshot,
)
from tests.fixtures.manifest_schema import FixtureManifest


@pytest.fixture(scope="module")
def current(fixture_dir: Path) -> WorkbookSnapshot:
    return load_workbook_snapshot(fixture_dir / "current.xlsx")


@pytest.fixture(scope="module")
def graph(current: WorkbookSnapshot):
    return build_dependency_graph(current)


def test_cross_sheet_dependents(graph, manifest: FixtureManifest) -> None:
    e01 = manifest.defect("E01")
    dependents = dependents_of(graph, "Long_Monthly", e01.cell or "")
    # The manifest's impact list is a subset (the full chain also passes
    # through Long_Monthly!E7 and Summary!B5).
    assert set(e01.impacts) <= set(dependents)
    assert "Long_Monthly!E7" in dependents


def test_summary_internal_chain(graph) -> None:
    assert dependents_of(graph, "Summary", "B2") == ["Summary!B4", "Summary!B6"]


def test_missing_growth_cell_still_has_impacts(graph, manifest: FixtureManifest) -> None:
    # E25 is empty (the E04 defect) but Summary!B5 sums E2:E25 — the impact
    # of the missing formula must still surface.
    e04 = manifest.defect("E04")
    assert dependents_of(graph, "Long_Monthly", e04.cell or "") == ["Summary!B5"]


def test_named_range_operands_resolve() -> None:
    sheet = SheetSnapshot(
        name="Sheet1",
        visibility="visible",
        max_row=2,
        max_column=2,
        cells={
            (1, 1): CellRecord(row=1, column=1, value=10.0),
            (1, 2): CellRecord(row=1, column=2, value=None, formula="=SUM(NR)"),
        },
    )
    workbook = WorkbookSnapshot(
        source_name="synthetic",
        file_format="xlsx",
        formulas_available=True,
        styles_available=True,
        sheets=[sheet],
        named_ranges=[NamedRange(name="NR", target="Sheet1!$A$1")],
    )
    graph = build_dependency_graph(workbook)
    assert dependents_of(graph, "Sheet1", "A1") == ["Sheet1!B1"]


def test_findings_annotated_with_impacts(
    fixture_dir: Path, current: WorkbookSnapshot, graph, manifest: FixtureManifest
) -> None:
    baseline = load_workbook_snapshot(fixture_dir / "baseline.xlsx")
    alignment = align_workbooks(baseline, current)
    findings = diff_workbook_values(baseline, current, alignment)
    findings += diff_workbook_formulas(baseline, current, alignment)
    annotate_impacts(findings, graph)

    e01 = manifest.defect("E01")
    value_finding = next(
        f
        for f in findings
        if f.finding_class is FindingClass.VALUE_CHANGED
        and (f.sheet, f.location) == ("Long_Monthly", e01.cell)
    )
    assert set(e01.impacts) <= set(value_finding.impacts)

    e02 = manifest.defect("E02")
    hardcode = next(
        f for f in findings if f.finding_class is FindingClass.FORMULA_HARDCODED
    )
    assert hardcode.impacts == e02.impacts  # ["Summary!B5"]

    not_extended = next(
        f for f in findings if f.finding_class is FindingClass.FORMULA_NOT_EXTENDED
    )
    assert not_extended.impacts == ["Summary!B5"]
