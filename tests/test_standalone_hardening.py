"""Step 3: standalone hardening — columnar error populations, XLSB externals."""

from __future__ import annotations

from pathlib import Path

from qc_tool.excel.align import AxisAlignment, RegionAlignment, WorkbookAlignment
from qc_tool.excel.formulas import diff_workbook_formulas
from qc_tool.excel.regions import TableRegion
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingEvidenceTag,
    FindingProvenance,
    FindingSubtype,
    Severity,
)
from qc_tool.io.loader import _xlsb_risks, load_workbook_snapshot
from qc_tool.io.model import (
    CellRecord,
    SheetSnapshot,
    WorkbookRisk,
    WorkbookRiskKind,
    WorkbookSnapshot,
)
from qc_tool.io.xlsb_formula import XlsbFormulaScan
from qc_tool.review import build_pattern_groups
from qc_tool.triage.rules import assign_severity, triage


def _workbook(cells: dict[tuple[int, int], CellRecord]) -> WorkbookSnapshot:
    max_row = max((r for r, _ in cells), default=1)
    max_col = max((c for _, c in cells), default=1)
    return WorkbookSnapshot(
        "standalone.xlsx",
        "xlsx",
        True,
        True,
        formula_presence_available=True,
        sheets=[SheetSnapshot("Raw", "visible", max_row, max_col, cells)],
    )


def _error_column(
    col: int,
    count: int,
    *,
    literal: str = "#N/A",
    start: int = 2,
    formula_presence: bool = False,
    formula: str | None = None,
) -> dict[tuple[int, int], CellRecord]:
    return {
        (row, col): CellRecord(
            row,
            col,
            literal,
            formula=formula,
            is_formula=formula_presence or formula is not None,
        )
        for row in range(start, start + count)
    }


def _preflight_errors(cells: dict[tuple[int, int], CellRecord]):
    findings = diff_workbook_formulas(
        _workbook({}), _workbook(cells), WorkbookAlignment(), cycle=False
    )
    return [f for f in findings if f.finding_class is FindingClass.FORMULA_ERROR]


class TestColumnarErrorPopulations:
    def test_formula_backed_concentrated_population_is_warning(self) -> None:
        cells = _error_column(60, 25, formula_presence=True)
        cells.update(
            {(row, 60): CellRecord(row, 60, float(row)) for row in range(27, 32)}
        )  # 25 of 30 populated = 83%
        cells[(2, 3)] = CellRecord(2, 3, "#REF!")  # scattered error elsewhere

        errors = _preflight_errors(cells)
        population = [
            f for f in errors if f.subtype is FindingSubtype.COLUMNAR_ERROR_POPULATION
        ]
        scattered = [f for f in errors if f.subtype is None]

        assert len(population) == 25
        assert all(assign_severity(f) is Severity.WARNING for f in population)
        assert len({f.event_key for f in population}) == 1
        assert all("error population: 25 of 30" in f.message for f in population)
        assert all(
            {
                FindingEvidenceTag.FORMULA_PRESENCE,
                FindingEvidenceTag.CONCENTRATED_POPULATION,
                FindingEvidenceTag.CONTIGUOUS_POPULATION,
            }
            <= f.evidence_tags
            for f in population
        )
        assert len(scattered) == 1
        assert assign_severity(scattered[0]) is Severity.CRITICAL

    def test_below_minimum_count_stays_critical(self) -> None:
        errors = _preflight_errors(_error_column(5, 19))
        assert all(f.subtype is None for f in errors)
        assert all(assign_severity(f) is Severity.CRITICAL for f in errors)

    def test_low_share_contiguous_value_only_population_stays_critical(self) -> None:
        cells = _error_column(5, 25)
        cells.update(
            {(row, 5): CellRecord(row, 5, float(row)) for row in range(50, 90)}
        )  # 25 of 65 populated = 38%
        errors = _preflight_errors(cells)
        assert all(
            f.subtype is FindingSubtype.COLUMNAR_ERROR_POPULATION for f in errors
        )
        assert all(assign_severity(f) is Severity.CRITICAL for f in errors)
        assert all(
            FindingEvidenceTag.CONTIGUOUS_POPULATION in f.evidence_tags
            for f in errors
        )

    def test_mixed_literals_count_separately(self) -> None:
        cells = _error_column(4, 15, literal="#N/A")
        cells.update(_error_column(4, 15, literal="#VALUE!", start=40))
        # 30 errors in the column but neither literal reaches 20.
        errors = _preflight_errors(cells)
        assert all(f.subtype is None for f in errors)

    def test_mass_rule_qualifies_low_share_floods(self) -> None:
        # 250 identical literals in a huge column: systemic by mass even at
        # a small share (the proven lookup-gap flood shape).
        cells = {
            (row * 4, 2): CellRecord(row * 4, 2, "#N/A")
            for row in range(1, 251)
        }
        cells.update(
            {
                (row, 2): CellRecord(row, 2, float(row))
                for row in range(1001, 5000)
            }
        )
        errors = _preflight_errors(cells)
        population = [
            f for f in errors if f.subtype is FindingSubtype.COLUMNAR_ERROR_POPULATION
        ]
        assert len(population) == 250
        assert all(assign_severity(f) is Severity.CRITICAL for f in population)
        assert len(build_pattern_groups(triage(population))) == 1
        assert all(
            {
                FindingEvidenceTag.CACHED_VALUE_ONLY,
                FindingEvidenceTag.SPARSE_MASS_POPULATION,
            }
            <= f.evidence_tags
            for f in population
        )

    def test_structural_breakage_never_demotes(self) -> None:
        # A column full of #REF! is one incident — but a broken one; no data
        # refresh heals it, so it stays CRITICAL by both share and mass.
        by_share = _preflight_errors(_error_column(3, 25, literal="#REF!"))
        by_mass = _preflight_errors(_error_column(4, 250, literal="#NAME?"))

        assert all(
            f.subtype is FindingSubtype.COLUMNAR_ERROR_POPULATION for f in by_share
        )
        assert all(
            f.subtype is FindingSubtype.COLUMNAR_ERROR_POPULATION for f in by_mass
        )
        assert all(assign_severity(f) is Severity.CRITICAL for f in by_share)
        assert all(assign_severity(f) is Severity.CRITICAL for f in by_mass)
        assert all(
            FindingEvidenceTag.STRUCTURAL_ERROR in f.evidence_tags
            for f in by_share + by_mass
        )

    def test_distinct_columns_and_literals_have_distinct_population_events(self) -> None:
        cells = _error_column(2, 20, literal="#N/A")
        cells.update(_error_column(3, 20, literal="#VALUE!"))

        errors = triage(_preflight_errors(cells))
        groups = build_pattern_groups(errors)

        assert len({finding.event_key for finding in errors}) == 2
        assert len(groups) == 2


def _cycle_errors(
    baseline_cells: dict[tuple[int, int], CellRecord],
    current_cells: dict[tuple[int, int], CellRecord],
) -> list[Finding]:
    rows = max(
        max((row for row, _ in baseline_cells), default=1),
        max((row for row, _ in current_cells), default=1),
    )
    baseline = _workbook(baseline_cells)
    current = _workbook(current_cells)
    region = TableRegion("Raw", 1, 1, rows, 1, "block", None, 1, "none")
    alignment = WorkbookAlignment(
        common_sheets=["Raw"],
        regions={
            "Raw": [
                RegionAlignment(
                    region,
                    region,
                    AxisAlignment(pairs=[(row, row) for row in range(1, rows + 1)]),
                    AxisAlignment(pairs=[(1, 1)]),
                )
            ]
        },
    )
    return triage(
        [
            finding
            for finding in diff_workbook_formulas(
                baseline, current, alignment
            )
            if finding.finding_class is FindingClass.FORMULA_ERROR
        ]
    )


class TestErrorSeverityMatrix:
    def test_inherited_explicit_na_is_info(self) -> None:
        cells = _error_column(1, 1, formula="=NA()")

        error = _cycle_errors(cells, cells)[0]

        assert error.provenance is FindingProvenance.INHERITED
        assert FindingEvidenceTag.EXPLICIT_NA in error.evidence_tags
        assert error.severity is Severity.INFO

    def test_structural_and_new_or_changed_populations_are_critical(self) -> None:
        inherited_structural = _cycle_errors(
            _error_column(1, 20, literal="#REF!", formula_presence=True),
            _error_column(1, 20, literal="#REF!", formula_presence=True),
        )
        new_population = _cycle_errors(
            {(row, 1): CellRecord(row, 1, float(row)) for row in range(1, 21)},
            _error_column(1, 20, start=1, formula_presence=True),
        )
        changed_population = _cycle_errors(
            _error_column(1, 20, literal="#DIV/0!", start=1, formula_presence=True),
            _error_column(1, 20, literal="#N/A", start=1, formula_presence=True),
        )

        assert {finding.provenance for finding in inherited_structural} == {
            FindingProvenance.INHERITED
        }
        assert {finding.provenance for finding in new_population} == {
            FindingProvenance.NEW
        }
        assert {finding.provenance for finding in changed_population} == {
            FindingProvenance.CHANGED
        }
        assert {
            finding.severity
            for finding in inherited_structural + new_population + changed_population
        } == {Severity.CRITICAL}


class TestXlsbExternalLinks:
    def test_scanner_proven_externals_map_to_snapshot_links(self) -> None:
        scan = XlsbFormulaScan(
            formula_cells={},
            risky_features=(
                "external relationships",
                "external workbook links",
                "pivot caches",
            ),
        )
        risks = _xlsb_risks(scan)
        assert risks == [
            WorkbookRisk(WorkbookRiskKind.EXTERNAL_RELATIONSHIP),
            WorkbookRisk(WorkbookRiskKind.EXTERNAL_WORKBOOK_LINK),
        ]
        assert _xlsb_risks(None) == []

    def test_clean_fixture_xlsb_has_no_external_entries(self, fixture_dir: Path) -> None:
        snapshot = load_workbook_snapshot(fixture_dir / "current.xlsb")
        assert snapshot.intrinsic_risks == []
