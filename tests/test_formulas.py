"""Formula QC tests (acceptance criterion 3) against the fixture manifest."""

import time
from collections.abc import Mapping
from pathlib import Path

import pytest

from qc_tool.excel.align import (
    AxisAlignment,
    RegionAlignment,
    WorkbookAlignment,
    align_workbooks,
)
from qc_tool.excel.formula_tokens import FormulaDiffKind, FormulaDiffSegment
from qc_tool.excel.formulas import (
    _FORMULA_PAIR_ANALYSIS_MEMO_CAP,
    FormulaComparisonTelemetry,
    FormulaPairAnalysis,
    FormulaPairAnalysisMemo,
    _differs_only_by_extension,
    diff_workbook_formulas,
    formula_text_comparable,
    formula_token_diff,
    rewrite_renamed_sheet_references,
    to_r1c1,
)
from qc_tool.excel.regions import TableRegion
from qc_tool.findings import Finding, FindingClass, FindingEvidenceTag
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.io.model import CellRecord, FormulaTextCoverage, SheetSnapshot, WorkbookSnapshot
from tests.fixtures.manifest_schema import FixtureManifest


def test_partial_formula_text_comparable_requires_same_nonempty_source() -> None:
    baseline = WorkbookSnapshot("base.xlsb", "xlsb", False, False)
    current = WorkbookSnapshot("current.xlsb", "xlsb", False, False)
    baseline.formula_source = current.formula_source = "libreoffice:test"
    baseline.formula_text_coverage = FormulaTextCoverage(
        state="partial", expected_count=2, merged_count=1, missing_count=1
    )
    current.formula_text_coverage = FormulaTextCoverage(
        state="partial", expected_count=2, merged_count=1, missing_count=1
    )

    assert formula_text_comparable(baseline, current)

    current.formula_source = "excel:test"
    assert not formula_text_comparable(baseline, current)

    current.formula_source = baseline.formula_source
    current.formula_text_coverage = FormulaTextCoverage()
    assert not formula_text_comparable(baseline, current)


@pytest.mark.parametrize(
    ("formula", "row", "col", "expected"),
    [
        ("=C10-D10", 10, 5, "=RC[-2]-RC[-1]"),
        ("=C11-D11", 11, 5, "=RC[-2]-RC[-1]"),  # same pattern one row down
        ("=$B$2-B3", 4, 2, "=R2C2-R[-1]C"),
        ("=SUM(Long_Monthly!C2:C25)", 2, 2, "=SUM(Long_Monthly!RC[1]:R[23]C[1])"),
        ("=B4/B2", 6, 2, "=R[-2]C/R[-4]C"),
        ("=SUM(RevenueData)", 3, 3, "=SUM(RevenueData)"),  # named range untouched
        ("=B2#", 5, 5, "=_xlfn.ANCHORARRAY(R[-3]C[-3])"),
        (
            "=_xlfn.ANCHORARRAY(B2)",
            5,
            5,
            "=_xlfn.ANCHORARRAY(R[-3]C[-3])",
        ),
        ("=@A1", 5, 5, "=_xlfn.SINGLE(R[-4]C[-4])"),
        ('="B2# and @A1"', 5, 5, '="B2# and @A1"'),
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


# --- plan-20260913 Step 3: confirmed-rename formula normalization ----------


@pytest.mark.parametrize(
    ("formula", "rename_map", "expected"),
    [
        ("=Sheet2025!A1", {"Sheet2025": "Sheet2026"}, "=Sheet2026!A1"),
        ("='Sheet 2025'!A1", {"Sheet 2025": "Sheet 2026"}, "='Sheet 2026'!A1"),
        # Untouched: no rename applies to this sheet.
        ("=Other!A1", {"Sheet2025": "Sheet2026"}, "=Other!A1"),
        # Untouched: same-sheet reference has no qualifying prefix at all.
        ("=A1+B1", {"Sheet2025": "Sheet2026"}, "=A1+B1"),
        # A string literal that happens to contain the old name is untouched.
        ('="Sheet2025 report"', {"Sheet2025": "Sheet2026"}, '="Sheet2025 report"'),
        # Empty map is always a no-op.
        ("=Sheet2025!A1", {}, "=Sheet2025!A1"),
    ],
)
def test_rewrite_renamed_sheet_references(
    formula: str, rename_map: Mapping[str, str], expected: str
) -> None:
    assert rewrite_renamed_sheet_references(formula, rename_map) == expected


def test_a_pure_rename_produces_no_formula_logic_changed_finding() -> None:
    """A formula that only differs because its own sheet was renamed must
    not be reported -- the rename itself is a separate structural finding
    (SHEET_RENAMED), not a per-cell formula cascade.
    """
    base_sheet = SheetSnapshot(
        name="Sheet2025",
        visibility="visible",
        max_row=1,
        max_column=1,
        cells={(1, 1): CellRecord(row=1, column=1, value=10, formula="=Sheet2025!B1")},
    )
    curr_sheet = SheetSnapshot(
        name="Sheet2026",
        visibility="visible",
        max_row=1,
        max_column=1,
        cells={(1, 1): CellRecord(row=1, column=1, value=10, formula="=Sheet2026!B1")},
    )
    baseline = WorkbookSnapshot(
        "b.xlsx", "xlsx", True, True, formula_presence_available=True
    )
    baseline.sheets = [base_sheet]
    current = WorkbookSnapshot(
        "c.xlsx", "xlsx", True, True, formula_presence_available=True
    )
    current.sheets = [curr_sheet]
    region = RegionAlignment(
        baseline=TableRegion("Sheet2025", 1, 1, 1, 1, "block", None, None, "none"),
        current=TableRegion("Sheet2026", 1, 1, 1, 1, "block", None, None, "none"),
        rows=AxisAlignment(pairs=[(1, 1)]),
        columns=AxisAlignment(pairs=[(1, 1)]),
    )
    alignment = WorkbookAlignment(
        common_sheets=["Sheet2026"],
        renamed_sheets={"Sheet2026": "Sheet2025"},
        regions={"Sheet2026": [region]},
    )

    findings = diff_workbook_formulas(baseline, current, alignment)

    assert [f.finding_class for f in findings] == []


def test_a_real_logic_change_still_reports_despite_a_confirmed_rename() -> None:
    """A rename must never mask a genuine formula-logic change on the same
    sheet -- unmapped/real differences remain findings.
    """
    base_sheet = SheetSnapshot(
        name="Sheet2025",
        visibility="visible",
        max_row=1,
        max_column=1,
        cells={(1, 1): CellRecord(row=1, column=1, value=10, formula="=Sheet2025!B1")},
    )
    curr_sheet = SheetSnapshot(
        name="Sheet2026",
        visibility="visible",
        max_row=1,
        max_column=1,
        cells={(1, 1): CellRecord(row=1, column=1, value=10, formula="=Sheet2026!B2")},
    )
    baseline = WorkbookSnapshot(
        "b.xlsx", "xlsx", True, True, formula_presence_available=True
    )
    baseline.sheets = [base_sheet]
    current = WorkbookSnapshot(
        "c.xlsx", "xlsx", True, True, formula_presence_available=True
    )
    current.sheets = [curr_sheet]
    region = RegionAlignment(
        baseline=TableRegion("Sheet2025", 1, 1, 1, 1, "block", None, None, "none"),
        current=TableRegion("Sheet2026", 1, 1, 1, 1, "block", None, None, "none"),
        rows=AxisAlignment(pairs=[(1, 1)]),
        columns=AxisAlignment(pairs=[(1, 1)]),
    )
    alignment = WorkbookAlignment(
        common_sheets=["Sheet2026"],
        renamed_sheets={"Sheet2026": "Sheet2025"},
        regions={"Sheet2026": [region]},
    )

    findings = diff_workbook_formulas(baseline, current, alignment)

    assert [f.finding_class for f in findings] == [FindingClass.FORMULA_LOGIC_CHANGED]


def test_formula_token_diff_marks_one_reference_replacement() -> None:
    assert formula_token_diff("=A1", "=B1", "C1", "C1") == (
        FormulaDiffSegment("=", FormulaDiffKind.EQUAL),
        FormulaDiffSegment("RC[-2]", FormulaDiffKind.REMOVED),
        FormulaDiffSegment("RC[-1]", FormulaDiffKind.ADDED),
    )


def test_formula_token_diff_normalizes_each_formula_at_its_own_host() -> None:
    assert formula_token_diff("=A1", "=B2", "B1", "C2") == (
        FormulaDiffSegment("=RC[-1]", FormulaDiffKind.EQUAL),
    )


def test_formula_token_diff_coalesces_identical_formula() -> None:
    assert formula_token_diff("=A1+B1", "=A1+B1", "C1", "C1") == (
        FormulaDiffSegment("=RC[-2]+RC[-1]", FormulaDiffKind.EQUAL),
    )


def test_formula_token_diff_preserves_text_operands() -> None:
    segments = formula_token_diff(
        '="Jan-26"',
        '="Feb-26"',
        "A1",
        "A1",
    )

    assert any(
        segment.kind is FormulaDiffKind.REMOVED and '"Jan-26"' in segment.text
        for segment in segments
    )
    assert any(
        segment.kind is FormulaDiffKind.ADDED and '"Feb-26"' in segment.text
        for segment in segments
    )


def test_formula_token_diff_handles_long_let_formula_without_losing_structure() -> None:
    segments = formula_token_diff(
        "=LET(x,A1,y,B1,x+y)",
        "=LET(x,A1,y,C1,x+y)",
        "E1",
        "E1",
    )

    assert segments[0].kind is FormulaDiffKind.EQUAL
    assert segments[0].text.startswith("=LET(")
    assert any(segment.kind is FormulaDiffKind.REMOVED for segment in segments)
    assert any(segment.kind is FormulaDiffKind.ADDED for segment in segments)
    assert "".join(segment.text for segment in segments).endswith("x+y)")


@pytest.mark.parametrize(
    ("baseline", "current", "baseline_location", "current_location"),
    [
        (None, "=A1", "A1", "A1"),
        ("=A1", None, "A1", "A1"),
        ("A1", "=B1", "A1", "A1"),
        ("=A1", "B1", "A1", "A1"),
        ("=A1", "=B1", "not-a-cell", "A1"),
        ("=A1", "=B1", "A1", "not-a-cell"),
    ],
)
def test_formula_token_diff_returns_empty_on_unrenderable_input(
    baseline: str | None,
    current: str | None,
    baseline_location: str,
    current_location: str,
) -> None:
    assert formula_token_diff(
        baseline,
        current,
        baseline_location,
        current_location,
    ) == ()


def test_formula_token_diff_keeps_named_ranges_as_symbols() -> None:
    assert formula_token_diff(
        "=SUM(RevenueData)",
        "=SUM(RevenueData)",
        "B3",
        "B3",
    ) == (
        FormulaDiffSegment("=SUM(RevenueData)", FormulaDiffKind.EQUAL),
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


def test_formula_comparison_telemetry_is_additive_only(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    alignment: WorkbookAlignment,
    findings: list[Finding],
) -> None:
    """Passing a telemetry object changes no finding, only accumulates counters."""
    telemetry = FormulaComparisonTelemetry()
    with_telemetry = diff_workbook_formulas(
        baseline, current, alignment, telemetry=telemetry
    )

    assert with_telemetry == findings

    assert telemetry.error_scan_cells > 0
    assert telemetry.error_scan_seconds >= 0.0

    assert telemetry.paired_pairs_considered > 0
    # Most paired formula cells are unchanged at the same host coordinate --
    # exactly the population the exact-text/same-host shortcut skips.
    assert 0 < telemetry.exact_text_same_host_pairs <= telemetry.paired_pairs_considered
    non_shortcut_pairs = (
        telemetry.paired_pairs_considered - telemetry.exact_text_same_host_pairs
    )
    # Every remaining pair always normalizes its baseline side (never memoized,
    # single-use); the memo may additionally avoid the current side's call, so
    # paired-only `to_r1c1` invocations sit in [non_shortcut, non_shortcut*2].
    assert non_shortcut_pairs <= telemetry.normalization_calls <= non_shortcut_pairs * 2
    assert telemetry.normalization_seconds >= 0.0
    assert telemetry.paired_traversal_seconds >= telemetry.normalization_seconds

    # `*_normalization_calls` count actual to_r1c1 invocations (misses only);
    # the shared memo is exercised by both the paired and consistency checks.
    assert telemetry.consistency_normalization_calls >= 0
    assert telemetry.consistency_normalization_seconds >= 0.0
    assert telemetry.consistency_seconds >= telemetry.consistency_normalization_seconds
    assert telemetry.memo_hits >= 0
    assert telemetry.memo_misses > 0

    assert telemetry.wrapper_reference_seconds >= 0.0
    assert telemetry.extension_seconds >= 0.0
    assert telemetry.extension_findings_seconds >= 0.0
    assert telemetry.finding_construction_seconds >= 0.0
    assert telemetry.complexity_assessment_seconds == 0.0


def test_run_qc_records_complexity_assessment_seconds_on_shared_telemetry(
    fixture_dir: Path,
) -> None:
    """plan-20260908-phase-b-guest-performance-followup.md: `run_qc()`'s own
    `assess_workbook_complexity()` call also lives inside the
    `RunPhase.COMPARING_FORMULAS` boundary, so its cost must land on the same
    telemetry object passed in rather than staying invisible to a
    `comparing_formulas` phase-duration reconciliation.
    """
    from qc_tool.engine import run_qc

    telemetry = FormulaComparisonTelemetry()
    result = run_qc(
        baseline_excel=fixture_dir / "baseline.xlsx",
        current_excel=fixture_dir / "current.xlsx",
        _formula_telemetry=telemetry,
    )

    assert len(list(result.findings)) > 0
    assert telemetry.paired_pairs_considered > 0
    assert telemetry.complexity_assessment_seconds >= 0.0


def test_formula_comparison_telemetry_default_omits_all_cost() -> None:
    """A fresh telemetry object starts at exactly zero on every field."""
    telemetry = FormulaComparisonTelemetry()
    assert telemetry == FormulaComparisonTelemetry(
        error_scan_seconds=0.0,
        error_scan_cells=0,
        paired_traversal_seconds=0.0,
        paired_pairs_considered=0,
        exact_text_same_host_pairs=0,
        normalization_calls=0,
        normalization_seconds=0.0,
        memo_hits=0,
        memo_misses=0,
        consistency_seconds=0.0,
        consistency_normalization_calls=0,
        consistency_normalization_seconds=0.0,
        wrapper_reference_seconds=0.0,
        extension_seconds=0.0,
        extension_findings_seconds=0.0,
        finding_construction_seconds=0.0,
    )


def test_normalization_memo_cap_fallback_stays_correct() -> None:
    """Past its cap the memo recomputes -- never caches -- but stays correct."""
    from qc_tool.excel.formulas import _CurrentNormalizationMemo

    memo = _CurrentNormalizationMemo(cap=1)
    first, first_hit = memo.normalize(1, 1, "=A1")
    assert not first_hit
    second, second_hit = memo.normalize(1, 1, "=A1")
    assert second_hit
    assert second == first == to_r1c1("=A1", 1, 1)

    # A second distinct coordinate exceeds the cap of 1: never cached, but
    # every call still recomputes the correct value.
    third, third_hit = memo.normalize(2, 2, "=B2")
    assert not third_hit
    fourth, fourth_hit = memo.normalize(2, 2, "=B2")
    assert not fourth_hit
    assert fourth == third == to_r1c1("=B2", 2, 2)


def test_normalization_memo_is_keyed_by_coordinate_not_text() -> None:
    """Two different coordinates with identical formula text stay distinct."""
    from qc_tool.excel.formulas import _CurrentNormalizationMemo

    memo = _CurrentNormalizationMemo()
    at_a1, _ = memo.normalize(1, 1, "=B1")
    at_c3, _ = memo.normalize(3, 3, "=B1")
    assert at_a1 != at_c3  # R1C1 output is host-relative


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


# --- Step 4b: coordinate-aware diff under partial formula-text coverage ---


def _long_column_workbook(formulas: dict[int, str | None]) -> WorkbookSnapshot:
    """A minimal 3-row "long" column: row 1 is a text key, rows 2-4 formulas."""
    cells: dict[tuple[int, int], CellRecord] = {
        (1, 1): CellRecord(1, 1, "Key"),
        (1, 2): CellRecord(1, 2, "Value"),
    }
    for row in (2, 3, 4):
        cells[(row, 1)] = CellRecord(row, 1, f"row{row}")
        cells[(row, 2)] = CellRecord(
            row, 2, None, formula=formulas.get(row), is_formula=True
        )
    return WorkbookSnapshot(
        source_name="partial.xlsb",
        file_format="xlsb",
        formulas_available=False,
        styles_available=False,
        formula_presence_available=True,
        formula_source="test-engine",
        formula_text_coverage=FormulaTextCoverage(state="partial", expected_count=3),
        sheets=[SheetSnapshot("Data", "visible", 4, 2, cells)],
    )


def _one_region_alignment() -> WorkbookAlignment:
    region = TableRegion("Data", 1, 1, 4, 2, "long", 1, 1, "none")
    alignment = RegionAlignment(
        baseline=region,
        current=region,
        rows=AxisAlignment(pairs=[(2, 2), (3, 3), (4, 4)]),
        columns=AxisAlignment(pairs=[(1, 1), (2, 2)]),
    )
    return WorkbookAlignment(common_sheets=["Data"], regions={"Data": [alignment]})


def test_partial_coverage_still_compares_paired_cells_with_merged_text() -> None:
    baseline = _long_column_workbook({2: "=A2*2", 3: "=A3*2", 4: "=A4*2"})
    current = _long_column_workbook({2: "=A2*2", 3: "=A3*3", 4: None})

    findings = diff_workbook_formulas(baseline, current, _one_region_alignment())

    logic_changed = _locations(findings, FindingClass.FORMULA_LOGIC_CHANGED)
    assert logic_changed == {("Data", "B3")}  # row 4's missing text is silently skipped


def test_partial_coverage_skips_consistency_for_an_incomplete_run() -> None:
    baseline = _long_column_workbook({2: "=A2*2", 3: "=A3*2", 4: "=A4*2"})
    current_incomplete = _long_column_workbook({2: "=A2*2", 3: "=A3*3", 4: None})

    findings = diff_workbook_formulas(
        baseline, current_incomplete, _one_region_alignment()
    )

    assert not _locations(findings, FindingClass.FORMULA_INCONSISTENT)

    # Control: the same deviation with every cell in the run merged still
    # flags row 3 -- proving the run-completeness gate, not the deviation
    # detector, suppressed the finding above.
    current_complete = _long_column_workbook({2: "=A2*2", 3: "=A3*3", 4: "=A4*2"})

    complete_findings = diff_workbook_formulas(
        baseline, current_complete, _one_region_alignment()
    )

    assert _locations(complete_findings, FindingClass.FORMULA_INCONSISTENT) == {
        ("Data", "B3")
    }


# --- B5: definition-level (adapter-supplied) R1C1 --------------------------


def _set_formula_r1c1(
    workbook: WorkbookSnapshot, sheet_name: str, values: dict[tuple[int, int], str]
) -> None:
    sheet = workbook.sheet(sheet_name)
    for coordinate, text in values.items():
        sheet.cells[coordinate].formula_r1c1 = text


def test_paired_and_consistency_checks_use_adapter_supplied_r1c1_without_recomputing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """When every paired/run cell already carries a correct `formula_r1c1`
    (as the native kernel adapter populates it), `to_r1c1` must never be
    called, and the findings must be byte-identical to the computed-R1C1
    path (Criterion 13(a)'s digest-equality intent, exercised on a fixture).
    """
    baseline = _long_column_workbook({2: "=A2*2", 3: "=A3*2", 4: "=A4*2"})
    current = _long_column_workbook({2: "=A2*2", 3: "=A3*3", 4: "=A4*2"})
    alignment = _one_region_alignment()

    # Ground truth: findings computed the existing way (no formula_r1c1).
    expected = diff_workbook_formulas(baseline, current, alignment)

    _set_formula_r1c1(
        baseline,
        "Data",
        {(2, 2): "=RC[-1]*2", (3, 2): "=RC[-1]*2", (4, 2): "=RC[-1]*2"},
    )
    _set_formula_r1c1(
        current,
        "Data",
        {(2, 2): "=RC[-1]*2", (3, 2): "=RC[-1]*3", (4, 2): "=RC[-1]*2"},
    )

    import qc_tool.excel.formulas as formulas_module

    def _fail_if_called(formula: str, row: int, col: int) -> str:
        raise AssertionError("to_r1c1 must not be called when formula_r1c1 is set")

    monkeypatch.setattr(formulas_module, "to_r1c1", _fail_if_called)

    actual = diff_workbook_formulas(baseline, current, alignment)

    assert _locations(actual, FindingClass.FORMULA_LOGIC_CHANGED) == _locations(
        expected, FindingClass.FORMULA_LOGIC_CHANGED
    )
    assert _locations(actual, FindingClass.FORMULA_INCONSISTENT) == _locations(
        expected, FindingClass.FORMULA_INCONSISTENT
    )
    assert [f.model_dump(mode="json") for f in actual] == [
        f.model_dump(mode="json") for f in expected
    ]


def test_paired_findings_fall_back_to_to_r1c1_when_formula_r1c1_is_absent_on_one_side() -> None:
    """A partially-populated formula_r1c1 (e.g. one adapter side is native,
    the other is OOXML/Excel/LibreOffice in a mixed scenario) must still
    compare correctly by falling back per-cell, not per-run.
    """
    baseline = _long_column_workbook({2: "=A2*2", 3: "=A3*2", 4: "=A4*2"})
    current = _long_column_workbook({2: "=A2*2", 3: "=A3*3", 4: "=A4*2"})
    alignment = _one_region_alignment()

    # Only the baseline side carries adapter-supplied R1C1; current falls
    # back to to_r1c1() entirely.
    _set_formula_r1c1(
        baseline,
        "Data",
        {(2, 2): "=RC[-1]*2", (3, 2): "=RC[-1]*2", (4, 2): "=RC[-1]*2"},
    )

    findings = diff_workbook_formulas(baseline, current, alignment)

    assert _locations(findings, FindingClass.FORMULA_LOGIC_CHANGED) == {("Data", "B3")}


# --- plan-20260910 Step 5: FormulaPairAnalysis memoization ------------------


def _formula_column_workbook(
    row_formulas: Mapping[int, str | None], *, max_row: int
) -> WorkbookSnapshot:
    """A minimal "long" column workbook with one formula cell per row in
    `row_formulas` (column B) -- generalizes `_long_column_workbook` to an
    arbitrary row set for Step 5's memoization tests.
    """
    cells: dict[tuple[int, int], CellRecord] = {
        (1, 1): CellRecord(1, 1, "Key"),
        (1, 2): CellRecord(1, 2, "Value"),
    }
    for row, formula in row_formulas.items():
        cells[(row, 1)] = CellRecord(row, 1, f"row{row}")
        cells[(row, 2)] = CellRecord(row, 2, None, formula=formula, is_formula=True)
    return WorkbookSnapshot(
        source_name="pair-analysis.xlsb",
        file_format="xlsb",
        formulas_available=False,
        styles_available=False,
        formula_presence_available=True,
        formula_source="test-engine",
        formula_text_coverage=FormulaTextCoverage(
            state="partial", expected_count=len(row_formulas)
        ),
        sheets=[SheetSnapshot("Data", "visible", max_row, 2, cells)],
    )


def _alignment_for_rows(rows: list[int], max_row: int) -> WorkbookAlignment:
    region = TableRegion("Data", 1, 1, max_row, 2, "long", 1, 1, "none")
    alignment = RegionAlignment(
        baseline=region,
        current=region,
        rows=AxisAlignment(pairs=[(row, row) for row in rows]),
        columns=AxisAlignment(pairs=[(1, 1), (2, 2)]),
    )
    return WorkbookAlignment(common_sheets=["Data"], regions={"Data": [alignment]})


def _diverse_construct_rows() -> tuple[dict[int, str], dict[int, str]]:
    """One row per notable classification construct -- a plain logic change,
    a range extension, an exact wrapper, an added reference with no wrapper,
    an absolute/relative mix, a sheet-qualified reference, and a
    whole-column reference -- each a DISTINCT canonical key (no repeats), so
    a parity check here proves the memo behaves correctly on a
    first-and-only MISS, complementing the repeated-pattern tests below that
    prove HIT behavior.
    """
    base = {
        2: "=B20",  # plain logic change
        3: "=SUM(C2:C10)",  # range extension
        4: "=B4",  # exact wrapper (curr wraps it in ROUND)
        5: "=B5",  # added reference, no wrapper
        6: "=$B$2+B6",  # absolute + relative mix
        7: "=Other!A1+B7",  # sheet-qualified reference
        8: "=SUM(C:C)",  # whole-column reference
    }
    curr = {
        2: "=B30",
        3: "=SUM(C2:C15)",
        4: "=ROUND(B4,2)",
        5: "=B5+C5",
        6: "=$B$2+B6*2",
        7: "=Other!A1+B7*2",
        8: "=SUM(C:C)*2",
    }
    return base, curr


def test_pair_analysis_memo_is_transparent_on_diverse_single_occurrence_constructs() -> (
    None
):
    base_formulas, curr_formulas = _diverse_construct_rows()
    rows = sorted(base_formulas)
    max_row = max(rows) + 1
    baseline = _formula_column_workbook(base_formulas, max_row=max_row)
    current = _formula_column_workbook(curr_formulas, max_row=max_row)
    alignment = _alignment_for_rows(rows, max_row)

    expected = diff_workbook_formulas(
        baseline, current, alignment, _use_native_delta=False
    )
    actual = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        pair_analysis_memo=FormulaPairAnalysisMemo(),
        _use_native_delta=False,
    )

    assert [f.model_dump(mode="json") for f in actual] == [
        f.model_dump(mode="json") for f in expected
    ]
    assert _locations(actual, FindingClass.FORMULA_LOGIC_CHANGED)


def _repeated_pattern_rows(
    count: int, *, start_row: int = 20
) -> tuple[dict[int, str], dict[int, str], list[int]]:
    rows = list(range(start_row, start_row + count))
    base_formulas = {row: f"=A{row}*2" for row in rows}
    curr_formulas = {row: f"=A{row}*3" for row in rows}
    return base_formulas, curr_formulas, rows


def test_pair_analysis_memo_reuses_the_single_classification_across_every_repeat(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """15 rows share one uniform relative pattern -- one canonical key. With
    the memo enabled, wrapper detection (the expensive, tokenization-based
    structural match that is genuinely a pure function of the normalized
    key) must run exactly once for that key, not once per row. The
    extension check runs on RAW formula text and is deliberately never
    memoized (see FormulaPairAnalysis's own docstring for why caching it
    was a real, confirmed bug) -- it must still run once per row, every
    time.
    """
    import qc_tool.excel.formulas as formulas_module

    base_formulas, curr_formulas, rows = _repeated_pattern_rows(15)
    max_row = max(rows) + 1
    baseline = _formula_column_workbook(base_formulas, max_row=max_row)
    current = _formula_column_workbook(curr_formulas, max_row=max_row)
    alignment = _alignment_for_rows(rows, max_row)

    extension_call_count = 0
    real_extension_check = formulas_module._differs_only_by_extension

    def _counting_extension_check(base_formula: str, curr_formula: str) -> bool:
        nonlocal extension_call_count
        extension_call_count += 1
        return real_extension_check(base_formula, curr_formula)

    monkeypatch.setattr(
        formulas_module, "_differs_only_by_extension", _counting_extension_check
    )

    wrapper_call_count = 0
    real_detect_wrapper = formulas_module.detect_formula_wrapper

    def _counting_detect_wrapper(base_norm: str, curr_norm: str):
        nonlocal wrapper_call_count
        wrapper_call_count += 1
        return real_detect_wrapper(base_norm, curr_norm)

    monkeypatch.setattr(
        formulas_module, "detect_formula_wrapper", _counting_detect_wrapper
    )

    memo = FormulaPairAnalysisMemo()
    findings = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        pair_analysis_memo=memo,
        _use_native_delta=False,
    )

    assert extension_call_count == len(rows)
    assert wrapper_call_count == 1
    assert len(memo) == 1
    assert len(_locations(findings, FindingClass.FORMULA_LOGIC_CHANGED)) == len(rows)


def test_pair_analysis_memo_recomputes_added_reference_per_occurrence_of_the_same_key() -> (
    None
):
    """Regression for the real-data bug found during plan-20260910 Step 8
    guest validation: two occurrences that normalize to the IDENTICAL
    (base_r1c1, curr_r1c1) canonical key can legitimately disagree on
    whether a reference was added, because that check runs on each
    occurrence's own RAW (pre-normalization) formula text, not the shared
    key. Forces two rows to share one canonical key via an explicit
    ``formula_r1c1`` override (bypassing ``to_r1c1``'s own computation)
    while giving each row raw formula text that must independently resolve
    to a DIFFERENT ADDED_REFERENCE verdict. With the memo enabled, each
    row's evidence_tags must still be computed correctly and independently
    -- never a stale value carried over from the other row's occurrence of
    the same key.
    """
    shared_base_r1c1 = "=X1+X2"
    shared_curr_r1c1 = "=X1+X2+X3"
    cells: dict[tuple[int, int], CellRecord] = {
        (1, 1): CellRecord(1, 1, "Key"),
        (1, 2): CellRecord(1, 2, "Value"),
        (20, 1): CellRecord(20, 1, "row20"),
        # Row 20: current genuinely adds a reference (C1) absent from
        # baseline -- ADDED_REFERENCE must be present.
        (20, 2): CellRecord(
            20, 2, None, formula="=B1+A1", formula_r1c1=shared_base_r1c1, is_formula=True
        ),
        (21, 1): CellRecord(21, 1, "row21"),
        # Row 21: shares the SAME canonical key (forced via formula_r1c1),
        # but its own raw text already references C2 on both sides -- only
        # a literal changed -- so ADDED_REFERENCE must be ABSENT.
        (21, 2): CellRecord(
            21, 2, None, formula="=B2+A2+C2*1", formula_r1c1=shared_base_r1c1, is_formula=True
        ),
    }
    baseline = WorkbookSnapshot(
        source_name="pair-analysis-conflict.xlsb",
        file_format="xlsb",
        formulas_available=False,
        styles_available=False,
        formula_presence_available=True,
        formula_source="test-engine",
        formula_text_coverage=FormulaTextCoverage(state="partial", expected_count=2),
        sheets=[SheetSnapshot("Data", "visible", 21, 2, cells)],
    )
    curr_cells: dict[tuple[int, int], CellRecord] = {
        (1, 1): CellRecord(1, 1, "Key"),
        (1, 2): CellRecord(1, 2, "Value"),
        (20, 1): CellRecord(20, 1, "row20"),
        (20, 2): CellRecord(
            20,
            2,
            None,
            formula="=B1+A1+C1",
            formula_r1c1=shared_curr_r1c1,
            is_formula=True,
        ),
        (21, 1): CellRecord(21, 1, "row21"),
        (21, 2): CellRecord(
            21,
            2,
            None,
            formula="=B2+A2+C2*2",
            formula_r1c1=shared_curr_r1c1,
            is_formula=True,
        ),
    }
    current = WorkbookSnapshot(
        source_name="pair-analysis-conflict.xlsb",
        file_format="xlsb",
        formulas_available=False,
        styles_available=False,
        formula_presence_available=True,
        formula_source="test-engine",
        formula_text_coverage=FormulaTextCoverage(state="partial", expected_count=2),
        sheets=[SheetSnapshot("Data", "visible", 21, 2, curr_cells)],
    )
    alignment = _alignment_for_rows([20, 21], 22)

    for memo in (None, FormulaPairAnalysisMemo()):
        findings = diff_workbook_formulas(
            baseline,
            current,
            alignment,
            pair_analysis_memo=memo,
            _use_native_delta=False,
        )
        by_location = {f.location: f for f in findings}
        assert FindingEvidenceTag.ADDED_REFERENCE in by_location["B20"].evidence_tags
        assert (
            FindingEvidenceTag.ADDED_REFERENCE
            not in by_location["B21"].evidence_tags
        )


def test_pair_analysis_memo_produces_byte_identical_findings_on_a_repeated_pattern() -> (
    None
):
    base_formulas, curr_formulas, rows = _repeated_pattern_rows(15)
    max_row = max(rows) + 1
    baseline = _formula_column_workbook(base_formulas, max_row=max_row)
    current = _formula_column_workbook(curr_formulas, max_row=max_row)
    alignment = _alignment_for_rows(rows, max_row)

    expected = diff_workbook_formulas(
        baseline, current, alignment, _use_native_delta=False
    )
    actual = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        pair_analysis_memo=FormulaPairAnalysisMemo(),
        _use_native_delta=False,
    )

    assert [f.model_dump(mode="json") for f in actual] == [
        f.model_dump(mode="json") for f in expected
    ]


def test_pair_analysis_memo_produces_byte_identical_findings_on_the_manifest_fixture(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot, alignment: WorkbookAlignment
) -> None:
    """Real-fixture regression check: enabling the memo on the standard
    E01-E18 manifest pair changes zero findings.
    """
    expected = diff_workbook_formulas(
        baseline, current, alignment, _use_native_delta=False
    )
    actual = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        pair_analysis_memo=FormulaPairAnalysisMemo(),
        _use_native_delta=False,
    )

    assert [f.model_dump(mode="json") for f in actual] == [
        f.model_dump(mode="json") for f in expected
    ]


def test_pair_analysis_memo_tracks_hits_and_misses_in_telemetry() -> None:
    base_formulas, curr_formulas, rows = _repeated_pattern_rows(15)
    max_row = max(rows) + 1
    baseline = _formula_column_workbook(base_formulas, max_row=max_row)
    current = _formula_column_workbook(curr_formulas, max_row=max_row)
    alignment = _alignment_for_rows(rows, max_row)

    telemetry = FormulaComparisonTelemetry()
    diff_workbook_formulas(
        baseline,
        current,
        alignment,
        telemetry=telemetry,
        pair_analysis_memo=FormulaPairAnalysisMemo(),
        _use_native_delta=False,
    )

    assert telemetry.pair_analysis_memo_misses == 1
    assert telemetry.pair_analysis_memo_hits == len(rows) - 1


def test_pair_analysis_memo_omitted_leaves_telemetry_counters_at_zero(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    alignment: WorkbookAlignment,
) -> None:
    telemetry = FormulaComparisonTelemetry()
    diff_workbook_formulas(baseline, current, alignment, telemetry=telemetry)

    assert telemetry.pair_analysis_memo_hits == 0
    assert telemetry.pair_analysis_memo_misses == 0


def test_pair_analysis_memo_cap_fallback_stays_correct() -> None:
    """Past its cap the memo recomputes -- never caches -- but stays
    correct, mirroring `_CurrentNormalizationMemo`'s own established
    cap-fallback contract.
    """
    memo = FormulaPairAnalysisMemo(cap=1)
    first = FormulaPairAnalysis(
        wrapper_kind=None,
        wrapper_exact=None,
        event_key="",
    )
    memo.put("=RC[-1]*2", "=RC[-1]*3", first)
    assert memo.get("=RC[-1]*2", "=RC[-1]*3") == first
    assert len(memo) == 1

    second = FormulaPairAnalysis(
        wrapper_kind="wrapped",
        wrapper_exact=True,
        event_key="formula-wrapper:wrapped:abc",
    )
    # A second distinct key exceeds the cap of 1: never cached, but the
    # already-cached first key is unaffected.
    memo.put("=RC[-2]*2", "=RC[-2]*3", second)
    assert memo.get("=RC[-2]*2", "=RC[-2]*3") is None
    assert len(memo) == 1
    assert memo.get("=RC[-1]*2", "=RC[-1]*3") == first


def test_pair_analysis_memo_extrapolated_full_cap_stays_within_128_mib() -> None:
    """Empirical proof for Step 5's 128 MiB budget: measure the deep byte
    size of a representative sample filled into the memo, then extrapolate
    linearly to the full cap (filling the actual ~150K-entry cap is
    unnecessarily slow for a unit test, and homogeneous dict entries scale
    linearly).
    """
    import random
    import sys

    def _deep_size(obj: object, seen: set[int] | None = None) -> int:
        if seen is None:
            seen = set()
        if id(obj) in seen:
            return 0
        seen.add(id(obj))
        size = sys.getsizeof(obj)
        if isinstance(obj, dict):
            for key, value in obj.items():
                size += _deep_size(key, seen) + _deep_size(value, seen)
        elif isinstance(obj, (list, tuple, set, frozenset)):
            for item in obj:
                size += _deep_size(item, seen)
        return size

    rng = random.Random(0)
    sample_size = 5_000
    memo = FormulaPairAnalysisMemo(cap=sample_size)
    for i in range(sample_size):
        base_norm = "=" + "+".join(
            f"R[{rng.randint(-50, 50)}]C[{rng.randint(-20, 20)}]"
            for _ in range(rng.randint(1, 5))
        )
        curr_norm = base_norm + f"+R[{rng.randint(-50, 50)}]C"
        memo.put(
            base_norm,
            curr_norm,
            FormulaPairAnalysis(
                wrapper_kind=("wrapped" if i % 3 == 0 else None),
                wrapper_exact=(bool(i % 2) if i % 3 == 0 else None),
                event_key=f"formula-wrapper:wrapped:{i:012x}" if i % 3 == 0 else "",
            ),
        )

    assert len(memo) > sample_size * 0.9  # a rare RNG string collision is fine
    measured_bytes = _deep_size(memo._values)  # white-box internal measurement
    bytes_per_entry = measured_bytes / len(memo)
    projected_full_cap_bytes = bytes_per_entry * _FORMULA_PAIR_ANALYSIS_MEMO_CAP

    budget_bytes = 128 * 1024 * 1024
    assert projected_full_cap_bytes < budget_bytes, (
        f"projected {projected_full_cap_bytes / (1024 * 1024):.1f} MiB at the "
        f"full {_FORMULA_PAIR_ANALYSIS_MEMO_CAP}-entry cap exceeds the 128 MiB "
        f"budget (measured {bytes_per_entry:.1f} bytes/entry from a "
        f"{sample_size}-entry sample)"
    )


def test_pair_analysis_memo_improves_compare_time_on_a_repeated_pattern_workload() -> (
    None
):
    """Step 5's required benchmark: on a workload where the SAME canonical
    key repeats across many rows (unlike Step 3's own overhead benchmark,
    which deliberately used a UNIQUE multiplier per row to isolate
    telemetry-only cost), memoizing wrapper detection must measurably speed
    up `diff_workbook_formulas()`'s own wall time -- this function's cost is
    exactly what `run_qc()`'s `comparing_formulas` phase measures. Uses
    `time.process_time()` (CPU time, immune to OS-scheduling noise) and
    `min()` across repeats -- the same robust methodology Step 3 needed for
    a reliable result on this shared host.

    The 5% budget (not Step 5's original 30%) reflects a real, confirmed
    correctness fix found during plan-20260910 Step 8 guest validation: the
    `expected` (extension) flag and the ADDED_REFERENCE evidence tag are
    computed from RAW formula text and are NOT a pure function of the
    normalized cache key (a real-data conflict probe found exactly one
    canonical key whose occurrences disagreed on ADDED_REFERENCE), so both
    are now always recomputed fresh -- only wrapper detection (a genuine
    pure function of the key) is still memoized. A smaller, correct
    improvement is the right outcome; do not restore the 30% budget by
    caching the unsafe fields again. The margin below 10% (rather than
    right at it) absorbs this shared-host benchmark's own run-to-run CPU
    noise (measured 9.3%-16.3% across repeated local runs).
    """
    rows = list(range(20, 2_020))
    base_formulas: dict[int, str] = {}
    curr_formulas: dict[int, str] = {}
    for row in rows:
        core = f"SUM(A{row}:B{row})+C{row}*D{row}+E{row}*F{row}+G{row}*H{row}"
        distractor = (
            f"SUM(I{row}:J{row})+K{row}*L{row}+M{row}*N{row}+O{row}*P{row}"
        )
        expression = core
        for _ in range(10):
            expression = f"IF(FALSE,{distractor},{expression})"
        base_formulas[row] = f"={core}"
        curr_formulas[row] = f"={expression}"
    max_row = max(rows) + 1
    baseline = _formula_column_workbook(base_formulas, max_row=max_row)
    current = _formula_column_workbook(curr_formulas, max_row=max_row)
    alignment = _alignment_for_rows(rows, max_row)

    repeats = 9

    def run(*, memoized: bool) -> float:
        started = time.process_time()
        diff_workbook_formulas(
            baseline,
            current,
            alignment,
            pair_analysis_memo=FormulaPairAnalysisMemo() if memoized else None,
            _use_native_delta=False,
        )
        return time.process_time() - started

    run(memoized=False)  # untimed warm-up each side
    run(memoized=True)
    unmemoized_times: list[float] = []
    memoized_times: list[float] = []
    for _ in range(repeats):
        unmemoized_times.append(run(memoized=False))
        memoized_times.append(run(memoized=True))

    unmemoized_best = min(unmemoized_times)
    memoized_best = min(memoized_times)
    assert unmemoized_best > 0.05, (
        "fixture too small to measure reliably "
        f"(unmemoized best {unmemoized_best:.4f}s CPU) -- widen it, don't "
        "loosen the 5% budget"
    )
    improvement = (unmemoized_best - memoized_best) / unmemoized_best

    assert improvement >= 0.05, (
        f"memoized compare time improved only {improvement:.4%}, below the "
        f"required 5% (unmemoized best {unmemoized_best:.4f}s CPU, memoized "
        f"best {memoized_best:.4f}s CPU)"
    )
