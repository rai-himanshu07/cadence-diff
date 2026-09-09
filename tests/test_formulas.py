"""Formula QC tests (acceptance criterion 3) against the fixture manifest."""

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
    FormulaComparisonTelemetry,
    _differs_only_by_extension,
    diff_workbook_formulas,
    formula_text_comparable,
    formula_token_diff,
    to_r1c1,
)
from qc_tool.excel.regions import TableRegion
from qc_tool.findings import Finding, FindingClass
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
