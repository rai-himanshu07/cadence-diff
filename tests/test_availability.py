"""General Excel and PowerPoint availability-boundary contracts."""

from pathlib import Path

import yaml
from openpyxl import Workbook

from qc_tool.availability import excel_blank_allowed
from qc_tool.config.lint import lint_profile
from qc_tool.config.profile import DeliverableProfile, PptProfile
from qc_tool.coverage import CoverageState, QCRunMode
from qc_tool.engine import run_qc
from qc_tool.excel.align import align_workbooks
from qc_tool.excel.formulas import diff_workbook_formulas
from qc_tool.excel.preflight import preflight_workbook
from qc_tool.findings import FindingClass, FindingExpectedReason
from qc_tool.io.model import CellRecord, SheetSnapshot, WorkbookSnapshot
from qc_tool.ppt.diff import diff_decks
from qc_tool.ppt.match import SlideMatching
from qc_tool.ppt.model import (
    ChartContent,
    DeckSnapshot,
    PptChartPlot,
    PptChartSeries,
    SlideContent,
    TableContent,
)
from qc_tool.ppt.preflight import preflight_deck

_PERIODS = ["Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26", "Jun-26", "Jul-26"]


def _excel_profile(
    *,
    allow_blank_after: bool = True,
    ignore_ranges: list[str] | None = None,
    refresh_ranges: list[str] | None = None,
) -> DeliverableProfile:
    return DeliverableProfile.model_validate(
        {
            "name": "availability",
            "excel": {
                "sheets": {
                    "Data": {
                        "ignore_ranges": ignore_ranges or [],
                        "refresh_ranges": refresh_ranges or [],
                        "availability_rules": [
                            {
                                "name": "actuals",
                                "range": "B2:H2",
                                "periods": "B1:H1",
                                "required_through": "Jun-26",
                                "allow_blank_after": allow_blank_after,
                                "role": "actual",
                            }
                        ],
                    }
                }
            },
        }
    )


def _formula_snapshot(missing_column: int) -> WorkbookSnapshot:
    cells = {(1, 1): CellRecord(1, 1, "Period"), (2, 1): CellRecord(2, 1, "Value")}
    for column, period in enumerate(_PERIODS, start=2):
        cells[(1, column)] = CellRecord(1, column, period)
        if column != missing_column:
            cells[(2, column)] = CellRecord(
                2,
                column,
                column * 10,
                formula=f"={column}",
            )
    return WorkbookSnapshot(
        "availability.xlsx",
        "xlsx",
        True,
        True,
        formula_presence_available=True,
        sheets=[SheetSnapshot("Data", "visible", 2, 8, cells)],
    )


def test_excel_formula_gap_allows_only_declared_future_blank() -> None:
    future = preflight_workbook(_formula_snapshot(8), _excel_profile())
    historical = preflight_workbook(_formula_snapshot(4), _excel_profile())

    assert not any(
        finding.finding_class is FindingClass.FORMULA_MISSING
        for finding in future.findings
    )
    historical_gaps = [
        finding
        for finding in historical.findings
        if finding.finding_class is FindingClass.FORMULA_MISSING
    ]
    assert [(finding.sheet, finding.location) for finding in historical_gaps] == [
        ("Data", "D2")
    ]


def test_excel_ignore_wins_and_role_does_not_imply_blank_permission() -> None:
    ignored = preflight_workbook(
        _formula_snapshot(4),
        _excel_profile(ignore_ranges=["D2"]),
    )
    required_future = preflight_workbook(
        _formula_snapshot(8),
        _excel_profile(allow_blank_after=False),
    )

    assert not any(
        finding.finding_class is FindingClass.FORMULA_MISSING
        for finding in ignored.findings
    )
    assert any(
        finding.finding_class is FindingClass.FORMULA_MISSING
        and finding.location == "H2"
        for finding in required_future.findings
    )


def test_excel_vertical_period_mapping_supports_long_tables() -> None:
    cells: dict[tuple[int, int], CellRecord] = {}
    for row, period in enumerate(_PERIODS, start=2):
        cells[(row, 1)] = CellRecord(row, 1, period)
    sheet = SheetSnapshot("Data", "visible", 8, 5, cells)
    profile = DeliverableProfile.model_validate(
        {
            "name": "long",
            "excel": {
                "sheets": {
                    "Data": {
                        "availability_rules": [
                            {
                                "range": "E2:E8",
                                "periods": "A2:A8",
                                "required_through": "Jun-26",
                            }
                        ]
                    }
                }
            },
        }
    ).sheet_profile("Data")

    assert not excel_blank_allowed(sheet, profile, 7, 5)
    assert excel_blank_allowed(sheet, profile, 8, 5)


def test_cycle_formula_growth_and_ignore_follow_availability_precedence() -> None:
    baseline = _formula_snapshot(99)
    baseline.sheets[0].max_column = 7
    baseline.sheets[0].cells = {
        key: cell for key, cell in baseline.sheets[0].cells.items() if key[1] <= 7
    }
    current = _formula_snapshot(8)
    profile = _excel_profile()
    findings = diff_workbook_formulas(
        baseline,
        current,
        align_workbooks(baseline, current, profile),
        profile,
    )
    assert not any(
        finding.finding_class is FindingClass.FORMULA_NOT_EXTENDED
        and finding.location == "H2"
        for finding in findings
    )

    ignored_current = _formula_snapshot(4)
    ignored_profile = _excel_profile(ignore_ranges=["D2"])
    ignored_findings = diff_workbook_formulas(
        _formula_snapshot(99),
        ignored_current,
        align_workbooks(_formula_snapshot(99), ignored_current, ignored_profile),
        ignored_profile,
    )
    assert not any(
        finding.location == "D2"
        and finding.finding_class
        in {FindingClass.FORMULA_REMOVED, FindingClass.FORMULA_MISSING}
        for finding in ignored_findings
    )


def _save_value_workbook(path: Path, future_value: int | None) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet.append(["Metric", "Jan-26", "Feb-26", "Mar-26", "Apr-26"])
    sheet.append(["Revenue", 10, 20, 30, future_value])
    workbook.save(path)


def _value_profile(
    *,
    refresh: bool = False,
    allow_blank_after: bool = True,
) -> DeliverableProfile:
    return DeliverableProfile.model_validate(
        {
            "name": "values",
            "excel": {
                "sheets": {
                    "Data": {
                        "refresh_ranges": ["E2"] if refresh else [],
                        "availability_rules": [
                            {
                                "range": "B2:E2",
                                "periods": "B1:E1",
                                "required_through": "Mar-26",
                                "allow_blank_after": allow_blank_after,
                            }
                        ],
                    }
                }
            },
        }
    )


def test_availability_controls_blankness_not_nonblank_refresh(
    tmp_path: Path,
) -> None:
    baseline = tmp_path / "baseline.xlsx"
    blank = tmp_path / "blank.xlsx"
    changed = tmp_path / "changed.xlsx"
    _save_value_workbook(baseline, 40)
    _save_value_workbook(blank, None)
    _save_value_workbook(changed, 50)

    blank_result = run_qc(
        baseline_excel=baseline,
        current_excel=blank,
        profile=_value_profile(),
    )
    changed_result = run_qc(
        baseline_excel=baseline,
        current_excel=changed,
        profile=_value_profile(),
    )
    refreshed_result = run_qc(
        baseline_excel=baseline,
        current_excel=changed,
        profile=_value_profile(refresh=True),
    )

    assert not any(
        finding.finding_class is FindingClass.VALUE_CHANGED
        and finding.location == "E2"
        for finding in blank_result.findings
    )
    changed_finding = next(
        finding
        for finding in changed_result.findings
        if finding.finding_class is FindingClass.VALUE_CHANGED
        and finding.location == "E2"
    )
    refreshed_finding = next(
        finding
        for finding in refreshed_result.findings
        if finding.finding_class is FindingClass.VALUE_CHANGED
        and finding.location == "E2"
    )
    assert not changed_finding.expected_growth
    assert refreshed_finding.expected_growth
    assert refreshed_finding.expected_reason is FindingExpectedReason.PROFILE_REFRESH


def test_refresh_range_never_excuses_a_required_blank(tmp_path: Path) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _save_value_workbook(baseline, 40)
    _save_value_workbook(current, None)

    result = run_qc(
        baseline_excel=baseline,
        current_excel=current,
        profile=_value_profile(refresh=True, allow_blank_after=False),
    )
    finding = next(
        finding
        for finding in result.findings
        if finding.finding_class is FindingClass.VALUE_CHANGED
        and finding.location == "E2"
    )

    assert not finding.expected_growth
    assert "required value is blank" in finding.message


def _ppt_profile(*, allow_blank_after: bool = True) -> PptProfile:
    return PptProfile.model_validate(
        {
            "availability_rules": [
                {
                    "slide": "Dashboard",
                    "scope": "table",
                    "series": "Revenue",
                    "required_through": "Mar-26",
                    "allow_blank_after": allow_blank_after,
                    "role": "actual",
                },
                {
                    "slide": "Dashboard",
                    "scope": "chart",
                    "element": "Trend",
                    "series": "Revenue",
                    "required_through": "Mar-26",
                    "allow_blank_after": allow_blank_after,
                    "role": "forecast",
                },
            ]
        }
    )


def _ppt_slide(values: list[float | None]) -> SlideContent:
    categories = ["Jan-26", "Feb-26", "Mar-26", "Apr-26"]
    series = PptChartSeries(
        index=0,
        order=0,
        name="Revenue",
        categories=categories,
        values=values,
        plot_index=0,
        source_index=0,
        source_id="series-0",
    )
    plot = PptChartPlot(
        index=0,
        chart_type="lineChart",
        series=[series],
        source_id="plot-0",
    )
    chart = ChartContent(
        chart_type="lineChart",
        categories=categories,
        series=[("Revenue", values)],
        source_id="chart-0",
        title="Trend",
        plots=[plot],
    )
    table = TableContent(
        rows=[
            ["Metric", *categories],
            ["Revenue", "10", "20", "", ""],
        ],
        source_id="table-0",
    )
    return SlideContent(
        index=0,
        title="Dashboard",
        texts=[],
        tables=[table],
        charts=[chart],
        shape_count=2,
    )


def test_ppt_preflight_allows_future_blanks_but_flags_historical_blanks() -> None:
    deck = DeckSnapshot("deck.pptx", slides=[_ppt_slide([10, 20, None, None])])

    result = preflight_deck(deck, _ppt_profile())

    table_blanks = [
        finding
        for finding in result.findings
        if finding.finding_class is FindingClass.PPT_TABLE_BLANK
    ]
    chart_blanks = [
        finding
        for finding in result.findings
        if finding.finding_class is FindingClass.PPT_CHART_VALUE_MISSING
    ]
    assert len(table_blanks) == 1 and "column 4" in (table_blanks[0].element or "")
    assert len(chart_blanks) == 1 and chart_blanks[0].element == "Mar-26"


def test_ppt_role_does_not_allow_future_blank_without_explicit_permission() -> None:
    deck = DeckSnapshot("deck.pptx", slides=[_ppt_slide([10, 20, 30, None])])

    result = preflight_deck(deck, _ppt_profile(allow_blank_after=False))

    assert any(
        finding.finding_class is FindingClass.PPT_CHART_VALUE_MISSING
        and finding.element == "Apr-26"
        for finding in result.findings
    )


def test_ppt_truncated_series_identifies_historical_missing_period() -> None:
    deck = DeckSnapshot("deck.pptx", slides=[_ppt_slide([10, 20])])

    result = preflight_deck(deck, _ppt_profile())

    assert any(
        finding.finding_class is FindingClass.PPT_CHART_LENGTH_MISMATCH
        for finding in result.findings
    )
    missing = [
        finding
        for finding in result.findings
        if finding.finding_class is FindingClass.PPT_CHART_VALUE_MISSING
    ]
    assert [finding.element for finding in missing] == ["Mar-26"]


def test_ppt_cycle_diff_suppresses_only_future_blank() -> None:
    baseline = _ppt_slide([10, 20, 30, 40])
    current = _ppt_slide([10, 20, None, None])

    findings = diff_decks(
        SlideMatching(pairs=[(baseline, current)]),
        _ppt_profile(),
    )
    value_findings = [
        finding
        for finding in findings
        if finding.finding_class is FindingClass.CHART_VALUE_CHANGED
    ]

    assert len(value_findings) == 1
    assert value_findings[0].element == "Mar-26"


def test_availability_profile_roundtrip_and_lint_validation() -> None:
    profile = _excel_profile()
    serialized = yaml.safe_dump(
        profile.model_dump(mode="json", by_alias=True),
        sort_keys=False,
    )
    reloaded = DeliverableProfile.model_validate(yaml.safe_load(serialized))
    assert reloaded == profile

    invalid = DeliverableProfile.model_validate(
        {
            "name": "invalid",
            "excel": {
                "sheets": {
                    "Data": {
                        "availability_rules": [
                            {
                                "range": "B2:H3",
                                "periods": "B1:D1",
                                "required_through": "not-a-period",
                            }
                        ]
                    }
                }
            },
            "ppt": {
                "availability_rules": [
                    {
                        "slide": "Dashboard",
                        "scope": "chart",
                        "required_through": "also-invalid",
                    }
                ]
            },
        }
    )
    issues = lint_profile(invalid)
    messages = [issue.message for issue in issues]
    assert any("required_through" in message for message in messages)
    assert any("period range width" in message for message in messages)


def test_lint_detects_period_kind_mismatch_against_live_workbook() -> None:
    profile = DeliverableProfile.model_validate(
        {
            "name": "kind-mismatch",
            "excel": {
                "sheets": {
                    "Data": {
                        "availability_rules": [
                            {
                                "range": "B2:H2",
                                "periods": "B1:H1",
                                "required_through": "Q2-26",
                            }
                        ]
                    }
                }
            },
        }
    )
    issues = lint_profile(profile, workbook=_formula_snapshot(99))

    assert any("do not match required_through kind" in issue.message for issue in issues)


def test_overlapping_rules_use_the_stricter_blank_decision() -> None:
    workbook = _formula_snapshot(8)
    profile = DeliverableProfile.model_validate(
        {
            "name": "overlap",
            "excel": {
                "sheets": {
                    "Data": {
                        "availability_rules": [
                            {
                                "range": "B2:H2",
                                "periods": "B1:H1",
                                "required_through": "Jun-26",
                            },
                            {
                                "range": "H2",
                                "periods": "H1",
                                "required_through": "Jul-26",
                            },
                        ]
                    }
                }
            },
        }
    ).sheet_profile("Data")

    assert not excel_blank_allowed(workbook.sheet("Data"), profile, 2, 8)


def test_single_period_cell_can_control_a_target_range() -> None:
    sheet = SheetSnapshot(
        "Data",
        "visible",
        2,
        5,
        {(1, 1): CellRecord(1, 1, "Jan-26")},
    )
    profile = DeliverableProfile.model_validate(
        {
            "name": "point-period",
            "excel": {
                "sheets": {
                    "Data": {
                        "availability_rules": [
                            {
                                "range": "B2:E2",
                                "periods": "A1",
                                "required_through": "Dec-25",
                            }
                        ]
                    }
                }
            },
        }
    ).sheet_profile("Data")

    assert excel_blank_allowed(sheet, profile, 2, 5)


def test_ppt_lint_resolves_untitled_slide_and_checks_element_series() -> None:
    slide = _ppt_slide([10, 20, 30, 40])
    slide.title = None
    slide.index = 2
    deck = DeckSnapshot("deck.pptx", slides=[slide])
    valid = PptProfile.model_validate(
        {
            "availability_rules": [
                {
                    "slide": "slide 3",
                    "scope": "chart",
                    "element": "Trend",
                    "series": "Revenue",
                    "required_through": "Mar-26",
                }
            ]
        }
    )
    assert lint_profile(DeliverableProfile(name="valid", ppt=valid), deck=deck) == []

    invalid = PptProfile.model_validate(
        {
            "availability_rules": [
                {
                    "slide": "slide 3",
                    "scope": "chart",
                    "element": "Missing",
                    "series": "Ghost",
                    "required_through": "Mar-26",
                }
            ]
        }
    )
    issues = lint_profile(DeliverableProfile(name="invalid", ppt=invalid), deck=deck)
    assert any("element 'Missing' not found" in issue.message for issue in issues)


def test_nonfunctional_rules_degrade_runtime_coverage() -> None:
    invalid_excel = DeliverableProfile.model_validate(
        {
            "name": "invalid-excel",
            "excel": {
                "sheets": {
                    "Data": {
                        "availability_rules": [
                            {
                                "range": "B2:H2",
                                "periods": "B1:D1",
                                "required_through": "Jun-26",
                            }
                        ]
                    }
                }
            },
        }
    )
    excel_result = preflight_workbook(_formula_snapshot(99), invalid_excel)
    excel_coverage = next(
        item
        for item in excel_result.coverage
        if item.check_id == "excel-availability"
    )
    assert excel_coverage.state is CoverageState.DEGRADED
    assert "misalign" in excel_coverage.detail

    invalid_ppt = PptProfile.model_validate(
        {
            "availability_rules": [
                {
                    "slide": "Dashboard",
                    "scope": "chart",
                    "element": "Missing",
                    "series": "Revenue",
                    "required_through": "Mar-26",
                }
            ]
        }
    )
    deck = DeckSnapshot("deck.pptx", slides=[_ppt_slide([10, 20, 30, 40])])
    ppt_result = preflight_deck(deck, invalid_ppt)
    ppt_coverage = next(
        item for item in ppt_result.coverage if item.check_id == "ppt-availability"
    )
    assert ppt_coverage.state is CoverageState.DEGRADED
    assert "element 'Missing' not found" in ppt_coverage.detail


def test_availability_coverage_discloses_strict_default(fixture_dir: Path) -> None:
    result = run_qc(
        current_ppt=fixture_dir / "current.pptx",
        mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
    )
    coverage = next(
        item for item in result.coverage if item.check_id == "ppt-availability"
    )
    assert coverage.state is CoverageState.CHECKED
    assert "strict" in coverage.detail.lower()

    excel_result = run_qc(
        current_excel=fixture_dir / "current.xlsx",
        mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
    )
    excel_coverage = next(
        item
        for item in excel_result.coverage
        if item.check_id == "excel-availability"
    )
    assert excel_coverage.state is CoverageState.CHECKED
    assert "strict" in excel_coverage.detail.lower()
