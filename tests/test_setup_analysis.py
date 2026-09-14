"""Tests for the Step 6 pure setup-analysis orchestration
(``qc_tool.setup.analysis``).

Plan: docs/plans/plan-20260913-mode-aware-configuration-wizard.md, Step 6.
"""

from __future__ import annotations

import random
from unittest.mock import patch

from qc_tool.io.model import CellRecord, CellValue, SheetSnapshot, WorkbookSnapshot
from qc_tool.setup.analysis import analyze_member

_LARGE_N = 6000


def _sheet(
    name: str,
    rows: list[list[CellValue]],
    *,
    visibility: str = "visible",
    formulas: dict[tuple[int, int], str] | None = None,
) -> SheetSnapshot:
    cells: dict[tuple[int, int], CellRecord] = {}
    formulas = formulas or {}
    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row, start=1):
            if value is None:
                continue
            formula = formulas.get((row_index, col_index))
            cells[(row_index, col_index)] = CellRecord(
                row=row_index,
                column=col_index,
                value=value,
                formula=formula,
                is_formula=formula is not None,
            )
    max_row = len(rows)
    max_col = max((len(row) for row in rows), default=0)
    return SheetSnapshot(name, visibility, max_row, max_col, cells)


def _workbook(*sheets: SheetSnapshot, name: str = "source.xlsx") -> WorkbookSnapshot:
    return WorkbookSnapshot(
        source_name=name,
        file_format="xlsx",
        formulas_available=True,
        styles_available=True,
        sheets=list(sheets),
    )


def _ranked_rows(n: int) -> list[list[CellValue]]:
    return [[f"ID{i}", f"Name{i}", 100.0 + i] for i in range(n)]


def test_analyze_member_detects_a_simple_region_on_both_sides() -> None:
    baseline_sheet = _sheet("Data", [["Header"], ["A1"], ["A2"]])
    current_sheet = _sheet("Data", [["Header"], ["A1"], ["A2"], ["A3"]])
    profile = analyze_member(
        member_id="primary",
        baseline=_workbook(baseline_sheet),
        current=_workbook(current_sheet),
        baseline_hash="a" * 64,
        current_hash="b" * 64,
    )

    assert profile.member_id == "primary"
    assert profile.failure_detail == ""
    assert len(profile.baseline_sheets) == 1
    assert len(profile.current_sheets) == 1
    assert profile.baseline_sheets[0].sheet_name == "Data"
    assert len(profile.baseline_sheets[0].regions) == 1
    assert len(profile.current_sheets[0].regions) == 1
    assert profile.current_sheets[0].regions[0].region.max_row == 4


def test_analyze_member_reports_hidden_and_very_hidden_sheets() -> None:
    hidden_sheet = _sheet("Hidden", [["x"]], visibility="hidden")
    very_hidden_sheet = _sheet("VeryHidden", [["x"]], visibility="veryHidden")
    profile = analyze_member(
        member_id="primary",
        baseline=_workbook(hidden_sheet, very_hidden_sheet),
        current=_workbook(hidden_sheet, very_hidden_sheet),
        baseline_hash="a" * 64,
        current_hash="b" * 64,
    )

    by_name = {sheet.sheet_name: sheet for sheet in profile.current_sheets}
    assert by_name["Hidden"].hidden is True
    assert by_name["Hidden"].very_hidden is False
    assert by_name["VeryHidden"].hidden is False
    assert by_name["VeryHidden"].very_hidden is True


def test_analyze_member_ranked_candidate_populated_when_sides_align() -> None:
    n = _LARGE_N
    base_rows = _ranked_rows(n)
    curr_rows = list(base_rows)
    random.Random(1234).shuffle(curr_rows)
    profile = analyze_member(
        member_id="primary",
        baseline=_workbook(_sheet("Panel", base_rows)),
        current=_workbook(_sheet("Panel", curr_rows)),
        baseline_hash="a" * 64,
        current_hash="b" * 64,
    )

    current_region = profile.current_sheets[0].regions[0]
    assert current_region.ranked_candidate is not None
    assert current_region.ranked_candidate.column_letters == ("A",)


def test_analyze_member_ranked_candidate_not_evaluated_without_a_matching_sheet() -> None:
    n = _LARGE_N
    curr_rows = _ranked_rows(n)
    random.Random(1234).shuffle(curr_rows)
    profile = analyze_member(
        member_id="primary",
        baseline=_workbook(_sheet("OtherSheet", [["x"]])),
        current=_workbook(_sheet("Panel", curr_rows)),
        baseline_hash="a" * 64,
        current_hash="b" * 64,
    )

    current_region = profile.current_sheets[0].regions[0]
    assert current_region.ranked_candidate is None


def test_analyze_member_ranked_candidate_not_evaluated_when_region_counts_differ() -> None:
    n = _LARGE_N
    rows = _ranked_rows(n)
    # Two disjoint regions on baseline (a gap row splits them into two
    # flood-fill components) vs one contiguous region on current.
    baseline_rows = [*rows[: n // 2], [None, None, None], *rows[n // 2 :]]
    profile = analyze_member(
        member_id="primary",
        baseline=_workbook(_sheet("Panel", baseline_rows)),
        current=_workbook(_sheet("Panel", rows)),
        baseline_hash="a" * 64,
        current_hash="b" * 64,
    )

    current_region = profile.current_sheets[0].regions[0]
    assert current_region.ranked_candidate is None


def test_analyze_member_isolates_a_single_sheet_detector_failure() -> None:
    good_sheet = _sheet("Good", [["x"], ["y"]])
    bad_sheet = _sheet("Bad", [["x"]])

    def _boom(sheet, *args, **kwargs):
        if sheet.name == "Bad":
            raise ValueError("synthetic detector failure")
        return []

    with patch("qc_tool.setup.analysis.detect_regions", side_effect=_boom):
        profile = analyze_member(
            member_id="primary",
            baseline=_workbook(good_sheet, bad_sheet),
            current=_workbook(good_sheet, bad_sheet),
            baseline_hash="a" * 64,
            current_hash="b" * 64,
        )

    by_name = {sheet.sheet_name: sheet for sheet in profile.current_sheets}
    assert by_name["Good"].failure_detail == ""
    assert "synthetic detector failure" in by_name["Bad"].failure_detail
    assert by_name["Bad"].regions == ()


def test_analyze_member_computes_a_workload_forecast_for_both_sides() -> None:
    sheet = _sheet("Data", [[1, 2], [3, 4]], formulas={(1, 1): "=B1+1"})
    profile = analyze_member(
        member_id="primary",
        baseline=_workbook(sheet),
        current=_workbook(sheet),
        baseline_hash="a" * 64,
        current_hash="b" * 64,
    )

    assert profile.baseline_complexity is not None
    assert profile.baseline_complexity.formula_count == 1
    assert profile.current_complexity is not None
    assert profile.current_complexity.formula_count == 1


def test_analyze_member_workload_forecast_failure_disables_only_the_forecast() -> None:
    sheet = _sheet("Data", [["x"]])
    with patch(
        "qc_tool.setup.analysis.assess_workbook_complexity",
        side_effect=ValueError("synthetic failure"),
    ):
        profile = analyze_member(
            member_id="primary",
            baseline=_workbook(sheet),
            current=_workbook(sheet),
            baseline_hash="a" * 64,
            current_hash="b" * 64,
        )

    assert profile.baseline_complexity is None
    assert profile.current_complexity is None
    # The rest of the scan still completed.
    assert len(profile.current_sheets) == 1


def test_analyze_member_never_touches_the_source_hashes() -> None:
    sheet = _sheet("Data", [["x"]])
    profile = analyze_member(
        member_id="primary",
        baseline=_workbook(sheet),
        current=_workbook(sheet),
        baseline_hash="a" * 64,
        current_hash="b" * 64,
    )

    assert profile.baseline_hash == "a" * 64
    assert profile.current_hash == "b" * 64
