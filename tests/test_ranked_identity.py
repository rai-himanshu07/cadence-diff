"""Step 6: pre-diff ranked-table suspicion detector (Acceptance Criterion 8).

Synthetic, hand-built regions -- no private data. Every threshold pinned in
`qc_tool.excel.ranked_identity` is exercised directly against a case designed
to sit clearly on one side of it, so a future threshold change is caught here
rather than silently drifting.
"""

from __future__ import annotations

import random

from qc_tool.excel.ranked_identity import (
    MIN_DATA_ROWS,
    MIN_MISMATCH_REDUCTION,
    MIN_PROJECTED_MISMATCHES,
    detect_ranked_table_candidate,
)
from qc_tool.excel.regions import TableRegion
from qc_tool.io.model import CellRecord, CellValue, SheetSnapshot

# Large enough that a fully displaced permutation clears
# MIN_PROJECTED_MISMATCHES (10,000) with just a couple of value columns.
# The detector is deliberately silent on small regions -- the negative cases
# below stay small because they are blocked by an earlier, N-independent
# gate instead (see test_below_minimum_row_count_never_suggests and friends).
_LARGE_N = 6000


def _sheet(
    rows: list[list[CellValue]], *, formulas: dict[tuple[int, int], str] | None = None
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
    return SheetSnapshot("Panel", "visible", max_row, max_col, cells)


def _region(rows: int, cols: int) -> TableRegion:
    return TableRegion("Panel", 1, 1, rows, cols, "block", None, 1, "none")


def _ranked_rows(n: int) -> list[list[CellValue]]:
    # A non-sequential unique text id: a genuine identity column must not
    # itself look like a plain row rank (see the dedicated rank-column test
    # below for that specific, deliberately numeric-sequence case).
    return [[f"ID{i}", f"Name{i}", 100.0 + i] for i in range(n)]


def test_pure_permutation_is_detected() -> None:
    n = _LARGE_N
    base_rows = _ranked_rows(n)
    curr_rows = list(base_rows)
    random.Random(1234).shuffle(curr_rows)
    base_sheet = _sheet(base_rows)
    curr_sheet = _sheet(curr_rows)
    region = _region(n, 3)

    candidate = detect_ranked_table_candidate(base_sheet, curr_sheet, region, region)

    assert candidate is not None
    assert candidate.column_letters == ("A",)
    assert candidate.displaced_ratio >= 0.20
    assert candidate.mismatch_reduction >= 0.50
    assert candidate.projected_positional_mismatches >= MIN_PROJECTED_MISMATCHES


def test_unshuffled_identical_order_is_not_suggested() -> None:
    n = 30
    rows = _ranked_rows(n)
    sheet = _sheet(rows)
    region = _region(n, 3)

    # Positional alignment is already correct; nothing displaced to fix.
    assert detect_ranked_table_candidate(sheet, sheet, region, region) is None


def test_below_minimum_row_count_never_suggests() -> None:
    n = MIN_DATA_ROWS - 1
    base_rows = _ranked_rows(n)
    curr_rows = list(reversed(base_rows))
    base_sheet = _sheet(base_rows)
    curr_sheet = _sheet(curr_rows)
    region = _region(n, 3)

    assert detect_ranked_table_candidate(base_sheet, curr_sheet, region, region) is None


def test_sequence_like_rank_column_is_excluded_in_favor_of_real_identity() -> None:
    n = _LARGE_N
    # Column A is a plain row rank; column B is the real stable identity.
    # Columns C/D/E are extra real content columns tied to identity (not to
    # position) -- a rank column is positionally self-consistent by
    # definition (row i's rank is always i), which would otherwise cancel
    # out the mismatch-reduction signal if it were the only value column.
    base_rows = [
        [i + 1, f"ID{i}", 100.0 + i, 200.0 + i * 2, 300.0 - i] for i in range(n)
    ]
    curr_rows = [list(row) for row in base_rows]
    random.Random(99).shuffle(curr_rows)
    # A rank column always reflects the CURRENT sort, never the original row.
    for position, row in enumerate(curr_rows):
        row[0] = position + 1
    base_sheet = _sheet(base_rows)
    curr_sheet = _sheet(curr_rows)
    region = _region(n, 5)

    candidate = detect_ranked_table_candidate(base_sheet, curr_sheet, region, region)

    assert candidate is not None
    assert candidate.column_letters == ("B",)  # the rank column is never chosen


def test_sequence_like_rank_does_not_cancel_minimal_table_evidence() -> None:
    n = MIN_PROJECTED_MISMATCHES + 100
    base_rows = [[i + 1, f"ID{i}", 100.0 + i] for i in range(n)]
    curr_rows = [list(row) for row in base_rows]
    random.Random(19).shuffle(curr_rows)
    for position, row in enumerate(curr_rows):
        row[0] = position + 1
    base_sheet = _sheet(base_rows)
    curr_sheet = _sheet(curr_rows)
    region = _region(n, 3)

    candidate = detect_ranked_table_candidate(base_sheet, curr_sheet, region, region)

    assert candidate is not None
    assert candidate.column_letters == ("B",)
    assert candidate.mismatch_reduction == 1.0
    assert candidate.projected_positional_mismatches >= MIN_PROJECTED_MISMATCHES


def test_duplicate_single_column_needs_a_composite_key() -> None:
    n = _LARGE_N
    half = n // 2
    # Column A repeats (two teams); column B *also* repeats, since the member
    # index is reused across both teams -- only the pair together is unique.
    # Columns C/D/E are extra arithmetic-sequence value columns purely so the
    # shuffled region has enough cells to clear the projected-mismatch floor.
    base_rows = [
        [
            f"Team {'A' if i < half else 'B'}",
            f"Member{i % half}",
            100.0 + i,
            200.0 + i * 2,
            300.0 - i,
        ]
        for i in range(n)
    ]
    curr_rows = [list(row) for row in base_rows]
    random.Random(7).shuffle(curr_rows)
    base_sheet = _sheet(base_rows)
    curr_sheet = _sheet(curr_rows)
    region = _region(n, 5)

    candidate = detect_ranked_table_candidate(base_sheet, curr_sheet, region, region)

    assert candidate is not None
    # Column order reflects a screening heuristic, not semantics -- both
    # columns are required together, in either order.
    assert set(candidate.columns) == {1, 2}


def test_large_absolute_noise_reduction_survives_genuine_changes() -> None:
    n = 12_000
    base_rows = [
        [i + 1, f"ID{i}", 100.0 + i, 200.0 + i, 300.0 + i, 400.0 + i]
        for i in range(n)
    ]
    curr_rows = [list(row) for row in base_rows]
    random.Random(23).shuffle(curr_rows)
    for position, row in enumerate(curr_rows):
        row[0] = position + 1
        # Three genuine changed measures remain after identity matching. The
        # fourth measure proves a large absolute population of positional noise.
        row[3] += 1.0
        row[4] += 1.0
        row[5] += 1.0
    base_sheet = _sheet(base_rows)
    curr_sheet = _sheet(curr_rows)
    region = _region(n, 6)

    candidate = detect_ranked_table_candidate(base_sheet, curr_sheet, region, region)

    assert candidate is not None
    assert candidate.column_letters == ("B",)
    assert candidate.mismatch_reduction < MIN_MISMATCH_REDUCTION
    assert candidate.projected_avoided_mismatches >= 10_000


def test_formula_derived_column_is_excluded() -> None:
    n = _LARGE_N
    # Column A would otherwise be a fine unique key, but it is formula-derived.
    base_rows = [[f"F{i}", f"ID{i}", 100.0 + i] for i in range(n)]
    curr_rows = [list(row) for row in base_rows]
    random.Random(55).shuffle(curr_rows)
    formulas = {(row, 1): "=ROW()" for row in range(1, n + 1)}
    base_sheet = _sheet(base_rows, formulas=formulas)
    curr_sheet = _sheet(curr_rows, formulas=formulas)
    region = _region(n, 3)

    candidate = detect_ranked_table_candidate(base_sheet, curr_sheet, region, region)

    assert candidate is not None
    assert candidate.column_letters == ("B",)


def test_low_key_overlap_across_cycles_is_not_suggested() -> None:
    n = 30
    # Fewer than 90% of current keys existed in baseline: mostly new rows,
    # not a reordering of the same population.
    base_rows = [[f"ID{i}", f"Name{i}", 100.0 + i] for i in range(n)]
    curr_rows = [[f"ID{i + 20}", f"Name{i + 20}", 200.0 + i] for i in range(n)]
    base_sheet = _sheet(base_rows)
    curr_sheet = _sheet(curr_rows)
    region = _region(n, 3)

    assert detect_ranked_table_candidate(base_sheet, curr_sheet, region, region) is None


def test_low_non_blank_coverage_column_is_excluded() -> None:
    n = 30
    half = n // 2
    # Same "needs both columns" shape as the composite-key test above, so
    # excluding column A for low coverage genuinely removes the only viable
    # identity -- column B alone still repeats across both teams.
    base_rows = [
        [f"Team {'A' if i < half else 'B'}", f"Member{i % half}", 100.0 + i]
        for i in range(n)
    ]
    curr_rows = [list(row) for row in base_rows]
    random.Random(3).shuffle(curr_rows)
    # Blank out more than 5% of column A -- below the 95% coverage floor.
    for row in base_rows[:5]:
        row[0] = None
    base_sheet = _sheet(base_rows)
    curr_sheet = _sheet(curr_rows)
    region = _region(n, 3)

    assert detect_ranked_table_candidate(base_sheet, curr_sheet, region, region) is None
