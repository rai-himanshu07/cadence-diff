"""Pre-diff ranked-table suspicion detector for unconfigured positional block
regions.

Screens a region that already fell back to positional row alignment for a
composite identity that would explain its rows more reliably than raw
position. Never applied automatically: this module only scores candidates.
Row correspondence changes only after the analyst confirms a matching
``RowIdentityRule`` in the profile and re-runs QC (see ``qc_tool.excel.align``
for the confirmed-rule alignment path).
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, replace
from itertools import combinations, pairwise

from openpyxl.utils import get_column_letter

from qc_tool.excel.regions import TableRegion
from qc_tool.io.model import CellValue, SheetSnapshot

#: Confirmation thresholds and search caps (Step 6 architecture; pinned).
#: Changing any of these requires new synthetic counterexamples and a plan
#: amendment -- never tune them against a private file.
MIN_DATA_ROWS = 25
MIN_NON_BLANK_COVERAGE = 0.95
MIN_UNIQUE_RATIO = 0.995
MIN_KEY_OVERLAP = 0.90
MAX_FORMULA_RATIO = 0.05
MIN_DISPLACED_RATIO = 0.20
MIN_MISMATCH_REDUCTION = 0.50
MIN_PROJECTED_MISMATCHES = 10_000
MIN_PROJECTED_AVOIDED_MISMATCHES = 10_000
MAX_SINGLE_CANDIDATES = 12
MAX_COMPOSITE_SIZE = 3
_SAMPLE_ROWS = 500
_HEADER_SCAN_ROWS = 32
_HEADER_SCAN_COLUMNS = 64
_HEADER_LOOKAHEAD_ROWS = 24


@dataclass(frozen=True, slots=True)
class RankedTableCandidate:
    """One bounded, evidence-scored composite-key suggestion for one region."""

    columns: tuple[int, ...]  # 1-based column indices, in the region
    non_blank_coverage: float
    unique_ratio: float
    key_overlap: float
    formula_ratio: float
    displaced_ratio: float
    mismatch_reduction: float
    projected_positional_mismatches: int
    ordinal_columns: tuple[int, ...] = ()
    header_row: int | None = None

    @property
    def column_letters(self) -> tuple[str, ...]:
        return tuple(get_column_letter(column) for column in self.columns)

    @property
    def ordinal_column_letters(self) -> tuple[str, ...]:
        return tuple(get_column_letter(column) for column in self.ordinal_columns)

    @property
    def projected_avoided_mismatches(self) -> int:
        return round(self.projected_positional_mismatches * self.mismatch_reduction)


def _is_blank(value: CellValue) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _key_component(value: CellValue) -> object:
    """Normalize a cell value into a hashable, comparison-stable component."""
    if isinstance(value, str):
        return value.strip().casefold()
    return value


def _cell_value(sheet: SheetSnapshot, row: int, column: int) -> CellValue:
    cell = sheet.cells.get((row, column))
    return None if cell is None else cell.value


def _column_values(
    sheet: SheetSnapshot, region: TableRegion, column: int
) -> list[CellValue]:
    return [
        _cell_value(sheet, row, column)
        for row in range(region.min_row, region.max_row + 1)
    ]


def _column_formula_ratio(sheet: SheetSnapshot, region: TableRegion, column: int) -> float:
    populated = [
        sheet.cells[(row, column)]
        for row in range(region.min_row, region.max_row + 1)
        if (row, column) in sheet.cells
    ]
    if not populated:
        return 0.0
    return sum(1 for cell in populated if cell.has_formula) / len(populated)


def _is_sequence_like(values: list[CellValue]) -> bool:
    """Whether non-blank numeric values form a simple arithmetic sequence.

    A plain row-rank or sequence column (1, 2, 3, ... or any constant step)
    is exactly the kind of column whose "identity" would just reproduce
    positional order, so it is never a safe candidate.
    """
    numeric = [
        value
        for value in values
        if isinstance(value, int | float) and not isinstance(value, bool)
    ]
    if len(numeric) < 3:
        return False
    steps = {round(b - a, 9) for a, b in pairwise(numeric)}
    return len(steps) == 1 and next(iter(steps)) != 0


@dataclass(frozen=True, slots=True)
class _ColumnScreen:
    column: int
    non_blank_base: float
    non_blank_curr: float
    unique_base: float
    unique_curr: float
    formula_ratio: float
    sequence_like: bool

    @property
    def safe(self) -> bool:
        return (
            not self.sequence_like
            and self.formula_ratio <= MAX_FORMULA_RATIO
            and self.non_blank_base >= MIN_NON_BLANK_COVERAGE
            and self.non_blank_curr >= MIN_NON_BLANK_COVERAGE
        )

    @property
    def score(self) -> float:
        return min(self.unique_base, self.unique_curr)


def _screen_column(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
    column: int,
) -> _ColumnScreen:
    base_values = _column_values(base_sheet, base_region, column)
    curr_values = _column_values(curr_sheet, curr_region, column)
    base_non_blank = [v for v in base_values if not _is_blank(v)]
    curr_non_blank = [v for v in curr_values if not _is_blank(v)]
    base_keys = [_key_component(v) for v in base_non_blank]
    curr_keys = [_key_component(v) for v in curr_non_blank]
    return _ColumnScreen(
        column=column,
        non_blank_base=len(base_non_blank) / len(base_values) if base_values else 0.0,
        non_blank_curr=len(curr_non_blank) / len(curr_values) if curr_values else 0.0,
        unique_base=(len(set(base_keys)) / len(base_keys)) if base_keys else 0.0,
        unique_curr=(len(set(curr_keys)) / len(curr_keys)) if curr_keys else 0.0,
        formula_ratio=max(
            _column_formula_ratio(base_sheet, base_region, column),
            _column_formula_ratio(curr_sheet, curr_region, column),
        ),
        sequence_like=_is_sequence_like(base_values) or _is_sequence_like(curr_values),
    )


def _row_keys(
    sheet: SheetSnapshot, region: TableRegion, columns: tuple[int, ...]
) -> dict[int, tuple[object, ...] | None]:
    """Row -> composite key, or None if any component is blank."""
    keys: dict[int, tuple[object, ...] | None] = {}
    for row in range(region.min_row, region.max_row + 1):
        parts: list[object] = []
        blank = False
        for column in columns:
            value = _cell_value(sheet, row, column)
            if _is_blank(value):
                blank = True
                break
            parts.append(_key_component(value))
        keys[row] = None if blank else tuple(parts)
    return keys


def _infer_header_row(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
    identity_columns: tuple[int, ...],
    ordinal_columns: tuple[int, ...],
) -> int | None:
    """Find a stable text header immediately before displaced data.

    Ranked tables can have titles or metadata above their actual header. The
    scan is deliberately bounded and cross-cycle: a candidate header must be
    text-rich and stable in both files, while following identity rows must show
    the displacement that caused this detector to fire. Unclear cases return
    ``None`` and preserve the existing first-row behavior.
    """
    base_height = base_region.max_row - base_region.min_row + 1
    curr_height = curr_region.max_row - curr_region.min_row + 1
    scan_rows = min(_HEADER_SCAN_ROWS, base_height, curr_height)
    display_columns = range(
        curr_region.min_col,
        min(curr_region.max_col + 1, curr_region.min_col + _HEADER_SCAN_COLUMNS),
    )
    minimum_text_cells = min(2, curr_region.max_col - curr_region.min_col + 1)
    if not ordinal_columns:
        return None
    candidates: list[tuple[int, float, float, int]] = []
    for offset in range(scan_rows):
        base_row = base_region.min_row + offset
        curr_row = curr_region.min_row + offset
        comparable = 0
        stable = 0
        stable_text = 0
        for column in display_columns:
            base_value = _cell_value(base_sheet, base_row, column)
            curr_value = _cell_value(curr_sheet, curr_row, column)
            if _is_blank(base_value) and _is_blank(curr_value):
                continue
            comparable += 1
            if _key_component(base_value) != _key_component(curr_value):
                continue
            stable += 1
            if isinstance(curr_value, str) and curr_value.strip():
                stable_text += 1
        if (
            comparable == 0
            or stable_text < minimum_text_cells
            or stable / comparable < 0.75
        ):
            continue
        ordinal_boundary = False
        for column in ordinal_columns:
            header_value = _cell_value(curr_sheet, curr_row, column)
            following_values = [
                _cell_value(curr_sheet, row, column)
                for row in range(curr_row + 1, min(curr_region.max_row, curr_row + 4) + 1)
            ]
            if (
                (isinstance(header_value, str) or _is_blank(header_value))
                and len(following_values) >= 3
                and _is_sequence_like(following_values)
            ):
                ordinal_boundary = True
                break
        if not ordinal_boundary:
            continue

        following = 0
        displaced = 0
        for next_offset in range(
            offset + 1,
            min(scan_rows, offset + 1 + _HEADER_LOOKAHEAD_ROWS),
        ):
            next_base_row = base_region.min_row + next_offset
            next_curr_row = curr_region.min_row + next_offset
            base_values = tuple(
                _cell_value(base_sheet, next_base_row, column)
                for column in identity_columns
            )
            curr_values = tuple(
                _cell_value(curr_sheet, next_curr_row, column)
                for column in identity_columns
            )
            if any(_is_blank(value) for value in (*base_values, *curr_values)):
                continue
            base_key = tuple(_key_component(value) for value in base_values)
            curr_key = tuple(_key_component(value) for value in curr_values)
            following += 1
            displaced += base_key != curr_key
        if following < 3 or displaced / following < MIN_DISPLACED_RATIO:
            continue
        candidates.append(
            (stable_text, stable / comparable, displaced / following, offset)
        )
    if not candidates:
        return None
    return curr_region.min_row + max(candidates)[3]


def _sample_mismatch_reduction(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
    identity_columns: tuple[int, ...],
    base_row_by_key: dict[tuple[object, ...], int],
    curr_row_by_key: dict[tuple[object, ...], int],
    matched_keys: set[tuple[object, ...]],
) -> tuple[float, int]:
    """Deterministic bounded sample comparing positional vs. key-based pairing."""
    value_columns = [
        column
        for column in range(curr_region.min_col, curr_region.max_col + 1)
        if column not in identity_columns
        and not (
            _is_sequence_like(_column_values(base_sheet, base_region, column))
            and _is_sequence_like(_column_values(curr_sheet, curr_region, column))
        )
    ]
    if not value_columns or not matched_keys:
        return 0.0, 0
    ordered_keys = sorted(matched_keys, key=repr)
    sample_size = min(len(ordered_keys), _SAMPLE_ROWS)
    if sample_size < len(ordered_keys):
        step = len(ordered_keys) / sample_size
        sampled = [ordered_keys[int(index * step)] for index in range(sample_size)]
    else:
        sampled = ordered_keys

    common_height = min(
        base_region.max_row - base_region.min_row + 1,
        curr_region.max_row - curr_region.min_row + 1,
    )
    positional_mismatches = 0
    key_mismatches = 0
    for key in sampled:
        base_row = base_row_by_key[key]
        curr_row = curr_row_by_key[key]
        offset = base_row - base_region.min_row
        positional_curr_row = (
            curr_region.min_row + offset if offset < common_height else None
        )
        for column in value_columns:
            base_value = _cell_value(base_sheet, base_row, column)
            key_curr_value = _cell_value(curr_sheet, curr_row, column)
            if base_value != key_curr_value:
                key_mismatches += 1
            if positional_curr_row is not None:
                positional_value = _cell_value(curr_sheet, positional_curr_row, column)
                if base_value != positional_value:
                    positional_mismatches += 1
    if positional_mismatches == 0:
        return 0.0, 0
    reduction = max(0.0, 1 - (key_mismatches / positional_mismatches))
    scale = len(matched_keys) / sample_size if sample_size else 1.0
    projected = round(positional_mismatches * scale)
    return reduction, projected


def _evaluate_combination(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
    columns: tuple[int, ...],
    formula_ratio: float,
) -> RankedTableCandidate | None:
    base_keys = _row_keys(base_sheet, base_region, columns)
    curr_keys = _row_keys(curr_sheet, curr_region, columns)
    base_non_blank = {row: key for row, key in base_keys.items() if key is not None}
    curr_non_blank = {row: key for row, key in curr_keys.items() if key is not None}
    if not base_keys or not curr_keys:
        return None
    non_blank_coverage = min(
        len(base_non_blank) / len(base_keys), len(curr_non_blank) / len(curr_keys)
    )
    if non_blank_coverage < MIN_NON_BLANK_COVERAGE:
        return None

    base_key_counts = Counter(base_non_blank.values())
    curr_key_counts = Counter(curr_non_blank.values())
    unique_base = sum(1 for count in base_key_counts.values() if count == 1) / len(
        base_non_blank
    )
    unique_curr = sum(1 for count in curr_key_counts.values() if count == 1) / len(
        curr_non_blank
    )
    unique_ratio = min(unique_base, unique_curr)
    if unique_ratio < MIN_UNIQUE_RATIO:
        return None

    base_unique_keys = {key for key, count in base_key_counts.items() if count == 1}
    curr_unique_keys = {key for key, count in curr_key_counts.items() if count == 1}
    union = base_unique_keys | curr_unique_keys
    key_overlap = len(base_unique_keys & curr_unique_keys) / len(union) if union else 0.0
    if key_overlap < MIN_KEY_OVERLAP:
        return None

    base_row_by_key = {
        key: row for row, key in base_non_blank.items() if base_key_counts[key] == 1
    }
    curr_row_by_key = {
        key: row for row, key in curr_non_blank.items() if curr_key_counts[key] == 1
    }
    matched = base_unique_keys & curr_unique_keys
    if not matched:
        return None
    displaced = sum(
        1
        for key in matched
        if (base_row_by_key[key] - base_region.min_row)
        != (curr_row_by_key[key] - curr_region.min_row)
    )
    displaced_ratio = displaced / len(matched)
    if displaced_ratio < MIN_DISPLACED_RATIO:
        return None

    mismatch_reduction, projected_mismatches = _sample_mismatch_reduction(
        base_sheet,
        curr_sheet,
        base_region,
        curr_region,
        columns,
        base_row_by_key,
        curr_row_by_key,
        matched,
    )
    if projected_mismatches < MIN_PROJECTED_MISMATCHES:
        return None
    if (
        mismatch_reduction < MIN_MISMATCH_REDUCTION
        and round(projected_mismatches * mismatch_reduction)
        < MIN_PROJECTED_AVOIDED_MISMATCHES
    ):
        return None

    return RankedTableCandidate(
        columns=columns,
        non_blank_coverage=non_blank_coverage,
        unique_ratio=unique_ratio,
        key_overlap=key_overlap,
        formula_ratio=formula_ratio,
        displaced_ratio=displaced_ratio,
        mismatch_reduction=mismatch_reduction,
        projected_positional_mismatches=projected_mismatches,
    )


def detect_ranked_table_candidate(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
) -> RankedTableCandidate | None:
    """Screen one unconfigured positional block region for a composite identity.

    Returns the single best-scoring candidate that clears every threshold, or
    ``None`` if nothing qualifies -- current behavior continues unchanged in
    that case, and no control is ever shown for the region.
    """
    row_count = curr_region.max_row - curr_region.min_row + 1
    if row_count < MIN_DATA_ROWS:
        return None
    columns = range(curr_region.min_col, curr_region.max_col + 1)
    screens = [
        _screen_column(base_sheet, curr_sheet, base_region, curr_region, column)
        for column in columns
    ]
    safe = sorted((s for s in screens if s.safe), key=lambda s: -s.score)
    top = safe[:MAX_SINGLE_CANDIDATES]
    if not top:
        return None
    top_columns = [screen.column for screen in top]
    formula_ratio_by_column = {screen.column: screen.formula_ratio for screen in top}

    best: RankedTableCandidate | None = None
    for size in range(1, min(MAX_COMPOSITE_SIZE, len(top_columns)) + 1):
        for combo in combinations(top_columns, size):
            candidate = _evaluate_combination(
                base_sheet,
                curr_sheet,
                base_region,
                curr_region,
                combo,
                max(formula_ratio_by_column[column] for column in combo),
            )
            if candidate is None:
                continue
            if best is None or candidate.unique_ratio > best.unique_ratio:
                best = candidate
    if best is None:
        return None
    ordinal_columns = tuple(
        screen.column
        for screen in screens
        if _is_sequence_like(_column_values(base_sheet, base_region, screen.column))
        and _is_sequence_like(_column_values(curr_sheet, curr_region, screen.column))
    )
    return replace(
        best,
        ordinal_columns=ordinal_columns,
        header_row=_infer_header_row(
            base_sheet,
            curr_sheet,
            base_region,
            curr_region,
            best.columns,
            ordinal_columns,
        ),
    )
