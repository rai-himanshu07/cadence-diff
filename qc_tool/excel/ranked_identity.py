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
MANUAL_MIN_NON_BLANK_COVERAGE = 0.25
MANUAL_MIN_PROJECTED_AVOIDED_MISMATCHES = 5_000
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
    manual_review: bool = False

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
    base_sequence_like: bool
    curr_sequence_like: bool
    base_components: tuple[int | None, ...]
    curr_components: tuple[int | None, ...]
    component_values: tuple[object, ...]

    @property
    def sequence_like(self) -> bool:
        return self.base_sequence_like or self.curr_sequence_like

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

    @property
    def manual_safe(self) -> bool:
        return (
            not self.sequence_like
            and self.non_blank_base >= MANUAL_MIN_NON_BLANK_COVERAGE
            and self.non_blank_curr >= MANUAL_MIN_NON_BLANK_COVERAGE
        )


def _screen_column(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
    column: int,
) -> _ColumnScreen:
    base_values = _column_values(base_sheet, base_region, column)
    curr_values = _column_values(curr_sheet, curr_region, column)
    component_ids: dict[object, int] = {}
    component_values: list[object] = []

    def encode(values: list[CellValue]) -> tuple[int | None, ...]:
        encoded: list[int | None] = []
        for value in values:
            if _is_blank(value):
                encoded.append(None)
                continue
            component = _key_component(value)
            try:
                component_id = component_ids[component]
            except KeyError:
                component_id = len(component_ids)
                component_ids[component] = component_id
                component_values.append(component)
            encoded.append(component_id)
        return tuple(encoded)

    base_components = encode(base_values)
    curr_components = encode(curr_values)
    base_keys = [value for value in base_components if value is not None]
    curr_keys = [value for value in curr_components if value is not None]
    return _ColumnScreen(
        column=column,
        non_blank_base=len(base_keys) / len(base_values) if base_values else 0.0,
        non_blank_curr=len(curr_keys) / len(curr_values) if curr_values else 0.0,
        unique_base=(len(set(base_keys)) / len(base_keys)) if base_keys else 0.0,
        unique_curr=(len(set(curr_keys)) / len(curr_keys)) if curr_keys else 0.0,
        formula_ratio=max(
            _column_formula_ratio(base_sheet, base_region, column),
            _column_formula_ratio(curr_sheet, curr_region, column),
        ),
        base_sequence_like=_is_sequence_like(base_values),
        curr_sequence_like=_is_sequence_like(curr_values),
        base_components=base_components,
        curr_components=curr_components,
        component_values=tuple(component_values),
    )


def _cannot_contribute_displacement(screen: _ColumnScreen) -> bool:
    base_offsets: dict[object, int] = {}
    for offset, component in enumerate(screen.base_components):
        if component is None:
            continue
        if component in base_offsets:
            return False
        base_offsets[component] = offset
    curr_offsets: dict[object, int] = {}
    for offset, component in enumerate(screen.curr_components):
        if component is None:
            continue
        if component in curr_offsets:
            return False
        curr_offsets[component] = offset
    return all(
        base_offsets[component] == curr_offsets[component]
        for component in base_offsets.keys() & curr_offsets.keys()
    )


def _combination_key_stats(
    region: TableRegion,
    columns: tuple[int, ...],
    screens: dict[int, _ColumnScreen],
    *,
    baseline: bool,
) -> tuple[
    int,
    Counter[tuple[int, ...]],
    dict[tuple[int, ...], int],
]:
    vectors = [
        screens[column].base_components
        if baseline
        else screens[column].curr_components
        for column in columns
    ]
    counts: Counter[tuple[int, ...]] = Counter()
    row_by_key: dict[tuple[int, ...], int] = {}
    non_blank = 0
    for offset, row in enumerate(range(region.min_row, region.max_row + 1)):
        parts: list[int] = []
        for vector in vectors:
            component = vector[offset]
            if component is None:
                break
            parts.append(component)
        else:
            key = tuple(parts)
            non_blank += 1
            counts[key] += 1
            row_by_key[key] = row
    return non_blank, counts, row_by_key


def _infer_header_row(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
    identity_columns: tuple[int, ...],
    ordinal_columns: tuple[int, ...],
    *,
    manual_review: bool = False,
) -> int | None:
    """Find a stable header immediately before sustained identity data.

    Ranked tables can have titles or metadata above the actual header. The
    scan is bounded and cross-cycle. A row is accepted when an ordinal column
    changes from text to a numeric sequence, or when every proposed identity
    column has the same text label in both files and the following rows show a
    clear increase in populated key data.
    """
    base_height = base_region.max_row - base_region.min_row + 1
    curr_height = curr_region.max_row - curr_region.min_row + 1
    scan_rows = min(_HEADER_SCAN_ROWS, base_height, curr_height)
    display_columns = range(
        curr_region.min_col,
        min(curr_region.max_col + 1, curr_region.min_col + _HEADER_SCAN_COLUMNS),
    )
    minimum_text_cells = min(2, curr_region.max_col - curr_region.min_col + 1)
    candidates: list[tuple[int, float, int, int, float, int]] = []
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
        text_rich = not (
            comparable == 0
            or stable_text < minimum_text_cells
            or stable / comparable < 0.75
        )

        ordinal_boundary = False
        if text_rich:
            for column in ordinal_columns:
                header_value = _cell_value(curr_sheet, curr_row, column)
                following_values = [
                    _cell_value(curr_sheet, row, column)
                    for row in range(
                        curr_row + 1,
                        min(curr_region.max_row, curr_row + 4) + 1,
                    )
                ]
                if (
                    (isinstance(header_value, str) or _is_blank(header_value))
                    and len(following_values) >= 3
                    and _is_sequence_like(following_values)
                ):
                    ordinal_boundary = True
                    break

        stable_identity_labels = 0
        for column in identity_columns:
            base_cell = base_sheet.cells.get((base_row, column))
            curr_cell = curr_sheet.cells.get((curr_row, column))
            base_label = _cell_value(base_sheet, base_row, column)
            curr_label = _cell_value(curr_sheet, curr_row, column)
            if not (
                isinstance(base_label, str)
                and base_label.strip()
                and isinstance(curr_label, str)
                and curr_label.strip()
                and base_cell is not None
                and not base_cell.has_formula
                and curr_cell is not None
                and not curr_cell.has_formula
                and _key_component(base_label) == _key_component(curr_label)
            ):
                break
            stable_identity_labels += 1
        identity_labels_are_literal = bool(identity_columns) and (
            stable_identity_labels == len(identity_columns)
        )
        ordinal_boundary = ordinal_boundary and identity_labels_are_literal

        identity_boundary = False
        identity_gain = 0.0
        if offset > 0 and identity_labels_are_literal:
            next_base_rows = range(
                base_row + 1, min(base_region.max_row, base_row + 8) + 1
            )
            next_curr_rows = range(
                curr_row + 1, min(curr_region.max_row, curr_row + 8) + 1
            )
            prev_base_rows = range(max(base_region.min_row, base_row - 3), base_row)
            prev_curr_rows = range(max(curr_region.min_row, curr_row - 3), curr_row)

            def key_coverage(sheet: SheetSnapshot, rows: range) -> float:
                values = [
                    _cell_value(sheet, row, column)
                    for row in rows
                    for column in identity_columns
                ]
                return (
                    sum(not _is_blank(item) for item in values) / len(values)
                    if values
                    else 0.0
                )

            next_coverage = min(
                key_coverage(base_sheet, next_base_rows),
                key_coverage(curr_sheet, next_curr_rows),
            )
            previous_coverage = max(
                key_coverage(base_sheet, prev_base_rows),
                key_coverage(curr_sheet, prev_curr_rows),
            )
            identity_gain = next_coverage - previous_coverage
            minimum_next_coverage = 0.35 if manual_review else 0.75
            minimum_coverage_gain = 0.10 if manual_review else 0.25
            identity_boundary = (
                next_coverage >= minimum_next_coverage
                and identity_gain >= minimum_coverage_gain
            )

        if not ordinal_boundary and not identity_boundary:
            continue
        candidates.append(
            (
                int(ordinal_boundary),
                identity_gain,
                stable_identity_labels,
                stable_text,
                stable / comparable if comparable else 0.0,
                -offset,
            )
        )
    if not candidates:
        return None
    return curr_region.min_row - max(candidates)[5]


def _sample_mismatch_reduction(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
    identity_columns: tuple[int, ...],
    base_row_by_key: dict[tuple[int, ...], int],
    curr_row_by_key: dict[tuple[int, ...], int],
    matched_keys: set[tuple[int, ...]],
    sequence_like_columns: frozenset[int],
    screens: dict[int, _ColumnScreen],
) -> tuple[float, int]:
    """Deterministic bounded sample comparing positional vs. key-based pairing."""
    value_columns = [
        column
        for column in range(curr_region.min_col, curr_region.max_col + 1)
        if column not in identity_columns
        and column not in sequence_like_columns
    ]
    if not value_columns or not matched_keys:
        return 0.0, 0
    ordered_keys = sorted(
        matched_keys,
        key=lambda key: repr(
            tuple(
                screens[column].component_values[component]
                for column, component in zip(identity_columns, key, strict=True)
            )
        ),
    )
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
    *,
    screens: dict[int, _ColumnScreen],
    sequence_like_columns: frozenset[int],
    manual_review: bool = False,
) -> RankedTableCandidate | None:
    base_non_blank, base_key_counts, base_rows = _combination_key_stats(
        base_region,
        columns,
        screens,
        baseline=True,
    )
    curr_non_blank, curr_key_counts, curr_rows = _combination_key_stats(
        curr_region,
        columns,
        screens,
        baseline=False,
    )
    base_row_count = base_region.max_row - base_region.min_row + 1
    curr_row_count = curr_region.max_row - curr_region.min_row + 1
    if base_row_count < 1 or curr_row_count < 1:
        return None
    non_blank_coverage = min(
        base_non_blank / base_row_count,
        curr_non_blank / curr_row_count,
    )
    minimum_coverage = (
        MANUAL_MIN_NON_BLANK_COVERAGE
        if manual_review
        else MIN_NON_BLANK_COVERAGE
    )
    if non_blank_coverage < minimum_coverage:
        return None

    unique_base = (
        sum(1 for count in base_key_counts.values() if count == 1) / base_non_blank
    )
    unique_curr = (
        sum(1 for count in curr_key_counts.values() if count == 1) / curr_non_blank
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
        key: base_rows[key] for key in base_unique_keys
    }
    curr_row_by_key = {
        key: curr_rows[key] for key in curr_unique_keys
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
        sequence_like_columns,
        screens,
    )
    if projected_mismatches < MIN_PROJECTED_MISMATCHES:
        return None
    projected_avoided = round(projected_mismatches * mismatch_reduction)
    if manual_review:
        if projected_avoided < MANUAL_MIN_PROJECTED_AVOIDED_MISMATCHES:
            return None
    elif (
        mismatch_reduction < MIN_MISMATCH_REDUCTION
        and projected_avoided < MIN_PROJECTED_AVOIDED_MISMATCHES
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
        manual_review=manual_review,
    )


def detect_ranked_table_candidate(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
    *,
    allow_manual_review: bool = False,
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
    safe = sorted(
        (
            screen
            for screen in screens
            if (
                screen.safe or (allow_manual_review and screen.manual_safe)
            )
            and not _cannot_contribute_displacement(screen)
        ),
        key=lambda screen: -screen.score,
    )
    top = safe[:MAX_SINGLE_CANDIDATES]
    if not top:
        return None
    if all(
        len(screen.base_components) == len(screen.curr_components)
        and all(
            baseline == current
            for baseline, current in zip(
                screen.base_components,
                screen.curr_components,
                strict=True,
            )
        )
        for screen in top
    ):
        return None
    top_columns = [screen.column for screen in top]
    formula_ratio_by_column = {screen.column: screen.formula_ratio for screen in top}
    screens_by_column = {screen.column: screen for screen in screens}
    sequence_like_columns = frozenset(
        screen.column
        for screen in screens
        if screen.base_sequence_like and screen.curr_sequence_like
    )

    best: RankedTableCandidate | None = None
    perfect_candidate = False
    for size in range(1, min(MAX_COMPOSITE_SIZE, len(top_columns)) + 1):
        for combo in combinations(top_columns, size):
            candidate = _evaluate_combination(
                base_sheet,
                curr_sheet,
                base_region,
                curr_region,
                combo,
                max(formula_ratio_by_column[column] for column in combo),
                screens=screens_by_column,
                sequence_like_columns=sequence_like_columns,
                manual_review=allow_manual_review,
            )
            if candidate is None:
                continue
            if best is None or candidate.unique_ratio > best.unique_ratio:
                best = candidate
            if best.unique_ratio == 1.0:
                perfect_candidate = True
                break
        if perfect_candidate:
            break
    if best is None:
        return None
    ordinal_columns = tuple(
        screen.column
        for screen in screens
        if screen.base_sequence_like and screen.curr_sequence_like
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
            manual_review=best.manual_review,
        ),
    )
