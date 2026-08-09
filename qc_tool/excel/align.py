"""Alignment engine: cell-level correspondence between baseline and current.

Per paired region, each axis (rows, columns) is aligned by identity keys:

- long regions:  rows keyed by leading label/period columns, columns by
  header text;
- wide regions:  columns keyed by header labels (periods), rows by the
  label column;
- block regions: rows keyed by label-column text, columns positionally.

Unmatched current entries whose period sorts *after* the last baseline
period are classified as expected cadence growth; any other unmatched
current entry is an unexpected insertion, and unmatched baseline entries
are deletions of historical data. If key matching pairs less than half of
the baseline axis, the axis falls back to positional alignment rather
than producing unreliable correspondences.
"""

import logging
import re
from collections import Counter
from collections.abc import Iterator
from dataclasses import dataclass, field
from typing import Literal

from pydantic import BaseModel, Field

from qc_tool.config.profile import DeliverableProfile
from qc_tool.excel.periods import Period, is_period_after, is_period_label, parse_period
from qc_tool.excel.regions import TableRegion, detect_regions
from qc_tool.io.model import SheetSnapshot, WorkbookSnapshot
from qc_tool.progress import CancellationToken, check_cancelled

logger = logging.getLogger(__name__)

_MIN_KEY_MATCH_RATIO = 0.5
_MIN_CONFIDENCE_AXIS_SIZE = 4
_MIN_BLOCK_LABEL_OVERLAP = 0.8

AxisKey = tuple[object, ...]


@dataclass(frozen=True, slots=True)
class AxisEntry:
    index: int  # absolute row or column index in its sheet
    key: AxisKey
    periods: tuple[Period | None, ...]  # period parse per key component


@dataclass(slots=True)
class AxisAlignment:
    pairs: list[tuple[int, int]] = field(default_factory=list)
    deleted: list[int] = field(default_factory=list)  # baseline indices
    inserted: list[int] = field(default_factory=list)  # current indices, unexpected
    growth: list[int] = field(default_factory=list)  # current indices, expected
    method: Literal["keys", "positional"] = "keys"
    low_confidence_fallback: bool = False


@dataclass(slots=True)
class RegionAlignment:
    baseline: TableRegion
    current: TableRegion
    rows: AxisAlignment
    columns: AxisAlignment

    def cell_pairs(self) -> Iterator[tuple[tuple[int, int], tuple[int, int]]]:
        """Yield ((baseline_row, baseline_col), (current_row, current_col))."""
        for base_row, curr_row in self.rows.pairs:
            for base_col, curr_col in self.columns.pairs:
                yield ((base_row, base_col), (curr_row, curr_col))

    @property
    def row_growth_count(self) -> int:
        return len(self.rows.growth)

    @property
    def column_growth_count(self) -> int:
        return len(self.columns.growth)

    @property
    def low_confidence(self) -> bool:
        return (
            self.rows.low_confidence_fallback
            or self.columns.low_confidence_fallback
        )


@dataclass(slots=True)
class WorkbookAlignment:
    common_sheets: list[str] = field(default_factory=list)
    added_sheets: list[str] = field(default_factory=list)
    removed_sheets: list[str] = field(default_factory=list)
    regions: dict[str, list[RegionAlignment]] = field(default_factory=dict)
    unpaired_baseline_regions: list[TableRegion] = field(default_factory=list)
    unpaired_current_regions: list[TableRegion] = field(default_factory=list)
    low_confidence_regions: list[str] = field(default_factory=list)


# --- alignment trust manifest ---------------------------------------------


class AlignmentAxisTrust(BaseModel):
    method: Literal["keys", "positional"]
    low_confidence_fallback: bool
    paired: int = Field(ge=0)
    deleted: int = Field(ge=0)
    inserted: int = Field(ge=0)
    growth: int = Field(ge=0)

    model_config = {"frozen": True}


class AlignmentRegionTrust(BaseModel):
    version: Literal[1] = 1
    artifact_member: str = Field(default="primary", min_length=1)
    sheet: str = Field(min_length=1)
    region_id: str = Field(min_length=1)
    baseline_range: str = Field(min_length=1)
    current_range: str = Field(min_length=1)
    row: AlignmentAxisTrust
    column: AlignmentAxisTrust
    comparable_cell_pairs: int = Field(ge=0)
    skipped_low_confidence_cells: int = Field(ge=0)
    low_confidence: bool

    model_config = {"frozen": True}


class AlignmentUnpairedRegion(BaseModel):
    version: Literal[1] = 1
    artifact_member: str = Field(default="primary", min_length=1)
    side: Literal["baseline", "current"]
    sheet: str = Field(min_length=1)
    region_id: str = Field(min_length=1)
    cell_range: str = Field(min_length=1)
    orientation: Literal["long", "wide", "block"]

    model_config = {"frozen": True}


class AlignmentTrustManifest(BaseModel):
    version: Literal[1] = 1
    regions: tuple[AlignmentRegionTrust, ...] = ()
    unpaired: tuple[AlignmentUnpairedRegion, ...] = ()

    model_config = {"frozen": True}


def build_alignment_trust_manifest(
    alignment: WorkbookAlignment,
    artifact_member: str = "primary",
) -> AlignmentTrustManifest:
    """Build factual, deterministic region correspondence metadata."""
    regions: list[AlignmentRegionTrust] = []
    for sheet_name in sorted(alignment.regions.keys()):
        region_list = alignment.regions[sheet_name]
        for region in region_list:
            # counts for rows/columns
            row_pairs = len(region.rows.pairs)
            col_pairs = len(region.columns.pairs)
            row_trust = AlignmentAxisTrust(
                method=region.rows.method,
                low_confidence_fallback=bool(region.rows.low_confidence_fallback),
                paired=row_pairs,
                deleted=len(region.rows.deleted),
                inserted=len(region.rows.inserted),
                growth=len(region.rows.growth),
            )
            col_trust = AlignmentAxisTrust(
                method=region.columns.method,
                low_confidence_fallback=bool(region.columns.low_confidence_fallback),
                paired=col_pairs,
                deleted=len(region.columns.deleted),
                inserted=len(region.columns.inserted),
                growth=len(region.columns.growth),
            )
            comparable_cell_pairs = row_pairs * col_pairs
            skipped = comparable_cell_pairs if region.low_confidence else 0
            regions.append(
                AlignmentRegionTrust(
                    version=1,
                    artifact_member=artifact_member,
                    sheet=sheet_name,
                    region_id=region.current.region_id,
                    baseline_range=region.baseline.cell_range,
                    current_range=region.current.cell_range,
                    row=row_trust,
                    column=col_trust,
                    comparable_cell_pairs=comparable_cell_pairs,
                    skipped_low_confidence_cells=skipped,
                    low_confidence=bool(region.low_confidence),
                )
            )

    unpaired: list[AlignmentUnpairedRegion] = []
    def unpaired_key(region: TableRegion) -> tuple[str, str, str]:
        return region.sheet, region.region_id, region.cell_range
    for reg in sorted(alignment.unpaired_baseline_regions, key=unpaired_key):
        unpaired.append(
            AlignmentUnpairedRegion(
                version=1,
                artifact_member=artifact_member,
                side="baseline",
                sheet=reg.sheet,
                region_id=reg.region_id,
                cell_range=reg.cell_range,
                orientation=reg.orientation,
            )
        )
    for reg in sorted(alignment.unpaired_current_regions, key=unpaired_key):
        unpaired.append(
            AlignmentUnpairedRegion(
                version=1,
                artifact_member=artifact_member,
                side="current",
                sheet=reg.sheet,
                region_id=reg.region_id,
                cell_range=reg.cell_range,
                orientation=reg.orientation,
            )
        )

    return AlignmentTrustManifest(regions=tuple(regions), unpaired=tuple(unpaired))



# --- axis key extraction --------------------------------------------------


def _column_is_labelish(sheet: SheetSnapshot, col: int, rows: range) -> bool:
    values = [
        sheet.cells[(row, col)].value for row in rows if (row, col) in sheet.cells
    ]
    if not values:
        return False
    non_numeric = sum(1 for v in values if not isinstance(v, int | float | bool))
    return non_numeric / len(values) >= 0.8


def _column_is_volatile_derived(sheet: SheetSnapshot, col: int, rows: range) -> bool:
    """Predominantly formula-derived AND period/date-valued: a display value
    (e.g. a derived "data through" date) that moves every cycle, never row
    identity."""
    cells = [sheet.cells[(row, col)] for row in rows if (row, col) in sheet.cells]
    if not cells:
        return False
    formulas = sum(1 for cell in cells if cell.has_formula)
    periods = sum(1 for cell in cells if parse_period(cell.value) is not None)
    return formulas / len(cells) >= 0.5 and periods / len(cells) >= 0.5


def _long_key_columns(sheet: SheetSnapshot, region: TableRegion) -> list[int]:
    """Leading label/period columns of a long region (identity columns).

    The first label column anchors identity even when formula-derived;
    additional components must be stable, so volatile derived period columns
    never churn the composite key.
    """
    data_rows = range((region.header_row or region.min_row) + 1, region.max_row + 1)
    key_cols: list[int] = []
    for col in range(region.min_col, region.max_col + 1):
        if not _column_is_labelish(sheet, col, data_rows):
            break
        if key_cols and _column_is_volatile_derived(sheet, col, data_rows):
            break
        key_cols.append(col)
    return key_cols or [region.key_col or region.min_col]


def _row_entries(
    sheet: SheetSnapshot, region: TableRegion, key_cols: list[int], first_row: int
) -> list[AxisEntry]:
    entries = []
    seen: Counter[AxisKey] = Counter()
    for row in range(first_row, region.max_row + 1):
        raw = tuple(
            sheet.cells[(row, col)].value if (row, col) in sheet.cells else None
            for col in key_cols
        )
        occurrence = seen[raw]
        seen[raw] += 1
        entries.append(
            AxisEntry(
                index=row,
                key=(*raw, occurrence),
                periods=tuple(parse_period(v) for v in raw),
            )
        )
    return entries


def _col_entries(
    sheet: SheetSnapshot, region: TableRegion, header_row: int, first_col: int
) -> list[AxisEntry]:
    entries = []
    seen: Counter[AxisKey] = Counter()
    for col in range(first_col, region.max_col + 1):
        value = (
            sheet.cells[(header_row, col)].value
            if (header_row, col) in sheet.cells
            else None
        )
        raw = (value,)
        occurrence = seen[raw]
        seen[raw] += 1
        entries.append(
            AxisEntry(index=col, key=(*raw, occurrence), periods=(parse_period(value),))
        )
    return entries


def _positional_entries(indices: range) -> list[AxisEntry]:
    return [
        AxisEntry(index=index, key=(offset,), periods=(None,))
        for offset, index in enumerate(indices)
    ]


# --- axis alignment -------------------------------------------------------


def _max_periods(entries: list[AxisEntry]) -> dict[tuple[int, str], Period]:
    """Per key-component position and cadence kind, the latest baseline period."""
    result: dict[tuple[int, str], Period] = {}
    for entry in entries:
        for position, period in enumerate(entry.periods):
            if period is None:
                continue
            key = (position, period.kind)
            best = result.get(key)
            if best is None or period.sort_key > best.sort_key:
                result[key] = period
    return result


def _is_growth(entry: AxisEntry, max_periods: dict[tuple[int, str], Period]) -> bool:
    for position, period in enumerate(entry.periods):
        if period is None:
            continue
        baseline_max = max_periods.get((position, period.kind))
        if baseline_max is not None and is_period_after(period, baseline_max):
            return True
    return False


def _align_axis(
    baseline: list[AxisEntry],
    current: list[AxisEntry],
    *,
    method: Literal["keys", "positional"] = "keys",
) -> AxisAlignment:
    current_by_key = {entry.key: entry for entry in current}
    matched_current: set[int] = set()
    alignment = AxisAlignment(method=method)

    for base_entry in baseline:
        match = current_by_key.get(base_entry.key)
        if match is not None:
            alignment.pairs.append((base_entry.index, match.index))
            matched_current.add(match.index)
        else:
            alignment.deleted.append(base_entry.index)

    max_periods = _max_periods(baseline)
    for curr_entry in current:
        if curr_entry.index in matched_current:
            continue
        if _is_growth(curr_entry, max_periods):
            alignment.growth.append(curr_entry.index)
        else:
            alignment.inserted.append(curr_entry.index)

    if baseline and len(alignment.pairs) / len(baseline) < _MIN_KEY_MATCH_RATIO:
        return _align_positionally(
            baseline,
            current,
            low_confidence=(
                len(baseline) >= _MIN_CONFIDENCE_AXIS_SIZE
                and len(current) >= _MIN_CONFIDENCE_AXIS_SIZE
            ),
        )
    return alignment


def _align_positionally(
    baseline: list[AxisEntry],
    current: list[AxisEntry],
    *,
    low_confidence: bool = False,
) -> AxisAlignment:
    alignment = AxisAlignment(
        method="positional",
        low_confidence_fallback=low_confidence,
    )
    shared = min(len(baseline), len(current))
    alignment.pairs = [
        (baseline[i].index, current[i].index) for i in range(shared)
    ]
    alignment.deleted = [entry.index for entry in baseline[shared:]]
    max_periods = _max_periods(baseline)
    for entry in current[shared:]:
        if _is_growth(entry, max_periods):
            alignment.growth.append(entry.index)
        else:
            alignment.inserted.append(entry.index)
    return alignment


# --- region alignment -----------------------------------------------------


def _align_long(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
) -> RegionAlignment:
    base_keys = _long_key_columns(base_sheet, base_region)
    curr_keys = _long_key_columns(curr_sheet, curr_region)
    base_header = base_region.header_row or base_region.min_row
    curr_header = curr_region.header_row or curr_region.min_row
    rows = _align_axis(
        _row_entries(base_sheet, base_region, base_keys, base_header + 1),
        _row_entries(curr_sheet, curr_region, curr_keys, curr_header + 1),
    )
    columns = _align_axis(
        _col_entries(base_sheet, base_region, base_header, base_region.min_col),
        _col_entries(curr_sheet, curr_region, curr_header, curr_region.min_col),
    )
    return RegionAlignment(base_region, curr_region, rows, columns)


def _align_wide(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
) -> RegionAlignment:
    base_header = base_region.header_row or base_region.min_row
    curr_header = curr_region.header_row or curr_region.min_row
    label_col_base = [base_region.key_col or base_region.min_col]
    label_col_curr = [curr_region.key_col or curr_region.min_col]
    rows = _align_axis(
        _row_entries(base_sheet, base_region, label_col_base, base_header),
        _row_entries(curr_sheet, curr_region, label_col_curr, curr_header),
    )
    columns = _align_axis(
        _col_entries(base_sheet, base_region, base_header, base_region.min_col),
        _col_entries(curr_sheet, curr_region, curr_header, curr_region.min_col),
    )
    return RegionAlignment(base_region, curr_region, rows, columns)


def _block_stable_labels(
    sheet: SheetSnapshot, region: TableRegion
) -> tuple[str, ...]:
    col = region.key_col or region.min_col
    labels: list[str] = []
    for row in range(region.min_row, region.max_row + 1):
        cell = sheet.cells.get((row, col))
        if (
            cell is not None
            and isinstance(cell.value, str)
            and cell.value.strip()
            and not cell.has_formula
            and parse_period(cell.value) is None
        ):
            labels.append(cell.value)
    return tuple(labels)


def _block_labels_are_stable(sheet: SheetSnapshot, region: TableRegion) -> bool:
    """Block rows may key-align only on constant text labels. Formula-derived
    or period-valued "labels" are display values that move every cycle; using
    them as identity turns KPI refreshes into phantom row events."""
    row_count = region.max_row - region.min_row + 1
    labels = _block_stable_labels(sheet, region)
    if not labels or row_count <= 0:
        return False
    return len(labels) / row_count >= 0.6 and len(set(labels)) == len(labels)


def _block_labels_overlap(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
) -> bool:
    baseline = set(_block_stable_labels(base_sheet, base_region))
    current = set(_block_stable_labels(curr_sheet, curr_region))
    total = max(len(baseline), len(current))
    return bool(total) and len(baseline & current) / total >= _MIN_BLOCK_LABEL_OVERLAP


def _align_block(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
) -> RegionAlignment:
    label_base = [base_region.key_col or base_region.min_col]
    label_curr = [curr_region.key_col or curr_region.min_col]
    if (
        _block_labels_are_stable(base_sheet, base_region)
        and _block_labels_are_stable(curr_sheet, curr_region)
        and _block_labels_overlap(
            base_sheet,
            curr_sheet,
            base_region,
            curr_region,
        )
    ):
        rows = _align_axis(
            _row_entries(base_sheet, base_region, label_base, base_region.min_row),
            _row_entries(curr_sheet, curr_region, label_curr, curr_region.min_row),
        )
    else:
        rows = _align_axis(
            _positional_entries(range(base_region.min_row, base_region.max_row + 1)),
            _positional_entries(range(curr_region.min_row, curr_region.max_row + 1)),
            method="positional",
        )
    columns = _align_axis(
        _positional_entries(range(base_region.min_col, base_region.max_col + 1)),
        _positional_entries(range(curr_region.min_col, curr_region.max_col + 1)),
        method="positional",
    )
    return RegionAlignment(base_region, curr_region, rows, columns)


def align_regions(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_region: TableRegion,
    curr_region: TableRegion,
) -> RegionAlignment:
    orientation = curr_region.orientation
    if orientation == "long":
        return _align_long(base_sheet, curr_sheet, base_region, curr_region)
    if orientation == "wide":
        return _align_wide(base_sheet, curr_sheet, base_region, curr_region)
    return _align_block(base_sheet, curr_sheet, base_region, curr_region)


# --- workbook alignment ---------------------------------------------------


def _pair_regions(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    base_regions: list[TableRegion],
    curr_regions: list[TableRegion],
) -> tuple[list[tuple[TableRegion, TableRegion]], list[TableRegion], list[TableRegion]]:
    """Pair regions by stable content and structure, not document order."""

    def labels(sheet: SheetSnapshot, region: TableRegion) -> frozenset[str]:
        result: set[str] = set()
        for (row, col), cell in sheet.cells.items():
            if not (
                region.min_row <= row <= region.max_row
                and region.min_col <= col <= region.max_col
            ):
                continue
            if not isinstance(cell.value, str) or is_period_label(cell.value):
                continue
            text = " ".join(cell.value.casefold().split())
            if text and not text.startswith("="):
                result.add(re.sub(r"\d+(?:[.,]\d+)*", "#", text))
        return frozenset(result)

    def similarity(base: TableRegion, curr: TableRegion) -> float:
        if base.orientation != curr.orientation:
            return -1.0
        base_labels = labels(base_sheet, base)
        curr_labels = labels(curr_sheet, curr)
        if base_labels or curr_labels:
            shared = len(base_labels & curr_labels)
            label_score = 2 * shared / max(len(base_labels) + len(curr_labels), 1)
        else:
            label_score = 0.25
        base_rows = base.max_row - base.min_row + 1
        curr_rows = curr.max_row - curr.min_row + 1
        base_cols = base.max_col - base.min_col + 1
        curr_cols = curr.max_col - curr.min_col + 1
        shape_score = (
            min(base_rows, curr_rows) / max(base_rows, curr_rows)
            + min(base_cols, curr_cols) / max(base_cols, curr_cols)
        ) / 2
        distance = abs(base.min_row - curr.min_row) + abs(base.min_col - curr.min_col)
        position_score = 1 / (1 + distance)
        return 0.65 * label_score + 0.25 * shape_score + 0.10 * position_score

    candidates = sorted(
        (
            (similarity(base, curr), base_index, curr_index)
            for base_index, base in enumerate(base_regions)
            for curr_index, curr in enumerate(curr_regions)
        ),
        key=lambda item: (-item[0], item[1], item[2]),
    )
    used_base: set[int] = set()
    used_curr: set[int] = set()
    indexed_pairs: list[tuple[int, int]] = []
    for score, base_index, curr_index in candidates:
        if score < 0.20:
            break
        if base_index in used_base or curr_index in used_curr:
            continue
        used_base.add(base_index)
        used_curr.add(curr_index)
        indexed_pairs.append((base_index, curr_index))
    indexed_pairs.sort()
    return (
        [(base_regions[b], curr_regions[c]) for b, c in indexed_pairs],
        [region for index, region in enumerate(base_regions) if index not in used_base],
        [region for index, region in enumerate(curr_regions) if index not in used_curr],
    )


def align_workbooks(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    profile: DeliverableProfile | None = None,
    *,
    cancellation_token: CancellationToken | None = None,
) -> WorkbookAlignment:
    ignore = set(profile.excel.ignore_sheets) if profile else set()
    base_names = [n for n in baseline.sheet_names if n not in ignore]
    curr_names = [n for n in current.sheet_names if n not in ignore]

    result = WorkbookAlignment(
        common_sheets=[n for n in base_names if n in set(curr_names)],
        added_sheets=[n for n in curr_names if n not in set(base_names)],
        removed_sheets=[n for n in base_names if n not in set(curr_names)],
    )

    for sheet_name in result.common_sheets:
        check_cancelled(cancellation_token)
        sheet_profile = profile.sheet_profile(sheet_name) if profile else None
        if sheet_profile is not None and sheet_profile.ignore:
            continue
        base_sheet = baseline.sheet(sheet_name)
        curr_sheet = current.sheet(sheet_name)
        base_regions = detect_regions(base_sheet, sheet_profile)
        curr_regions = detect_regions(curr_sheet, sheet_profile)
        pairs, unpaired_base, unpaired_curr = _pair_regions(
            base_sheet, curr_sheet, base_regions, curr_regions
        )
        result.unpaired_baseline_regions.extend(unpaired_base)
        result.unpaired_current_regions.extend(unpaired_curr)
        region_alignments = [
            align_regions(base_sheet, curr_sheet, base_region, curr_region)
            for base_region, curr_region in pairs
        ]
        result.regions[sheet_name] = region_alignments
        result.low_confidence_regions.extend(
            f"{sheet_name}!{region.current.cell_range}"
            for region in region_alignments
            if region.low_confidence
        )
    return result
