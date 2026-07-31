"""Formula QC: errors, hardcoding, logic changes, extension, consistency.

Comparisons are shift-invariant: every formula is normalized to R1C1
relative to its host cell, so a formula that merely moved with inserted
cadence data compares equal. Where normalized forms differ only in range
*end* references that grow (same sheet, same anchor, larger end), the
change is classified as an expected range extension, not a logic change.

Checks (paired regions come from the alignment engine):

- error values anywhere in the current workbook, and error literals
  inside formula text;
- baseline formula replaced by a constant (hardcode);
- normalized formula changed vs baseline (minus expected extensions);
- growth rows/columns missing the formula pattern of their axis run
  (not extended);
- formula cells deviating from the dominant R1C1 pattern of their run
  (inconsistent within range).

xlsb workbooks may expose formula-record presence without formula text. That
presence safely enables hardcode, removal, and missing fill checks; semantic
logic and consistency checks require compatible decoded text on both sides.
"""

import logging
import re
from collections import Counter

from openpyxl.utils import get_column_letter

from qc_tool.availability import cell_in_ranges, excel_blank_allowed
from qc_tool.config.profile import DeliverableProfile, SheetProfile
from qc_tool.excel.align import RegionAlignment, WorkbookAlignment
from qc_tool.excel.formula_tokens import tokenize_formula
from qc_tool.findings import Finding, FindingClass
from qc_tool.io.model import (
    ERROR_LITERALS,
    CellRecord,
    SheetSnapshot,
    WorkbookSnapshot,
    display_cell_value,
)
from qc_tool.progress import CancellationToken, check_cancelled

logger = logging.getLogger(__name__)

_ENDPOINT_RE = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$")
_RUN_MIN_CELLS = 3
_RUN_DOMINANCE = 0.5
_FORMULAIC_SHARE = 0.5


def _ref(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


# --- R1C1 normalization ---------------------------------------------------


def _endpoint_to_r1c1(endpoint: str, host_row: int, host_col: int) -> str | None:
    match = _ENDPOINT_RE.match(endpoint)
    if match is None:
        return None
    col_abs, col_letters, row_abs, row_digits = match.groups()
    from openpyxl.utils.cell import column_index_from_string

    col = column_index_from_string(col_letters)
    row = int(row_digits)
    row_part = f"R{row}" if row_abs else ("R" if row == host_row else f"R[{row - host_row}]")
    col_part = f"C{col}" if col_abs else ("C" if col == host_col else f"C[{col - host_col}]")
    return row_part + col_part


def _range_token_to_r1c1(token: str, host_row: int, host_col: int) -> str:
    sheet_prefix, sep, ref = token.rpartition("!")
    parts = ref.split(":")
    converted: list[str] = []
    for part in parts:
        r1c1 = _endpoint_to_r1c1(part, host_row, host_col)
        if r1c1 is None:
            return token  # named range, whole-row/col ref, etc. — keep verbatim
        converted.append(r1c1)
    return sheet_prefix + sep + ":".join(converted)


def to_r1c1(formula: str, host_row: int, host_col: int) -> str:
    """Normalize an A1-style formula to R1C1 relative to its host cell."""
    try:
        tokens = tokenize_formula(formula)
    except Exception:  # malformed formulas must not kill a run
        logger.warning("unparseable formula at %s: %r", _ref(host_row, host_col), formula)
        return formula
    rendered = [
        _range_token_to_r1c1(t.value, host_row, host_col)
        if t.type == "OPERAND" and t.subtype == "RANGE"
        else t.value
        for t in tokens
    ]
    return "=" + "".join(rendered)


# --- expected range extension ---------------------------------------------


def _split_range(token: str) -> tuple[str, str, str] | None:
    sheet, _, ref = token.rpartition("!")
    start, sep, end = ref.partition(":")
    if not sep:
        return None
    return sheet, start, end


def _is_range_extension(base_token: str, curr_token: str) -> bool:
    base_parts = _split_range(base_token)
    curr_parts = _split_range(curr_token)
    if base_parts is None or curr_parts is None:
        return False
    if base_parts[0] != curr_parts[0] or base_parts[1] != curr_parts[1]:
        return False
    base_end = _ENDPOINT_RE.match(base_parts[2])
    curr_end = _ENDPOINT_RE.match(curr_parts[2])
    if base_end is None or curr_end is None:
        return False
    same_col = base_end[2] == curr_end[2]
    same_row = base_end[4] == curr_end[4]
    if same_col and int(curr_end[4]) >= int(base_end[4]):
        return True
    return bool(
        same_row
        and (
            len(curr_end[2]) > len(base_end[2])
            or (len(curr_end[2]) == len(base_end[2]) and curr_end[2] >= base_end[2])
        )
    )


def _differs_only_by_extension(base_formula: str, curr_formula: str) -> bool:
    try:
        base_tokens = tokenize_formula(base_formula)
        curr_tokens = tokenize_formula(curr_formula)
    except Exception:  # malformed formulas cannot be extension-classified
        return False
    if len(base_tokens) != len(curr_tokens):
        return False
    extension_seen = False
    for base_tok, curr_tok in zip(base_tokens, curr_tokens, strict=True):
        if (base_tok.type, base_tok.subtype) != (curr_tok.type, curr_tok.subtype):
            return False
        if base_tok.value == curr_tok.value:
            continue
        if base_tok.type == "OPERAND" and base_tok.subtype == "RANGE":
            if not _is_range_extension(base_tok.value, curr_tok.value):
                return False
            extension_seen = True
        else:
            return False
    return extension_seen


# --- error scan -------------------------------------------------------------


def _ignored(profile: SheetProfile | None, row: int, column: int) -> bool:
    return bool(
        profile is not None
        and cell_in_ranges(row, column, profile.ignore_ranges)
    )


def _error_findings(
    current: WorkbookSnapshot,
    profile: DeliverableProfile | None,
) -> list[Finding]:
    findings = []
    for sheet in current.sheets:
        sheet_profile = profile.sheet_profile(sheet.name) if profile is not None else None
        for (row, col), cell in sorted(sheet.cells.items()):
            if _ignored(sheet_profile, row, col):
                continue
            location = _ref(row, col)
            if cell.is_error:
                findings.append(
                    Finding(
                        artifact="excel",
                        finding_class=FindingClass.FORMULA_ERROR,
                        sheet=sheet.name,
                        location=location,
                        current_value=display_cell_value(cell.value),
                        message=f"{sheet.name}!{location}: error value {cell.value}",
                    )
                )
            if cell.formula and any(err in cell.formula for err in ERROR_LITERALS):
                findings.append(
                    Finding(
                        artifact="excel",
                        finding_class=FindingClass.FORMULA_ERROR,
                        sheet=sheet.name,
                        location=location,
                        current_value=cell.formula,
                        message=(
                            f"{sheet.name}!{location}: formula contains an error "
                            f"reference ({cell.formula})"
                        ),
                    )
                )
    return findings


# --- paired-cell checks ------------------------------------------------------


def _paired_cell_findings(
    base_sheet: SheetSnapshot,
    curr_sheet: SheetSnapshot,
    region: RegionAlignment,
    *,
    compare_text: bool,
    profile: SheetProfile | None,
) -> list[Finding]:
    findings = []
    sheet_name = curr_sheet.name
    for (base_row, base_col), (curr_row, curr_col) in region.cell_pairs():
        base_cell = base_sheet.cells.get((base_row, base_col))
        curr_cell = curr_sheet.cells.get((curr_row, curr_col))
        if _ignored(profile, curr_row, curr_col):
            continue
        base_has_formula = base_cell is not None and base_cell.has_formula
        curr_has_formula = curr_cell is not None and curr_cell.has_formula
        if not base_has_formula and not curr_has_formula:
            continue
        location = _ref(curr_row, curr_col)

        if base_has_formula and not curr_has_formula:
            baseline_value = (
                base_cell.formula
                if base_cell is not None and base_cell.formula is not None
                else "formula record"
            )
            if curr_cell is not None and curr_cell.value is not None:
                findings.append(
                    Finding(
                        artifact="excel",
                        finding_class=FindingClass.FORMULA_HARDCODED,
                        sheet=sheet_name,
                        location=location,
                        baseline_location=_ref(base_row, base_col),
                        baseline_value=baseline_value,
                        current_value=display_cell_value(curr_cell.value),
                        message=(
                            f"{sheet_name}!{location}: formula replaced by "
                            f"hardcoded value {curr_cell.value}"
                        ),
                    )
                )
            else:
                if excel_blank_allowed(
                    curr_sheet,
                    profile,
                    curr_row,
                    curr_col,
                ):
                    continue
                findings.append(
                    Finding(
                        artifact="excel",
                        finding_class=FindingClass.FORMULA_REMOVED,
                        sheet=sheet_name,
                        location=location,
                        baseline_location=_ref(base_row, base_col),
                        baseline_value=baseline_value,
                        message=f"{sheet_name}!{location}: formula cleared or removed",
                    )
                )
            continue

        if base_has_formula and curr_has_formula and compare_text:
            base_formula = base_cell.formula if base_cell is not None else None
            curr_formula = curr_cell.formula if curr_cell is not None else None
            if base_formula is None or curr_formula is None:
                logger.warning(
                    "%s!%s: formula text capability contradicted cell data",
                    sheet_name,
                    location,
                )
                continue
            base_norm = to_r1c1(base_formula, base_row, base_col)
            curr_norm = to_r1c1(curr_formula, curr_row, curr_col)
            if base_norm == curr_norm:
                continue
            expected = _differs_only_by_extension(base_formula, curr_formula)
            findings.append(
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
                    expected_growth=expected,
                    sheet=sheet_name,
                    location=location,
                    baseline_location=_ref(base_row, base_col),
                    baseline_value=base_formula,
                    current_value=curr_formula,
                    message=(
                        f"{sheet_name}!{location}: "
                        + (
                            "formula range extended with new-cycle data"
                            if expected
                            else "formula logic changed"
                        )
                    ),
                )
            )
    return findings


# --- run helpers (formula pattern along the data axis) -----------------------


def _run_axes(
    region: RegionAlignment,
) -> tuple[list[int], list[int]] | None:
    """(run positions, cell positions) as (fixed axis indices, moving indices)."""
    current = region.current
    if current.orientation == "long":
        data_rows = sorted([c for _, c in region.rows.pairs] + region.rows.growth)
        columns = sorted([c for _, c in region.columns.pairs] + region.columns.growth)
        return columns, data_rows
    if current.orientation == "wide":
        header = current.header_row or current.min_row
        data_rows = sorted(
            [c for _, c in region.rows.pairs if c != header] + region.rows.growth
        )
        columns = sorted([c for _, c in region.columns.pairs] + region.columns.growth)
        return data_rows, columns
    return None


def _run_cells(
    sheet: SheetSnapshot, region: RegionAlignment, fixed: int, moving: list[int]
) -> list[tuple[int, int, CellRecord]]:
    is_long = region.current.orientation == "long"
    out = []
    for position in moving:
        key = (position, fixed) if is_long else (fixed, position)
        cell = sheet.cells.get(key)
        if cell is not None:
            out.append((key[0], key[1], cell))
    return out


def _consistency_findings(
    curr_sheet: SheetSnapshot,
    region: RegionAlignment,
    profile: SheetProfile | None,
) -> list[Finding]:
    axes = _run_axes(region)
    if axes is None:
        return []
    run_positions, cell_positions = axes
    findings = []
    for fixed in run_positions:
        cells = _run_cells(curr_sheet, region, fixed, cell_positions)
        formula_cells = [
            (row, col, cell)
            for row, col, cell in cells
            if cell.has_formula and not _ignored(profile, row, col)
        ]
        if len(formula_cells) < _RUN_MIN_CELLS:
            continue
        patterns = Counter(
            to_r1c1(cell.formula or "", row, col) for row, col, cell in formula_cells
        )
        dominant, dominant_count = patterns.most_common(1)[0]
        if dominant_count / len(formula_cells) <= _RUN_DOMINANCE:
            continue
        for row, col, cell in formula_cells:
            if to_r1c1(cell.formula or "", row, col) == dominant:
                continue
            location = _ref(row, col)
            findings.append(
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.FORMULA_INCONSISTENT,
                    sheet=curr_sheet.name,
                    location=location,
                    baseline_value=dominant,
                    current_value=cell.formula,
                    message=(
                        f"{curr_sheet.name}!{location}: formula deviates from the "
                        f"dominant pattern of its range ({dominant})"
                    ),
                )
            )
    return findings


def _extension_findings(
    curr_sheet: SheetSnapshot,
    region: RegionAlignment,
    profile: SheetProfile | None,
) -> list[Finding]:
    """Growth rows/columns must carry the formula pattern of their run."""
    current = region.current
    findings = []

    def check(
        growth: list[int], run_positions: list[int], paired: list[int], *, rows_grow: bool
    ) -> None:
        for fixed in run_positions:
            paired_cells = [
                curr_sheet.cells.get((p, fixed) if rows_grow else (fixed, p)) for p in paired
            ]
            populated = [c for c in paired_cells if c is not None]
            with_formula = [c for c in populated if c.has_formula]
            if len(populated) < 2 or len(with_formula) / len(populated) < _FORMULAIC_SHARE:
                continue
            for g in growth:
                key = (g, fixed) if rows_grow else (fixed, g)
                if _ignored(profile, *key):
                    continue
                cell = curr_sheet.cells.get(key)
                if cell is not None and cell.has_formula:
                    continue
                if excel_blank_allowed(curr_sheet, profile, *key):
                    continue
                location = _ref(*key)
                findings.append(
                    Finding(
                        artifact="excel",
                        finding_class=FindingClass.FORMULA_NOT_EXTENDED,
                        sheet=curr_sheet.name,
                        location=location,
                        current_value=None if cell is None else display_cell_value(cell.value),
                        message=(
                            f"{curr_sheet.name}!{location}: new-cycle cell is missing "
                            "the formula used by its historical range"
                        ),
                    )
                )

    if current.orientation == "long" and region.rows.growth:
        columns = sorted([c for _, c in region.columns.pairs] + region.columns.growth)
        paired_rows = [c for _, c in region.rows.pairs]
        check(region.rows.growth, columns, paired_rows, rows_grow=True)
    elif current.orientation == "wide" and region.columns.growth:
        header = current.header_row or current.min_row
        data_rows = [c for _, c in region.rows.pairs if c != header] + region.rows.growth
        paired_cols = [c for _, c in region.columns.pairs]
        check(region.columns.growth, sorted(data_rows), paired_cols, rows_grow=False)
    return findings


# --- entry point --------------------------------------------------------------


def formula_text_compatible(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot
) -> bool:
    """Whether formula text can be compared without crossing engine dialects."""
    return bool(
        baseline.formulas_available
        and current.formulas_available
        and baseline.formula_source == current.formula_source
    )


def diff_workbook_formulas(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    alignment: WorkbookAlignment,
    profile: DeliverableProfile | None = None,
    *,
    cancellation_token: CancellationToken | None = None,
) -> list[Finding]:
    findings = _error_findings(current, profile)
    presence_pair = bool(
        baseline.formula_presence_available and current.formula_presence_available
    )
    compare_text = formula_text_compatible(baseline, current)
    if not presence_pair:
        logger.info(
            "formula presence unavailable (%s/%s): formula QC limited to error values",
            baseline.file_format,
            current.file_format,
        )
        return findings
    if not compare_text:
        logger.info(
            "compatible formula text unavailable (%s/%s): semantic formula QC skipped",
            baseline.formula_source,
            current.formula_source,
        )
    for sheet_name, regions in alignment.regions.items():
        check_cancelled(cancellation_token)
        base_sheet = baseline.sheet(sheet_name)
        curr_sheet = current.sheet(sheet_name)
        sheet_profile = profile.sheet_profile(sheet_name) if profile is not None else None
        for region in regions:
            check_cancelled(cancellation_token)
            if region.low_confidence:
                continue
            findings.extend(
                _paired_cell_findings(
                    base_sheet,
                    curr_sheet,
                    region,
                    compare_text=compare_text,
                    profile=sheet_profile,
                )
            )
            if compare_text:
                findings.extend(
                    _consistency_findings(curr_sheet, region, sheet_profile)
                )
            findings.extend(
                _extension_findings(curr_sheet, region, sheet_profile)
            )
    return findings
