"""Fail-closed contracts for enriching XLSB snapshots with formula text."""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass

from qc_tool.io.model import (
    CellRecord,
    ExternalLinkReachability,
    FormulaTextCoverage,
    WorkbookSnapshot,
)
from qc_tool.io.xlsb_formula import XlsbFormulaScan

FormulaCoordinate = tuple[int, int]
FormulaMap = dict[str, dict[FormulaCoordinate, str]]

#: Bounds on adapter-supplied defined names. Real workbooks rarely declare
#: more than a few hundred names, so these bound a malformed or hostile
#: adapter response without ever needing to grow for a legitimate workbook.
MAX_EXTRACTED_DEFINED_NAMES = 10_000
MAX_DEFINED_NAME_TARGET_CHARS = 4_096
#: Quadratic name-to-name closure stays bounded below this limit. Larger,
#: otherwise-valid collections remain usable formula data but cannot prove an
#: external link inactive, so reachability degrades fail-closed.
MAX_REACHABILITY_DEFINED_NAMES = 1_000

#: Excel/OOXML external-workbook-reference syntax: a bracketed numeric
#: external-workbook index preceding a sheet reference (e.g. ``[1]Sheet1!A1``
#: or ``'[1]Sheet1'!A1``). A bounded textual heuristic, not a formula parser
#: -- it only recognizes this well-known external-reference shape.
_EXTERNAL_REFERENCE_PATTERN = re.compile(r"\[\d+\]")
_EXTERNAL_WORKBOOK_FILE_PATTERN = re.compile(
    r"\[[^\]]+\.(?:xls|xlsx|xlsm|xlsb)\]", re.IGNORECASE
)


def _contains_external_reference(text: str) -> bool:
    """Conservative recognition across Excel and LibreOffice serializations."""
    lowered = text.casefold()
    return bool(
        _EXTERNAL_REFERENCE_PATTERN.search(text)
        or _EXTERNAL_WORKBOOK_FILE_PATTERN.search(text)
        or "file://" in lowered
        or "file:\\" in lowered
    )


class FormulaEnrichmentError(RuntimeError):
    """Formula text cannot be trusted for this workbook."""


@dataclass(frozen=True, slots=True)
class ExtractedDefinedName:
    """One defined name returned by an external formula-extraction adapter.

    Held only in worker memory and this module's bounded, count-only
    reachability verdict; the name and target text are never logged, stored,
    or surfaced to reports/history.
    """

    name: str
    target: str
    sheet: str | None = None
    hidden: bool = False


@dataclass(frozen=True, slots=True)
class FormulaExtraction:
    """Formula text produced by one external spreadsheet engine."""

    formulas: FormulaMap
    engine: str
    detail: str
    #: Workbook- and worksheet-scoped defined names, when the adapter could
    #: collect them. Empty when unavailable; see `defined_names_complete`.
    defined_names: tuple[ExtractedDefinedName, ...] = ()
    #: Whether `defined_names` reflects a clean, uncapped, duplicate-free
    #: collection. False means defined-name reachability cannot be proven,
    #: even though any collected names remain valid data.
    defined_names_complete: bool = False
    #: Canonical, `=`-prefixed R1C1 text per coordinate, when the adapter can
    #: supply it per definition (currently: the native XLSB kernel only).
    #: `None` for adapters that never produce it (Excel COM, LibreOffice);
    #: `merge_formula_extraction` leaves `CellRecord.formula_r1c1` at its
    #: default `None` in that case, and formula comparison falls back to
    #: computing R1C1 itself from `CellRecord.formula`.
    formulas_r1c1: FormulaMap | None = None

    @property
    def coordinates(self) -> frozenset[tuple[str, int, int]]:
        return frozenset(
            (sheet, row, column)
            for sheet, cells in self.formulas.items()
            for row, column in cells
        )


def bounded_defined_names(
    raw: Sequence[ExtractedDefinedName],
) -> tuple[tuple[ExtractedDefinedName, ...], bool]:
    """Cap and de-duplicate adapter-supplied defined names.

    Returns ``(names, complete)``. ``complete`` is False when any entry was
    dropped for a duplicate ``(name, sheet)`` identity, an oversized target,
    or the total count cap, so callers never claim defined-name reachability
    is exhaustive over degraded input. Shared by the Windows Excel worker and
    the Linux LibreOffice adapter.
    """
    seen: set[tuple[str, str | None]] = set()
    accepted: list[ExtractedDefinedName] = []
    complete = True
    for item in raw:
        key = (item.name, item.sheet)
        if key in seen or len(item.target) > MAX_DEFINED_NAME_TARGET_CHARS:
            complete = False
            continue
        if len(accepted) >= MAX_EXTRACTED_DEFINED_NAMES:
            complete = False
            continue
        seen.add(key)
        accepted.append(item)
    return tuple(accepted), complete


def scan_coordinates(scan: XlsbFormulaScan) -> frozenset[tuple[str, int, int]]:
    """Flatten scan coordinates for exact two-way parity checks."""
    return frozenset(
        (sheet, row, column)
        for sheet, cells in scan.formula_cells.items()
        for row, column in cells
    )


def validate_formula_extraction(
    scan: XlsbFormulaScan, extraction: FormulaExtraction
) -> None:
    """Reject any unexpected coordinate or invalid formula text before mutation.

    A scanned coordinate the extraction did not return (missing) is
    tolerated -- it becomes partial coverage rather than a fatal error. Any
    coordinate the external engine returned that the independent scan did
    NOT expect remains fatal, with no mutation to the snapshot.
    """
    expected = scan_coordinates(scan)
    actual = extraction.coordinates
    unexpected = actual - expected
    if unexpected:
        raise FormulaEnrichmentError(
            f"formula-coordinate mismatch: {len(unexpected)} unexpected coordinates"
        )
    for sheet, cells in extraction.formulas.items():
        for coordinate, formula in cells.items():
            if not isinstance(formula, str) or not formula.startswith("="):
                raise FormulaEnrichmentError(
                    f"{sheet}!{coordinate}: external engine returned invalid formula text"
                )


def compute_formula_text_coverage(
    scan: XlsbFormulaScan, extraction: FormulaExtraction
) -> FormulaTextCoverage:
    """Compute coverage state; call only after `validate_formula_extraction` passes."""
    expected = scan_coordinates(scan)
    merged = extraction.coordinates
    missing = len(expected - merged)
    if missing == 0:
        return FormulaTextCoverage(
            state="complete",
            expected_count=len(expected),
            merged_count=len(merged),
            missing_count=0,
            detail="every scanned formula coordinate was merged",
        )
    return FormulaTextCoverage(
        state="partial",
        expected_count=len(expected),
        merged_count=len(merged),
        missing_count=missing,
        detail=(
            f"{missing} of {len(expected)} scanned formula coordinates were not "
            "returned by the external engine"
        ),
    )


def _references_name(formula: str, name: str) -> bool:
    """Whether `formula` contains `name` as a whole token (bounded heuristic)."""
    pattern = rf"(?<![A-Za-z0-9_.]){re.escape(name)}(?![A-Za-z0-9_.])"
    return re.search(pattern, formula, re.IGNORECASE) is not None


def _compile_name_matcher(names: set[str]) -> re.Pattern[str] | None:
    """One alternation equivalent to ``any(_references_name(text, n) for n in names)``."""
    if not names:
        return None
    alternatives = "|".join(re.escape(name) for name in sorted(names, key=len, reverse=True))
    return re.compile(rf"(?<![A-Za-z0-9_.])(?:{alternatives})(?![A-Za-z0-9_.])", re.IGNORECASE)


def classify_external_reachability(
    extraction: FormulaExtraction, coverage: FormulaTextCoverage
) -> ExternalLinkReachability:
    """Whether merged formula text and defined names prove a passive link unused.

    ``proven`` requires complete formula-text coverage (`coverage.state ==
    "complete"`) AND a complete defined-name collection; anything less
    cannot prove inactivity and stays a bounded degraded disclosure. Only
    counts are retained here -- never the referencing formula or
    defined-name text.
    """
    if coverage.state != "complete" or not extraction.defined_names_complete:
        return ExternalLinkReachability(
            proven=False,
            detail=(
                "formula text or defined-name collection is incomplete; "
                "external-link usage cannot be proven"
            ),
        )
    if len(extraction.defined_names) > MAX_REACHABILITY_DEFINED_NAMES:
        return ExternalLinkReachability(
            proven=False,
            detail=(
                "defined-name population exceeds the bounded reachability limit; "
                "external-link usage cannot be proven"
            ),
        )
    external_names = {
        item.name
        for item in extraction.defined_names
        if _contains_external_reference(item.target)
    }
    # Close over defined-name aliases before checking formulas. A formula may
    # reference LocalAlias -> WorkbookAlias -> [1]ExternalSheet!A1; only the
    # first draft's single-hop check would miss that live dependency.
    changed = True
    while changed:
        changed = False
        for item in extraction.defined_names:
            if item.name in external_names:
                continue
            if any(_references_name(item.target, name) for name in external_names):
                external_names.add(item.name)
                changed = True
    direct = 0
    transitive = 0
    # Millions of formulas x dozens of names: match all names in one pass.
    matcher = _compile_name_matcher(external_names)
    for cells in extraction.formulas.values():
        for formula in cells.values():
            if _contains_external_reference(formula):
                direct += 1
                continue
            if matcher is not None and matcher.search(formula) is not None:
                transitive += 1
    live = bool(direct or transitive)
    return ExternalLinkReachability(
        proven=True,
        live=live,
        direct_reference_count=direct,
        transitive_reference_count=transitive,
        detail=(
            "at least one current formula depends on the external link"
            if live
            else "no current formula depends on the external link"
        ),
    )


def merge_formula_extraction(
    snapshot: WorkbookSnapshot,
    scan: XlsbFormulaScan,
    extraction: FormulaExtraction,
) -> None:
    """Merge trusted formula text without replacing original cached values.

    Tolerates a partial merge (some scanned formula coordinates missing from
    `extraction`); only an unexpected coordinate or invalid formula text
    remains fatal before any mutation.
    """
    if snapshot.file_format != "xlsb":
        raise FormulaEnrichmentError("formula enrichment is only valid for XLSB snapshots")
    validate_formula_extraction(scan, extraction)

    for sheet_name, formulas in extraction.formulas.items():
        try:
            sheet = snapshot.sheet(sheet_name)
        except KeyError as exc:
            raise FormulaEnrichmentError(
                f"external engine returned unknown sheet {sheet_name!r}"
            ) from exc
        r1c1_for_sheet = (
            extraction.formulas_r1c1.get(sheet_name)
            if extraction.formulas_r1c1 is not None
            else None
        )
        for (row, column), formula in formulas.items():
            cell = sheet.cells.get((row, column))
            if cell is None:
                cell = CellRecord(row=row, column=column, value=None, is_formula=True)
                sheet.cells[(row, column)] = cell
                sheet.max_row = max(sheet.max_row, row)
                sheet.max_column = max(sheet.max_column, column)
            cell.formula = formula
            cell.is_formula = True
            if r1c1_for_sheet is not None:
                cell.formula_r1c1 = r1c1_for_sheet.get((row, column))

    coverage = compute_formula_text_coverage(scan, extraction)
    snapshot.formula_presence_available = True
    snapshot.formulas_available = coverage.state == "complete"
    snapshot.formula_source = extraction.engine
    snapshot.formula_detail = (
        extraction.detail
        if coverage.state == "complete"
        else f"{extraction.detail}; {coverage.detail}"
    )
    snapshot.formula_text_coverage = coverage
    if scan.passive_features:
        snapshot.external_link_reachability = classify_external_reachability(
            extraction, coverage
        )
