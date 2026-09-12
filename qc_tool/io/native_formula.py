"""Native Excel formula surface adapter (Rust/PyO3,
``native/cadence_diff_native/``'s `formula_surface_report`).

Exposes B2's "definitions + per-cell ids + names surface": per-cell eager A1
text (every existing `CellRecord.formula` consumer expects per-cell A1),
canonical R1C1 rendered once per shared/array definition (a `definition_id`
links each cell to its group), and defined names with rendered A1 targets.

The helper distribution can resolve to a native extension or an importable
pure-Python fallback module. Availability and compatibility are therefore
owned exclusively by ``qc_tool.io.native_kernel``.

`extract_formulas_with_native_kernel` (B3) is the actual `FormulaExtraction`
adapter `qc_tool/io/loader.py` dispatches to for `formula_engine: native`;
`formula_surface_report` above stays the raw, unshaped B2 probe surface.

See ``docs/plans/plan-20260906-group-first-and-native-kernel.md``'s Phase B
for the full design.
"""

from __future__ import annotations

from dataclasses import dataclass

from qc_tool.io import native_kernel
from qc_tool.io.formula_enrichment import (
    ExtractedDefinedName,
    FormulaEnrichmentError,
    FormulaExtraction,
    FormulaMap,
    bounded_defined_names,
)
from qc_tool.io.xlsb_formula import XlsbFormulaScan


@dataclass(slots=True)
class FormulaSheetSurface:
    """One sheet's formula cells plus its deduplicated definition table.

    ``cell_definition_id[i]`` indexes into ``definition_r1c1``: every
    follower of one shared/array group shares the same entry, rendered once.
    """

    name: str
    cell_rows: tuple[int, ...]
    cell_cols: tuple[int, ...]
    cell_a1: tuple[str, ...]
    cell_definition_id: tuple[int, ...]
    definition_r1c1: tuple[str, ...]


@dataclass(slots=True)
class DefinedNameSurface:
    """One defined name; ``target_a1`` is ``None`` if it failed to decode
    (fail closed, never guessed)."""

    name: str
    target_a1: str | None


@dataclass(slots=True)
class WorkbookFormulaSurface:
    sheets: tuple[FormulaSheetSurface, ...]
    defined_names: tuple[DefinedNameSurface, ...]


def native_formula_available() -> bool:
    return native_kernel.native_kernel_available()


def _validate_sheet_surface(sheet: FormulaSheetSurface) -> None:
    """Fail closed on a malformed/mismatched native surface shape before it
    is ever indexed or zipped elsewhere -- never let a native-side shape
    violation escape as a bare ``IndexError``/``ValueError``.

    Messages never name a sheet, coordinate, or formula text; a shape
    violation is bounded, content-free evidence by construction.
    """
    lengths = {
        len(sheet.cell_rows),
        len(sheet.cell_cols),
        len(sheet.cell_a1),
        len(sheet.cell_definition_id),
    }
    if len(lengths) > 1:
        raise FormulaEnrichmentError(
            "the cadence-diff native engine returned a formula surface with "
            "mismatched per-cell vector lengths"
        )
    definition_count = len(sheet.definition_r1c1)
    if any(
        not (0 <= definition_id < definition_count)
        for definition_id in sheet.cell_definition_id
    ):
        raise FormulaEnrichmentError(
            "the cadence-diff native engine returned a formula surface with an "
            "out-of-range definition id"
        )


def _validated_r1c1_cells(sheet: FormulaSheetSurface) -> dict[tuple[int, int], str]:
    """Per-cell canonical R1C1 (1-based coordinates), cross-checked once per
    distinct definition rather than trusted blindly.

    The kernel renders R1C1 once per shared/array definition
    (``definition_r1c1``, indexed by ``cell_definition_id``) -- far cheaper
    than re-tokenizing every cell's A1 text in Python. A narrow, disclosed
    residual gap in the kernel's R1C1 rendering has not been root-caused
    (see ``native/cadence_diff_native/src/ptg.rs``'s module doc comment), so this
    validates the kernel's per-definition text against the already-Excel-
    proven ``to_r1c1()`` -- using one representative cell's own A1 text --
    exactly once per distinct definition, and falls back to that proven
    value for every cell sharing a definition that fails validation. Cost is
    bounded by the number of distinct definitions (hundreds of thousands on
    a large real workbook), not the number of formula cells (millions),
    while the correctness risk is eliminated regardless of the gap's root
    cause.
    """
    from qc_tool.excel.formulas import to_r1c1

    validated: dict[int, str] = {}
    result: dict[tuple[int, int], str] = {}
    for row, col, a1_text, definition_id in zip(
        sheet.cell_rows,
        sheet.cell_cols,
        sheet.cell_a1,
        sheet.cell_definition_id,
        strict=True,
    ):
        cached = validated.get(definition_id)
        if cached is None:
            kernel_r1c1 = "=" + sheet.definition_r1c1[definition_id]
            reference_r1c1 = to_r1c1(f"={a1_text}", row + 1, col + 1)
            cached = kernel_r1c1 if kernel_r1c1 == reference_r1c1 else reference_r1c1
            validated[definition_id] = cached
        result[(row + 1, col + 1)] = cached
    return result


def formula_surface_report(data: bytes) -> WorkbookFormulaSurface:
    """Decodes the full production formula surface from in-memory XLSB bytes.

    Raises ``RuntimeError`` if the extension is not installed; callers must
    check ``native_formula_available()`` first and fall back to the existing
    Excel-COM/LibreOffice adapters instead of relying on this exception as
    control flow.
    """
    sheets_raw, names_raw = native_kernel.formula_surface_report(data)
    sheets = tuple(
        FormulaSheetSurface(
            name=name,
            cell_rows=tuple(rows),
            cell_cols=tuple(cols),
            cell_a1=tuple(a1_texts),
            cell_definition_id=tuple(def_ids),
            definition_r1c1=tuple(def_r1c1),
        )
        for name, rows, cols, a1_texts, def_ids, def_r1c1 in sheets_raw
    )
    defined_names = tuple(DefinedNameSurface(name=n, target_a1=target) for n, target in names_raw)
    return WorkbookFormulaSurface(sheets=sheets, defined_names=defined_names)


def _kernel_version() -> str | None:
    return native_kernel.native_distribution_version()


def native_adapter_fingerprint() -> str | None:
    """A cheap kernel-version probe for formula-cache keys, or None.

    Mirrors ``excel_adapter_fingerprint``/``libreoffice_adapter_fingerprint``:
    never opens a workbook, returns None (caching disabled) rather than
    guessing when the installed distribution's version cannot be read.
    """
    version = _kernel_version()
    return f"cadence-diff-native:{version}" if version else None


def extract_formulas_with_native_kernel(
    data: bytes, formula_scan: XlsbFormulaScan
) -> FormulaExtraction:
    """The native kernel's `FormulaExtraction` adapter (`formula_engine: native`).

    Converts B2's raw production surface (0-based coordinates, bare
    token-stream A1 text with no leading ``=``) into the same shape the
    Excel-COM/LibreOffice adapters return (1-based coordinates, ``=``-
    prefixed text) so every downstream consumer -- `merge_formula_extraction`,
    `validate_formula_extraction`, `qc_tool/excel/formulas.py` -- stays
    unchanged regardless of engine. A cell the kernel could not decode (an
    unknown token, or a malformed record) is simply absent from the surface
    and so absent here too -- fail-closed, never a guess; `merge_formula_
    extraction`'s existing partial-coverage handling applies unchanged.

    Raises ``FormulaEnrichmentError`` (not a bare ``RuntimeError``) when the
    extension is not installed, so callers already catching that error for
    the Excel/LibreOffice adapters degrade this engine identically instead of
    crashing the whole load. The same error is raised -- instead of a bare
    ``IndexError``/``ValueError``/``RuntimeError`` escaping -- for any
    malformed or mismatched surface the native extension returns, and for
    any other exception the PyO3 call itself raises (including a Rust panic,
    which PyO3 surfaces as a ``pyo3_runtime.PanicException`` that does not
    subclass ``Exception``): the native boundary must be total and
    degradable, never a crash of the whole load.
    """
    if not native_formula_available():
        raise FormulaEnrichmentError(
            "the 'native' formula engine was requested but a compatible "
            f"cadence-diff native engine is unavailable "
            f"({native_kernel.native_kernel_status().value})"
        )
    try:
        surface = formula_surface_report(data)
    except FormulaEnrichmentError:
        raise
    except (SystemExit, KeyboardInterrupt, GeneratorExit):
        raise
    except BaseException as exc:
        # Deliberately broad: the native extension is an untrusted FFI
        # boundary and any failure there (including a Rust panic) must
        # degrade the same way as a returned error, never crash the load.
        # Never log/include `exc`'s own text -- it could echo a path or
        # other adapter-internal detail; the exception class name is safe.
        raise FormulaEnrichmentError(
            "the cadence-diff native engine failed to decode this workbook's "
            f"formula surface ({type(exc).__name__})"
        ) from exc
    formulas: FormulaMap = {}
    formulas_r1c1: FormulaMap = {}
    for sheet in surface.sheets:
        _validate_sheet_surface(sheet)
        cells = {
            (row + 1, col + 1): f"={text}"
            for row, col, text in zip(
                sheet.cell_rows, sheet.cell_cols, sheet.cell_a1, strict=True
            )
            if text  # fail closed: an absent/failed render is never a bare "="
        }
        if cells:
            formulas[sheet.name] = cells
        r1c1_cells = _validated_r1c1_cells(sheet)
        if r1c1_cells:
            formulas_r1c1[sheet.name] = r1c1_cells
    names = [
        ExtractedDefinedName(name=item.name, target=item.target_a1)
        for item in surface.defined_names
        if item.target_a1 is not None
    ]
    defined_names, defined_names_complete = bounded_defined_names(names)
    version = _kernel_version() or "unknown"
    return FormulaExtraction(
        formulas=formulas,
        engine=f"native-biff12:{version}",
        detail="Formula text decoded by the native BIFF12 kernel",
        defined_names=defined_names,
        defined_names_complete=defined_names_complete,
        formulas_r1c1=formulas_r1c1,
    )


def restrict_formula_extraction(
    extraction: FormulaExtraction, coordinates: frozenset[tuple[str, int, int]]
) -> FormulaExtraction:
    """A copy of `extraction` with `.formulas` filtered to only `coordinates`.

    Used only by the native engine's compatibility mode (plan Criterion
    13(a)'s parity oracle): the native kernel decodes cells (e.g. dynamic-
    array spills) a legacy external engine cannot return text for at all, so
    proving native's TEXT is correct first requires restricting its broader
    coverage down to exactly what a legacy engine also covers, before
    comparing findings against a legacy-engine oracle. Never used for a real
    production run -- a real run wants the kernel's full coverage.
    """
    formulas: FormulaMap = {}
    for sheet, cells in extraction.formulas.items():
        kept = {
            (row, column): text
            for (row, column), text in cells.items()
            if (sheet, row, column) in coordinates
        }
        if kept:
            formulas[sheet] = kept
    formulas_r1c1: FormulaMap | None = None
    if extraction.formulas_r1c1 is not None:
        formulas_r1c1 = {}
        for sheet, cells in extraction.formulas_r1c1.items():
            kept_r1c1 = {
                (row, column): text
                for (row, column), text in cells.items()
                if (sheet, row, column) in coordinates
            }
            if kept_r1c1:
                formulas_r1c1[sheet] = kept_r1c1
    return FormulaExtraction(
        formulas=formulas,
        engine=extraction.engine,
        detail=f"{extraction.detail}; restricted to legacy-engine coordinates for compatibility",
        defined_names=extraction.defined_names,
        defined_names_complete=extraction.defined_names_complete,
        formulas_r1c1=formulas_r1c1,
    )
