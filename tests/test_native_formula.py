"""Tests for the native formula surface adapter
(`qc_tool/io/native_formula.py`) -- B2's "definitions + per-cell ids + names
surface" deliverable, plus B3's `FormulaExtraction` adapter
(`extract_formulas_with_native_kernel`) and compatibility-mode restriction
(`restrict_formula_extraction`).
"""

from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

import qc_tool.io.native_formula as native_formula_module
import qc_tool.io.native_kernel as native_kernel_module
from qc_tool.io.formula_enrichment import FormulaEnrichmentError, FormulaExtraction
from qc_tool.io.native_formula import (
    DefinedNameSurface,
    FormulaSheetSurface,
    WorkbookFormulaSurface,
    extract_formulas_with_native_kernel,
    formula_surface_report,
    native_adapter_fingerprint,
    native_formula_available,
    restrict_formula_extraction,
)
from qc_tool.io.xlsb_formula import XlsbFormulaScan

FIXTURE = Path(__file__).parent / "fixtures" / "generated" / "current.xlsb"


def _available_native_module() -> SimpleNamespace:
    return SimpleNamespace(
        __version__="2.0.0",
        __kernel_api_version__=1,
        __native_available__=True,
        raw_values_report=lambda data: [],
        formula_surface_report=lambda data: ([], []),
        formula_r1c1_report=lambda data: [],
        formula_delta_batch=lambda pairs: [],
    )


@pytest.mark.skipif(
    not native_formula_available(),
    reason="native/cadence_diff_native/ not built in this environment (optional accelerator)",
)
def test_formula_surface_report_returns_a_typed_surface() -> None:
    surface = formula_surface_report(FIXTURE.read_bytes())
    assert isinstance(surface.sheets, tuple)
    assert isinstance(surface.defined_names, tuple)
    # This fixture has 0 formulas (a values-only fixture) but must still
    # decode cleanly to an empty-but-valid surface, never raise.
    for sheet in surface.sheets:
        lengths = {
            len(sheet.cell_rows),
            len(sheet.cell_cols),
            len(sheet.cell_a1),
            len(sheet.cell_definition_id),
        }
        assert len(lengths) == 1
        for def_id in sheet.cell_definition_id:
            assert 0 <= def_id < len(sheet.definition_r1c1)


def test_formula_surface_report_raises_a_clear_error_when_unavailable(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(native_kernel_module, "_native_module", None)
    with pytest.raises(RuntimeError, match=r"native engine unavailable \(missing\)"):
        formula_surface_report(FIXTURE.read_bytes())


def _fake_surface() -> WorkbookFormulaSurface:
    """A deterministic surface standing in for the Rust extension's real
    output -- 0-based coordinates, bare token-stream A1 text with no leading
    ``=`` (both confirmed empirically against the built extension).
    """
    return WorkbookFormulaSurface(
        sheets=(
            FormulaSheetSurface(
                name="Data",
                cell_rows=(0, 1),
                cell_cols=(0, 1),
                cell_a1=("A2+1", "SUM(A1:A2)"),
                cell_definition_id=(0, 1),
                definition_r1c1=("RC[1]+1", "SUM(R[-1]C:R[-2]C)"),
            ),
        ),
        defined_names=(
            DefinedNameSurface(name="MyRange", target_a1="Data!$A$1:$A$2"),
            DefinedNameSurface(name="Undecoded", target_a1=None),
        ),
    )


def _scan() -> XlsbFormulaScan:
    return XlsbFormulaScan(formula_cells={"Data": frozenset({(1, 1), (2, 2)})})


def test_extract_formulas_with_native_kernel_prefixes_and_converts_coordinates(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The kernel's bare, 0-based surface must become the same shape the
    Excel/LibreOffice adapters return: ``=``-prefixed text at 1-based
    coordinates (Criterion 9 depends on every consumer seeing one contract
    regardless of engine).
    """
    monkeypatch.setattr(
        native_kernel_module, "_native_module", _available_native_module()
    )
    monkeypatch.setattr(
        native_formula_module, "formula_surface_report", lambda data: _fake_surface()
    )
    monkeypatch.setattr(native_kernel_module, "native_distribution_version", lambda: "9.9.9")

    extraction = extract_formulas_with_native_kernel(b"unused", _scan())

    assert extraction.formulas == {
        "Data": {(1, 1): "=A2+1", (2, 2): "=SUM(A1:A2)"},
    }
    assert extraction.engine == "native-biff12:9.9.9"
    assert all(formula.startswith("=") for formula in extraction.formulas["Data"].values())
    # The undecoded name (target_a1=None) is dropped, never guessed.
    assert [item.name for item in extraction.defined_names] == ["MyRange"]
    assert extraction.defined_names[0].target == "Data!$A$1:$A$2"


def test_extract_formulas_with_native_kernel_raises_formula_enrichment_error_when_unavailable(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Must raise `FormulaEnrichmentError` (not a bare `RuntimeError`) so
    `qc_tool/io/loader.py`'s existing degrade-gracefully `except` clause
    catches it identically to the Excel/LibreOffice adapters.
    """
    monkeypatch.setattr(native_kernel_module, "_native_module", None)
    with pytest.raises(FormulaEnrichmentError, match=r"native.*formula engine"):
        extract_formulas_with_native_kernel(b"unused", _scan())


def test_extract_formulas_with_native_kernel_also_populates_r1c1(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """B5: the adapter must also expose canonical R1C1 per cell (this
    fixture's placeholder `definition_r1c1` values do not correspond to a
    real `to_r1c1()` computation from `cell_a1`, so this exercises the
    validate-and-fallback path -- see the dedicated `_validated_r1c1_cells`
    tests below for the trusted-kernel-value path).
    """
    monkeypatch.setattr(
        native_kernel_module, "_native_module", _available_native_module()
    )
    monkeypatch.setattr(
        native_formula_module, "formula_surface_report", lambda data: _fake_surface()
    )
    monkeypatch.setattr(native_kernel_module, "native_distribution_version", lambda: "9.9.9")

    extraction = extract_formulas_with_native_kernel(b"unused", _scan())

    assert extraction.formulas_r1c1 is not None
    assert set(extraction.formulas_r1c1["Data"]) == {(1, 1), (2, 2)}
    assert all(
        text.startswith("=") for text in extraction.formulas_r1c1["Data"].values()
    )


def test_native_adapter_fingerprint_none_when_unavailable(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(native_kernel_module, "_native_module", None)
    assert native_adapter_fingerprint() is None


def test_restrict_formula_extraction_keeps_only_given_coordinates() -> None:
    extraction = FormulaExtraction(
        formulas={
            "Data": {(1, 1): "=A2+1", (2, 2): "=SUM(A1:A2)"},
            "Other": {(1, 1): "=1"},
        },
        engine="native-biff12:0.1.0",
        detail="test",
    )

    restricted = restrict_formula_extraction(
        extraction, frozenset({("Data", 1, 1)})
    )

    assert restricted.formulas == {"Data": {(1, 1): "=A2+1"}}
    assert restricted.engine == extraction.engine  # engine identity unchanged
    assert "restricted to legacy-engine coordinates" in restricted.detail


def test_restrict_formula_extraction_drops_a_sheet_left_with_no_cells() -> None:
    extraction = FormulaExtraction(
        formulas={"Data": {(1, 1): "=A2+1"}},
        engine="native-biff12:0.1.0",
        detail="test",
    )

    restricted = restrict_formula_extraction(extraction, frozenset())

    assert restricted.formulas == {}


def test_restrict_formula_extraction_also_restricts_formulas_r1c1() -> None:
    extraction = FormulaExtraction(
        formulas={"Data": {(1, 1): "=A2+1", (2, 2): "=SUM(A1:A2)"}},
        engine="native-biff12:0.1.0",
        detail="test",
        formulas_r1c1={"Data": {(1, 1): "=RC[1]+1", (2, 2): "=SUM(R[-1]C:R[-2]C)"}},
    )

    restricted = restrict_formula_extraction(
        extraction, frozenset({("Data", 1, 1)})
    )

    assert restricted.formulas_r1c1 == {"Data": {(1, 1): "=RC[1]+1"}}


def test_restrict_formula_extraction_handles_no_formulas_r1c1() -> None:
    """Excel/LibreOffice extractions never set `formulas_r1c1`; restriction
    must not invent one."""
    extraction = FormulaExtraction(
        formulas={"Data": {(1, 1): "=A2+1"}},
        engine="test-engine",
        detail="test",
    )

    restricted = restrict_formula_extraction(extraction, frozenset({("Data", 1, 1)}))

    assert restricted.formulas_r1c1 is None


# --- B5: per-definition R1C1 validate-and-fallback --------------------------


def test_validated_r1c1_cells_trusts_a_kernel_value_that_matches_to_r1c1() -> None:
    from qc_tool.io.native_formula import _validated_r1c1_cells

    # host (row=1, col=1); "A2" is row 2 col A -> row offset +1, col offset 0
    # -> to_r1c1("=A2", 1, 1) == "=R[1]C", matching this kernel value exactly.
    sheet = FormulaSheetSurface(
        name="Data",
        cell_rows=(0,),
        cell_cols=(0,),
        cell_a1=("A2",),
        cell_definition_id=(0,),
        definition_r1c1=("R[1]C",),
    )

    assert _validated_r1c1_cells(sheet) == {(1, 1): "=R[1]C"}


def test_validated_r1c1_cells_falls_back_to_to_r1c1_on_a_mismatch() -> None:
    from qc_tool.io.native_formula import _validated_r1c1_cells

    # Same cell as above, but the kernel's claimed R1C1 is deliberately wrong.
    sheet = FormulaSheetSurface(
        name="Data",
        cell_rows=(0,),
        cell_cols=(0,),
        cell_a1=("A2",),
        cell_definition_id=(0,),
        definition_r1c1=("R[99]C",),
    )

    assert _validated_r1c1_cells(sheet) == {(1, 1): "=R[1]C"}


def test_validated_r1c1_cells_validates_once_per_definition_not_per_cell(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import qc_tool.excel.formulas as formulas_module
    from qc_tool.io.native_formula import _validated_r1c1_cells

    calls: list[tuple[str, int, int]] = []
    real_to_r1c1 = formulas_module.to_r1c1

    def counting_to_r1c1(formula: str, row: int, col: int) -> str:
        calls.append((formula, row, col))
        return real_to_r1c1(formula, row, col)

    monkeypatch.setattr(formulas_module, "to_r1c1", counting_to_r1c1)
    # Three followers of one shared definition at rows 2/3/4 (1-based),
    # column B, each referencing the cell one row up and one column left --
    # the same relative shape for every follower.
    sheet = FormulaSheetSurface(
        name="Data",
        cell_rows=(1, 2, 3),
        cell_cols=(1, 1, 1),
        cell_a1=("A1", "A2", "A3"),
        cell_definition_id=(0, 0, 0),
        definition_r1c1=("R[-1]C[-1]",),
    )

    result = _validated_r1c1_cells(sheet)

    assert len(calls) == 1  # validated once for the whole shared definition
    assert result == {
        (2, 2): "=R[-1]C[-1]",
        (3, 2): "=R[-1]C[-1]",
        (4, 2): "=R[-1]C[-1]",
    }

