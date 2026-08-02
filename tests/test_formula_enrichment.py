"""Formula enrichment parity, merging, and sandbox-adapter behavior."""

import json
import os
import subprocess
from pathlib import Path
from typing import Any

import pytest
from openpyxl import Workbook

import qc_tool.io.libreoffice_formula as libreoffice_formula_module
from qc_tool.io.excel_formula import _load_worker_result, extract_formulas_with_excel
from qc_tool.io.excel_formula_worker import (
    _grid,
    _rectangles,
    _set_manual_calculation,
)
from qc_tool.io.formula_enrichment import (
    FormulaEnrichmentError,
    FormulaExtraction,
    merge_formula_extraction,
)
from qc_tool.io.libreoffice_formula import (
    _convert_with_libreoffice,
    _extract_formulas_with_converter,
)
from qc_tool.io.model import CellRecord, SheetSnapshot, WorkbookSnapshot
from qc_tool.io.xlsb_formula import XlsbFormulaScan


def _snapshot() -> WorkbookSnapshot:
    return WorkbookSnapshot(
        source_name="source.xlsb",
        file_format="xlsb",
        formulas_available=False,
        styles_available=False,
        formula_presence_available=True,
        sheets=[
            SheetSnapshot(
                name="Data",
                visibility="visible",
                max_row=1,
                max_column=1,
                cells={(1, 1): CellRecord(1, 1, 42.0, is_formula=True)},
            )
        ],
    )


def _scan(*, risky: tuple[str, ...] = ()) -> XlsbFormulaScan:
    return XlsbFormulaScan(
        formula_cells={"Data": frozenset({(1, 1)})},
        risky_features=risky,
    )


def test_merge_preserves_original_cached_value() -> None:
    snapshot = _snapshot()
    extraction = FormulaExtraction(
        formulas={"Data": {(1, 1): "=20+22"}},
        engine="test-engine",
        detail="test extraction",
    )

    merge_formula_extraction(snapshot, _scan(), extraction)

    cell = snapshot.sheet("Data").cells[(1, 1)]
    assert cell.value == 42.0
    assert cell.formula == "=20+22" and cell.has_formula
    assert snapshot.formulas_available
    assert snapshot.formula_source == "test-engine"


def test_coordinate_mismatch_fails_before_mutation() -> None:
    snapshot = _snapshot()
    extraction = FormulaExtraction(
        formulas={"Data": {(1, 2): "=20+22"}},
        engine="test-engine",
        detail="test extraction",
    )

    with pytest.raises(FormulaEnrichmentError, match="coordinate mismatch"):
        merge_formula_extraction(snapshot, _scan(), extraction)

    assert snapshot.sheet("Data").cells[(1, 1)].formula is None
    assert not snapshot.formulas_available


def test_libreoffice_adapter_reads_formula_only_and_checks_parity() -> None:
    def fake_converter(work_dir: Path, timeout: float) -> tuple[str, str]:
        assert timeout == 12.0
        if os.name == "posix":
            assert (work_dir / "input.xlsb").stat().st_mode & 0o777 == 0o600
        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "Data"
        sheet["A1"] = "=20+22"
        workbook.save(work_dir / "output" / "input.xlsx")
        return "fake-libreoffice", "sandbox test"

    extraction = _extract_formulas_with_converter(
        b"decrypted xlsb bytes",
        _scan(),
        timeout=12.0,
        converter=fake_converter,
    )

    assert extraction.formulas == {"Data": {(1, 1): "=20+22"}}
    assert extraction.engine == "fake-libreoffice"


def test_libreoffice_adapter_rejects_risky_content_before_conversion() -> None:
    called = False

    def fake_converter(work_dir: Path, timeout: float) -> tuple[str, str]:
        nonlocal called
        called = True
        return "unused", "unused"

    with pytest.raises(FormulaEnrichmentError, match="VBA project"):
        _extract_formulas_with_converter(
            b"xlsb",
            _scan(risky=("VBA project",)),
            timeout=300.0,
            converter=fake_converter,
        )

    assert not called


def test_libreoffice_adapter_rejects_formula_coordinate_mismatch() -> None:
    def fake_converter(work_dir: Path, timeout: float) -> tuple[str, str]:
        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "Data"
        sheet["B1"] = "=20+22"
        workbook.save(work_dir / "output" / "input.xlsx")
        return "fake-libreoffice", "mismatch test"

    with pytest.raises(FormulaEnrichmentError, match="coordinate mismatch"):
        _extract_formulas_with_converter(
            b"xlsb",
            _scan(),
            timeout=300.0,
            converter=fake_converter,
        )


@pytest.mark.skipif(os.name != "posix", reason="POSIX process-group cleanup")
def test_libreoffice_timeout_race_still_returns_enrichment_error(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    class FinishedDuringTimeout:
        pid = 12345
        returncode = None

        def communicate(self, timeout: float | None = None) -> tuple[str, str]:
            if timeout is not None:
                raise subprocess.TimeoutExpired("libreoffice", timeout)
            self.returncode = 0
            return "", ""

    monkeypatch.setattr(
        libreoffice_formula_module,
        "_sandbox_command",
        lambda work_dir: (["libreoffice"], "/usr/bin/libreoffice"),
    )
    monkeypatch.setattr(
        libreoffice_formula_module.subprocess,
        "run",
        lambda *args, **kwargs: subprocess.CompletedProcess(
            args[0], 0, "LibreOffice test\n", ""
        ),
    )
    monkeypatch.setattr(
        libreoffice_formula_module.subprocess,
        "Popen",
        lambda *args, **kwargs: FinishedDuringTimeout(),
    )
    monkeypatch.setattr(
        libreoffice_formula_module.os,
        "killpg",
        lambda *args: (_ for _ in ()).throw(ProcessLookupError()),
    )

    with pytest.raises(FormulaEnrichmentError, match="exceeded"):
        _convert_with_libreoffice(tmp_path, 1.0)


def test_excel_adapter_protocol_is_testable_without_windows() -> None:
    def fake_worker(
        work_dir: Path, request: dict[str, object], timeout: float
    ) -> FormulaExtraction:
        assert timeout == 7.0
        assert request["formula_count"] == 1
        assert request["formula_cells"] == {"Data": [[1, 1]]}
        assert (work_dir / "input.xlsb").read_bytes() == b"xlsb"
        return FormulaExtraction(
            formulas={"Data": {(1, 1): "=20+22"}},
            engine="excel:test:formula2",
            detail="test Excel worker",
        )

    extraction = extract_formulas_with_excel(
        b"xlsb", _scan(), timeout=7.0, worker_runner=fake_worker
    )

    assert extraction.formulas == {"Data": {(1, 1): "=20+22"}}


def test_excel_formula_rectangles_and_shapes() -> None:
    assert _rectangles({(1, 1), (1, 2), (2, 1), (2, 2), (4, 3)}) == [
        (1, 1, 2, 2),
        (4, 3, 4, 3),
    ]
    assert _grid("=A1", 1, 1) == [["=A1"]]
    assert _grid(("=A1", "=B1"), 1, 2) == [["=A1", "=B1"]]
    assert _grid((("=A1",), ("=A2",)), 2, 1) == [["=A1"], ["=A2"]]


def test_excel_manual_calculation_retries_with_a_guard_workbook() -> None:
    class FakeComError(Exception):
        pass

    class Guard:
        closed = False

        def Close(self, *, SaveChanges: bool) -> None:
            assert SaveChanges is False
            self.closed = True

    class Workbooks:
        def __init__(self, app: Any) -> None:
            self.app = app
            self.guard = Guard()

        def Add(self) -> Guard:
            self.app.workbook_exists = True
            return self.guard

    class App:
        def __init__(self) -> None:
            self.workbook_exists = False
            self.calculation = 0
            self.Workbooks = Workbooks(self)

        @property
        def Calculation(self) -> int:
            return self.calculation

        @Calculation.setter
        def Calculation(self, value: int) -> None:
            if not self.workbook_exists:
                raise FakeComError
            self.calculation = value

    app = App()

    guard = _set_manual_calculation(app, FakeComError)

    assert guard is app.Workbooks.guard
    assert app.Calculation == -4135
    assert not app.Workbooks.guard.closed


def test_excel_manual_calculation_retry_stays_fail_closed() -> None:
    class FakeComError(Exception):
        pass

    class Guard:
        closed = False

        def Close(self, *, SaveChanges: bool) -> None:
            assert SaveChanges is False
            self.closed = True

    class Workbooks:
        def __init__(self) -> None:
            self.guard = Guard()

        def Add(self) -> Guard:
            return self.guard

    class App:
        def __init__(self) -> None:
            self.Workbooks = Workbooks()

        @property
        def Calculation(self) -> int:
            return 0

        @Calculation.setter
        def Calculation(self, value: int) -> None:
            raise FakeComError

    app = App()

    with pytest.raises(FakeComError):
        _set_manual_calculation(app, FakeComError)

    assert app.Workbooks.guard.closed


@pytest.mark.parametrize(
    "payload",
    [
        {"schema_version": 2, "ok": True},
        {"schema_version": 1, "ok": False, "error": "Excel unavailable"},
        {"schema_version": 1, "ok": True, "formulas": {}, "engine": 1, "detail": "x"},
    ],
)
def test_excel_worker_result_validation_fails_closed(
    payload: dict[str, object], tmp_path: Path
) -> None:
    path = tmp_path / "result.json"
    path.write_text(json.dumps(payload), encoding="utf-8")

    with pytest.raises(FormulaEnrichmentError):
        _load_worker_result(path)
