"""Killable pywin32 worker for reading XLSB Formula2 text from desktop Excel."""

from __future__ import annotations

import json
import os
import sys
import threading
from collections.abc import Iterable
from contextlib import suppress
from pathlib import Path
from typing import Any


def _rectangles(coordinates: Iterable[tuple[int, int]]) -> list[tuple[int, int, int, int]]:
    """Merge horizontal formula runs with identical runs on adjacent rows."""
    rows: dict[int, list[int]] = {}
    for row, column in sorted(set(coordinates)):
        rows.setdefault(row, []).append(column)
    spans: dict[int, list[tuple[int, int]]] = {}
    for row, columns in rows.items():
        row_spans: list[tuple[int, int]] = []
        start = previous = columns[0]
        for column in columns[1:]:
            if column == previous + 1:
                previous = column
                continue
            row_spans.append((start, previous))
            start = previous = column
        row_spans.append((start, previous))
        spans[row] = row_spans

    rectangles: list[tuple[int, int, int, int]] = []
    active: dict[tuple[int, int], int] = {}
    previous_row: int | None = None
    for row in sorted(spans):
        if previous_row is None or row != previous_row + 1:
            rectangles.extend(
                (start_row, c1, previous_row or start_row, c2)
                for (c1, c2), start_row in active.items()
            )
            active.clear()
        current_spans = set(spans[row])
        for span, start_row in list(active.items()):
            if span not in current_spans:
                rectangles.append((start_row, span[0], previous_row or start_row, span[1]))
                del active[span]
        for span in current_spans:
            active.setdefault(span, row)
        previous_row = row
    if previous_row is not None:
        rectangles.extend(
            (start_row, c1, previous_row, c2)
            for (c1, c2), start_row in active.items()
        )
    return sorted(rectangles)


def _grid(value: object, rows: int, columns: int) -> list[list[object]]:
    if rows == columns == 1:
        return [[value]]
    if not isinstance(value, tuple | list):
        raise RuntimeError("Excel returned a scalar for a multi-cell formula range")
    raw = list(value)
    if rows == 1 and len(raw) == columns and not any(
        isinstance(item, tuple | list) for item in raw
    ):
        return [raw]
    if columns == 1 and len(raw) == rows and not any(
        isinstance(item, tuple | list) for item in raw
    ):
        return [[item] for item in raw]
    grid = [list(row) if isinstance(row, tuple | list) else [row] for row in raw]
    if len(grid) != rows or any(len(row) != columns for row in grid):
        raise RuntimeError("Excel returned an unexpected Formula2 array shape")
    return grid


def _write_json(path: Path, payload: dict[str, object]) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=True, separators=(",", ":")),
        encoding="utf-8",
    )
    os.replace(temporary, path)


def _watch_parent(parent_pid: int, excel_handle: Any) -> None:
    import win32api  # pyright: ignore[reportMissingModuleSource]
    import win32con  # pyright: ignore[reportMissingModuleSource]
    import win32event  # pyright: ignore[reportMissingModuleSource]

    try:
        parent_handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, parent_pid)
    except OSError:
        win32api.TerminateProcess(excel_handle, 1)
        os._exit(1)
    try:
        if win32event.WaitForSingleObject(parent_handle, win32event.INFINITE) == 0:
            try:
                win32api.TerminateProcess(excel_handle, 1)
            finally:
                os._exit(1)
    finally:
        win32api.CloseHandle(parent_handle)


def _excel_process(app: Any) -> tuple[int, Any, float]:
    import win32api  # pyright: ignore[reportMissingModuleSource]
    import win32con  # pyright: ignore[reportMissingModuleSource]
    import win32process  # pyright: ignore[reportMissingModuleSource]

    hwnd = int(app.Hwnd)
    if hwnd <= 0:
        raise RuntimeError("Excel did not expose an application window handle")
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    if pid <= 0:
        raise RuntimeError("Excel did not expose an owned process id")
    handle = win32api.OpenProcess(
        win32con.PROCESS_QUERY_INFORMATION
        | win32con.PROCESS_TERMINATE
        | win32con.SYNCHRONIZE,
        False,
        pid,
    )
    created = win32process.GetProcessTimes(handle)["CreationTime"].timestamp()
    return pid, handle, created


def _extract(request: dict[str, object]) -> dict[str, object]:
    import pythoncom  # pyright: ignore[reportMissingModuleSource]
    import pywintypes  # pyright: ignore[reportMissingModuleSource]
    import win32api  # pyright: ignore[reportMissingModuleSource]
    import win32com.client  # pyright: ignore[reportMissingModuleSource]

    input_path = Path(str(request["input_path"]))
    status_path = Path(str(request["status_path"]))
    raw_parent_pid = request["parent_pid"]
    raw_expected_count = request["formula_count"]
    if (
        not isinstance(raw_parent_pid, int)
        or isinstance(raw_parent_pid, bool)
        or not isinstance(raw_expected_count, int)
        or isinstance(raw_expected_count, bool)
    ):
        raise RuntimeError("worker process identity request is invalid")
    parent_pid = raw_parent_pid
    expected_count = raw_expected_count
    formula_cells = request["formula_cells"]
    if not isinstance(formula_cells, dict):
        raise RuntimeError("worker formula_cells request is invalid")

    pythoncom.CoInitialize()
    app = workbook = excel_handle = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        pid, excel_handle, created = _excel_process(app)
        _write_json(status_path, {"pid": pid, "created": created})
        threading.Thread(
            target=_watch_parent,
            args=(parent_pid, excel_handle),
            daemon=True,
        ).start()

        app.Visible = False
        app.AutomationSecurity = 3
        app.EnableEvents = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        app.Calculation = -4135
        app.CalculateBeforeSave = False
        app.Interactive = False
        app.ScreenUpdating = False
        with suppress(pywintypes.com_error):
            app.AutoRecover.Enabled = False

        workbook = app.Workbooks.Open(
            Filename=str(input_path),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            Notify=False,
            AddToMru=False,
            CorruptLoad=0,
        )
        version = str(app.Version)
        build = str(app.Build)
        formulas: dict[str, list[dict[str, object]]] = {}
        extracted_count = 0
        for sheet_name, raw_coordinates in formula_cells.items():
            if not isinstance(sheet_name, str) or not isinstance(raw_coordinates, list):
                raise RuntimeError("worker formula coordinate request is invalid")
            coordinates: list[tuple[int, int]] = []
            for item in raw_coordinates:
                if (
                    not isinstance(item, list)
                    or len(item) != 2
                    or not all(
                        isinstance(value, int) and not isinstance(value, bool)
                        for value in item
                    )
                ):
                    raise RuntimeError("worker formula coordinate request is invalid")
                coordinates.append((item[0], item[1]))
            sheet = workbook.Worksheets(sheet_name)
            cells: list[dict[str, object]] = []
            for row_1, col_1, row_2, col_2 in _rectangles(coordinates):
                target = sheet.Range(sheet.Cells(row_1, col_1), sheet.Cells(row_2, col_2))
                if target.HasFormula is not True:
                    raise RuntimeError(
                        f"Excel did not confirm every requested formula in {sheet_name}"
                    )
                try:
                    raw_grid = target.Formula2
                except pywintypes.com_error as exc:
                    raise RuntimeError(
                        "this Excel build does not provide reliable Formula2 access"
                    ) from exc
                grid = _grid(raw_grid, row_2 - row_1 + 1, col_2 - col_1 + 1)
                for row_offset, values in enumerate(grid):
                    for column_offset, formula in enumerate(values):
                        if not isinstance(formula, str) or not formula.startswith("="):
                            raise RuntimeError(
                                f"Excel returned invalid Formula2 text in {sheet_name}"
                            )
                        cells.append(
                            {
                                "row": row_1 + row_offset,
                                "column": col_1 + column_offset,
                                "formula": formula,
                            }
                        )
                        extracted_count += 1
            if cells:
                formulas[sheet_name] = cells
        if extracted_count != expected_count:
            raise RuntimeError(
                f"Excel extracted {extracted_count} formulas; expected {expected_count}"
            )
        return {
            "schema_version": 1,
            "ok": True,
            "engine": f"excel:{version}:{build}:formula2",
            "detail": f"Formula2 text extracted by Microsoft Excel {version} build {build}",
            "formulas": formulas,
        }
    finally:
        if workbook is not None:
            with suppress(Exception):
                workbook.Close(SaveChanges=False)
        if app is not None:
            with suppress(Exception):
                app.Quit()
        if excel_handle is not None:
            win32api.CloseHandle(excel_handle)
        pythoncom.CoUninitialize()


def main() -> int:
    result_path: Path | None = None
    try:
        request = json.loads(sys.stdin.read())
        if not isinstance(request, dict) or request.get("schema_version") != 1:
            raise RuntimeError("invalid Excel formula worker request")
        result_path = Path(str(request["result_path"]))
        payload = _extract(request)
    except Exception as exc:
        payload = {
            "schema_version": 1,
            "ok": False,
            "error": f"{type(exc).__name__}: {exc}",
        }
    if result_path is None:
        return 2
    _write_json(result_path, payload)
    return 0 if payload.get("ok") is True else 1


if __name__ == "__main__":
    raise SystemExit(main())
