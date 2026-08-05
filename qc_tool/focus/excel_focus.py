"""Excel worksheet and range focus inside the short-lived helper.

The navigation core is duck-typed over the object model, so it is exercised with
fake COM objects on any platform. It activates a window, activates an already
visible worksheet, and scrolls to a bounded range. It never opens, saves,
recalculates, refreshes, closes, or quits anything, and it never touches a
global Office setting.
"""

from __future__ import annotations

import contextlib
import sys
from collections.abc import Callable

from qc_tool.focus.discovery import OpenDocument, as_bool, as_count, read_attribute
from qc_tool.focus.locator import normalize_address
from qc_tool.focus.protocol import FocusOutcome, path_digest

#: ``xlSheetVisible``. Focus refuses anything else and never unhides a sheet.
XL_SHEET_VISIBLE = -1

Foreground = Callable[[int], bool]


def _worksheet(workbook: object, name: str) -> object | None:
    sheets = read_attribute(workbook, "Worksheets")
    total = as_count(read_attribute(sheets, "Count"))
    for index in range(1, total + 1):
        sheet: object | None = None
        with contextlib.suppress(Exception):
            sheet = sheets.Item(index)  # type: ignore[attr-defined]
        if sheet is None:
            continue
        if read_attribute(sheet, "Name") == name:
            return sheet
    return None


def navigate_excel(
    window_object: object,
    *,
    sheet_name: str,
    address: str | None,
    expected_path_digest: str,
    salt: bytes,
    window_handle: int,
    set_foreground: Foreground,
) -> str:
    """Activate one exact worksheet and range in one exact bound workbook."""
    workbook = read_attribute(window_object, "Parent")
    application = read_attribute(window_object, "Application")
    if workbook is None or application is None:
        return FocusOutcome.BOUND_DOCUMENT_NOT_FOUND.value
    full_name = read_attribute(workbook, "FullName")
    if not isinstance(full_name, str) or path_digest(full_name, salt) != (
        expected_path_digest
    ):
        return FocusOutcome.BOUND_DOCUMENT_NOT_FOUND.value
    if as_count(read_attribute(read_attribute(workbook, "Windows"), "Count")) != 1:
        return FocusOutcome.TARGET_WINDOW_MISSING.value
    saved_before = as_bool(read_attribute(workbook, "Saved"))
    sheet = _worksheet(workbook, sheet_name)
    if sheet is None:
        return FocusOutcome.TARGET_SHEET_MISSING.value
    if read_attribute(sheet, "Visible") != XL_SHEET_VISIBLE:
        return FocusOutcome.TARGET_SHEET_HIDDEN.value
    target: object | None = None
    if address is not None:
        normalized = normalize_address(address)
        if normalized is None:
            return FocusOutcome.TARGET_ADDRESS_INVALID.value
        try:
            target = sheet.Range(normalized)  # type: ignore[attr-defined]
        except Exception:
            return FocusOutcome.TARGET_ADDRESS_INVALID.value
        if target is None:
            return FocusOutcome.TARGET_ADDRESS_INVALID.value
    try:
        window_object.Activate()  # type: ignore[attr-defined]
        sheet.Activate()  # type: ignore[attr-defined]
        if target is not None:
            # Application-level navigation, but only with a Range already taken
            # from the exact bound workbook.
            application.Goto(target, True)  # type: ignore[attr-defined]
    except Exception:
        return FocusOutcome.HELPER_FAILED.value
    if (
        saved_before is True
        and as_bool(read_attribute(workbook, "Saved")) is False
    ):
        return FocusOutcome.SIDE_EFFECT_DETECTED.value
    if not set_foreground(window_handle):
        # Windows can deny foreground activation. Never retarget another window.
        return FocusOutcome.FOCUSED_WITHOUT_FOREGROUND.value
    return FocusOutcome.FOCUSED.value


def focus_document(document: OpenDocument, request: dict[str, object]) -> str:
    """Windows adapter: re-acquire the bound window and navigate it."""
    if sys.platform != "win32":
        return FocusOutcome.UNSUPPORTED_PLATFORM.value
    from qc_tool.focus import win32_office

    handles = sorted(set(document.visible_window_handles))
    if len(handles) != 1:
        return FocusOutcome.TARGET_WINDOW_MISSING.value
    sheet_name = request.get("sheet")
    salt_hex = request.get("path_salt")
    expected = request.get("expected_path_digest")
    if (
        not isinstance(sheet_name, str)
        or not sheet_name
        or not isinstance(salt_hex, str)
        or not isinstance(expected, str)
    ):
        return FocusOutcome.INVALID_REQUEST.value
    raw_address = request.get("address")
    address = raw_address if isinstance(raw_address, str) and raw_address else None
    if not document.object_model_window_handle:
        return FocusOutcome.TARGET_WINDOW_MISSING.value
    window_object = win32_office.object_from_window(
        document.object_model_window_handle
    )
    if window_object is None:
        return FocusOutcome.TARGET_WINDOW_MISSING.value
    return navigate_excel(
        window_object,
        sheet_name=sheet_name,
        address=address,
        expected_path_digest=expected,
        salt=bytes.fromhex(salt_hex),
        window_handle=handles[0],
        set_foreground=win32_office.set_foreground_window,
    )
