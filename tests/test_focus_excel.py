"""Step 5 contracts: exact Excel worksheet and range focus."""

import pytest

from qc_tool.focus.discovery import FocusApplication, OpenDocument
from qc_tool.focus.excel_focus import XL_SHEET_VISIBLE, focus_document, navigate_excel
from qc_tool.focus.protocol import FocusOutcome, path_digest

SALT = b"\x11\x22\x33"
FULL_NAME = r"C:\work\current.xlsx"
DIGEST = path_digest(FULL_NAME, SALT)


class FakeRange:
    def __init__(self, address: str) -> None:
        self.address = address


class FakeSheet:
    def __init__(self, name: str, *, visible: int = XL_SHEET_VISIBLE) -> None:
        self.Name = name
        self.Visible = visible
        self.activated = False
        self.requested: list[str] = []

    def Activate(self) -> None:
        self.activated = True

    def Range(self, address: str) -> FakeRange:
        self.requested.append(address)
        if address.startswith("ZZ"):
            raise RuntimeError("no such range")
        return FakeRange(address)


class FakeCollection:
    def __init__(self, items: list[object]) -> None:
        self._items = items
        self.Count = len(items)

    def Item(self, index: int) -> object:
        return self._items[index - 1]


class FakeApplication:
    def __init__(self) -> None:
        self.goto: list[tuple[object, bool]] = []

    def Goto(self, reference: object, scroll: bool) -> None:
        self.goto.append((reference, scroll))


class FakeWorkbook:
    def __init__(
        self,
        *,
        full_name: str = FULL_NAME,
        sheets: list[FakeSheet] | None = None,
        windows: int = 1,
        saved: bool = True,
    ) -> None:
        self.FullName = full_name
        self.Worksheets = FakeCollection(
            list(sheets if sheets is not None else [FakeSheet("Summary")])
        )
        self.Windows = FakeCollection([object()] * windows)
        self.Saved = saved


class FakeWindow:
    def __init__(self, workbook: FakeWorkbook, application: FakeApplication) -> None:
        self.Parent = workbook
        self.Application = application
        self.activated = False

    def Activate(self) -> None:
        self.activated = True


def _navigate(window: FakeWindow, *, address: str | None = "B5", foreground=lambda _h: True):
    return navigate_excel(
        window,
        sheet_name="Summary",
        address=address,
        expected_path_digest=DIGEST,
        salt=SALT,
        window_handle=101,
        set_foreground=foreground,
    )


def _window(**kwargs) -> tuple[FakeWindow, FakeApplication, FakeWorkbook]:
    application = FakeApplication()
    workbook = FakeWorkbook(**kwargs)
    return FakeWindow(workbook, application), application, workbook


def test_focus_activates_the_window_sheet_and_range() -> None:
    window, application, workbook = _window()
    assert _navigate(window) == FocusOutcome.FOCUSED.value
    sheet = workbook.Worksheets.Item(1)
    assert window.activated
    assert isinstance(sheet, FakeSheet) and sheet.activated
    assert sheet.requested == ["B5"]
    assert application.goto and application.goto[0][1] is True


def test_sheet_only_target_activates_without_a_range() -> None:
    window, application, workbook = _window()
    assert _navigate(window, address=None) == FocusOutcome.FOCUSED.value
    sheet = workbook.Worksheets.Item(1)
    assert isinstance(sheet, FakeSheet) and sheet.requested == []
    assert application.goto == []


def test_wrong_document_is_refused_even_with_the_right_window() -> None:
    window, _application, _workbook = _window(full_name=r"C:\work\other.xlsx")
    assert _navigate(window) == FocusOutcome.BOUND_DOCUMENT_NOT_FOUND.value


def test_missing_sheet_refuses() -> None:
    window, _application, _workbook = _window(sheets=[FakeSheet("Detail")])
    assert _navigate(window) == FocusOutcome.TARGET_SHEET_MISSING.value


def test_hidden_sheet_is_never_unhidden() -> None:
    sheet = FakeSheet("Summary", visible=0)
    window, _application, _workbook = _window(sheets=[sheet])
    assert _navigate(window) == FocusOutcome.TARGET_SHEET_HIDDEN.value
    assert sheet.Visible == 0
    assert not sheet.activated


@pytest.mark.parametrize("address", ["not-an-address", "A:B", "A1048577", "XFE1"])
def test_malformed_or_oversized_address_refuses(address: str) -> None:
    window, application, _workbook = _window()
    assert _navigate(window, address=address) == (
        FocusOutcome.TARGET_ADDRESS_INVALID.value
    )
    assert application.goto == []


def test_a_range_office_rejects_refuses() -> None:
    window, application, _workbook = _window()
    assert _navigate(window, address="ZZ9") == FocusOutcome.TARGET_ADDRESS_INVALID.value
    assert application.goto == []


def test_multiple_workbook_windows_refuse() -> None:
    window, application, _workbook = _window(windows=2)
    assert _navigate(window) == FocusOutcome.TARGET_WINDOW_MISSING.value
    assert application.goto == []


def test_foreground_denial_is_partial_success_not_a_retarget() -> None:
    window, application, _workbook = _window()
    outcome = _navigate(window, foreground=lambda _h: False)
    assert outcome == FocusOutcome.FOCUSED_WITHOUT_FOREGROUND.value
    assert len(application.goto) == 1


def test_a_document_that_becomes_dirty_during_navigation_is_a_side_effect() -> None:
    window, _application, workbook = _window()

    class DirtyingSheet(FakeSheet):
        def Activate(self) -> None:
            super().Activate()
            workbook.Saved = False

    workbook.Worksheets = FakeCollection([DirtyingSheet("Summary")])
    assert _navigate(window) == FocusOutcome.SIDE_EFFECT_DETECTED.value


def test_dirty_before_navigation_is_not_reported_as_a_side_effect() -> None:
    window, _application, _workbook = _window(saved=False)
    assert _navigate(window) == FocusOutcome.FOCUSED.value


def test_focus_document_refuses_off_windows() -> None:
    document = OpenDocument(
        application=FocusApplication.EXCEL,
        process_id=1,
        process_created=1.0,
        windows_session_id=1,
        full_name=FULL_NAME,
        window_count=1,
        visible_window_count=1,
        visible_window_handles=(101,),
        saved=True,
        autosave=False,
    )
    assert focus_document(document, {}) == FocusOutcome.UNSUPPORTED_PLATFORM.value
