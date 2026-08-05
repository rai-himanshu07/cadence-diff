"""Step 6 contracts: exact PowerPoint slide focus."""

import pytest

from qc_tool.focus.discovery import FocusApplication, OpenDocument
from qc_tool.focus.ppt_focus import (
    MSO_TRUE,
    PP_VIEW_NORMAL,
    focus_document,
    navigate_powerpoint,
)
from qc_tool.focus.protocol import FocusOutcome, path_digest

SALT = b"\x44\x55"
FULL_NAME = r"C:\work\deck.pptx"
DIGEST = path_digest(FULL_NAME, SALT)


class FakeTransition:
    def __init__(self, hidden: int = 0) -> None:
        self.Hidden = hidden


class FakeSlide:
    def __init__(self, *, hidden: bool = False) -> None:
        self.SlideShowTransition = FakeTransition(MSO_TRUE if hidden else 0)


class FakeCollection:
    def __init__(self, items: list[object]) -> None:
        self._items = items
        self.Count = len(items)

    def Item(self, index: int) -> object:
        return self._items[index - 1]


class FakeView:
    def __init__(self) -> None:
        self.visited: list[int] = []

    def GotoSlide(self, index: int) -> None:
        self.visited.append(index)


class FakeApplication:
    def __init__(self, *, slide_shows: int = 0) -> None:
        self.SlideShowWindows = FakeCollection([object()] * slide_shows)


class FakePresentation:
    def __init__(
        self,
        *,
        full_name: str = FULL_NAME,
        slides: list[FakeSlide] | None = None,
        windows: int = 1,
        saved: bool = True,
    ) -> None:
        self.FullName = full_name
        self.Slides = FakeCollection(
            list(slides if slides is not None else [FakeSlide(), FakeSlide(), FakeSlide()])
        )
        self.Windows = FakeCollection([object()] * windows)
        self.Saved = saved


class FakeDocumentWindow:
    def __init__(
        self,
        presentation: FakePresentation,
        application: FakeApplication,
        *,
        view_type: int = PP_VIEW_NORMAL,
    ) -> None:
        self.Presentation = presentation
        self.Application = application
        self.View = FakeView()
        self.ViewType = view_type
        self.activated = False

    def Activate(self) -> None:
        self.activated = True


def _window(
    *,
    application: FakeApplication | None = None,
    view_type: int = PP_VIEW_NORMAL,
    **kwargs,
):
    presentation = FakePresentation(**kwargs)
    app = application or FakeApplication()
    return FakeDocumentWindow(presentation, app, view_type=view_type), presentation


def _navigate(window: FakeDocumentWindow, *, slide_index: int = 2, foreground=lambda _h: True):
    return navigate_powerpoint(
        window,
        slide_index=slide_index,
        expected_path_digest=DIGEST,
        salt=SALT,
        window_handle=201,
        set_foreground=foreground,
    )


def test_focus_activates_the_window_and_moves_the_normal_view() -> None:
    window, _presentation = _window()
    assert _navigate(window) == FocusOutcome.FOCUSED.value
    assert window.activated
    assert window.View.visited == [2]


def test_wrong_presentation_is_refused_even_with_the_right_window() -> None:
    window, _presentation = _window(full_name=r"C:\work\other.pptx")
    assert _navigate(window) == FocusOutcome.BOUND_DOCUMENT_NOT_FOUND.value
    assert window.View.visited == []


@pytest.mark.parametrize("slide_index", [0, 4, 99])
def test_a_removed_or_out_of_range_slide_refuses(slide_index: int) -> None:
    window, _presentation = _window()
    assert _navigate(window, slide_index=slide_index) == (
        FocusOutcome.TARGET_SLIDE_MISSING.value
    )
    assert window.View.visited == []


def test_hidden_slide_refuses_because_normal_view_behaviour_is_unproved() -> None:
    window, _presentation = _window(slides=[FakeSlide(), FakeSlide(hidden=True)])
    assert _navigate(window) == FocusOutcome.TARGET_SLIDE_MISSING.value
    assert window.View.visited == []


def test_multiple_presentation_windows_refuse() -> None:
    window, _presentation = _window(windows=2)
    assert _navigate(window) == FocusOutcome.TARGET_WINDOW_MISSING.value
    assert window.View.visited == []


def test_zero_presentation_windows_refuse() -> None:
    window, _presentation = _window(windows=0)
    assert _navigate(window) == FocusOutcome.TARGET_WINDOW_MISSING.value


def test_a_running_slide_show_refuses() -> None:
    window, _presentation = _window(application=FakeApplication(slide_shows=1))
    assert _navigate(window) == FocusOutcome.TARGET_WINDOW_MISSING.value
    assert window.View.visited == []


@pytest.mark.parametrize("view_type", [1, 2, 3, 12])
def test_a_non_normal_editing_view_refuses(view_type: int) -> None:
    window, _presentation = _window(view_type=view_type)
    assert _navigate(window) == FocusOutcome.TARGET_WINDOW_MISSING.value
    assert window.View.visited == []


def test_foreground_denial_is_partial_success() -> None:
    window, _presentation = _window()
    assert _navigate(window, foreground=lambda _h: False) == (
        FocusOutcome.FOCUSED_WITHOUT_FOREGROUND.value
    )
    assert window.View.visited == [2]


def test_a_presentation_that_becomes_dirty_is_a_side_effect() -> None:
    window, presentation = _window()

    class DirtyingView(FakeView):
        def GotoSlide(self, index: int) -> None:
            super().GotoSlide(index)
            presentation.Saved = False

    window.View = DirtyingView()
    assert _navigate(window) == FocusOutcome.SIDE_EFFECT_DETECTED.value


def test_dirty_before_navigation_is_not_reported_as_a_side_effect() -> None:
    window, _presentation = _window(saved=False)
    assert _navigate(window) == FocusOutcome.FOCUSED.value


def test_baseline_and_current_slide_indices_are_independent() -> None:
    window, _presentation = _window()
    assert _navigate(window, slide_index=1) == FocusOutcome.FOCUSED.value
    assert _navigate(window, slide_index=3) == FocusOutcome.FOCUSED.value
    assert window.View.visited == [1, 3]


def test_focus_document_refuses_off_windows() -> None:
    document = OpenDocument(
        application=FocusApplication.POWERPOINT,
        process_id=1,
        process_created=1.0,
        windows_session_id=1,
        full_name=FULL_NAME,
        window_count=1,
        visible_window_count=1,
        visible_window_handles=(201,),
        saved=True,
        autosave=False,
    )
    assert focus_document(document, {}) == FocusOutcome.UNSUPPORTED_PLATFORM.value


def test_focus_document_uses_native_object_handle_and_visible_foreground_handle(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from qc_tool.focus import ppt_focus, win32_office

    window, _presentation = _window()
    acquired: list[int] = []
    foreground: list[int] = []
    monkeypatch.setattr(ppt_focus.sys, "platform", "win32")
    monkeypatch.setattr(
        win32_office,
        "object_from_window",
        lambda handle: acquired.append(handle) or window,
    )
    monkeypatch.setattr(
        win32_office,
        "set_foreground_window",
        lambda handle: foreground.append(handle) or True,
    )
    document = OpenDocument(
        application=FocusApplication.POWERPOINT,
        process_id=1,
        process_created=1.0,
        windows_session_id=1,
        full_name=FULL_NAME,
        window_count=1,
        visible_window_count=1,
        visible_window_handles=(201,),
        saved=True,
        autosave=False,
        object_model_window_handle=808,
    )

    assert focus_document(
        document,
        {
            "slide_index": 2,
            "path_salt": SALT.hex(),
            "expected_path_digest": DIGEST,
        },
    ) == FocusOutcome.FOCUSED.value
    assert acquired == [808]
    assert foreground == [201]
