"""PowerPoint slide focus inside the short-lived helper.

Slide-level only: exact shape focus is a later step. The navigation core is
duck-typed over the object model so it runs against fake COM objects on any
platform. It activates one normal editing window and moves its view to one
slide index. It never opens, saves, closes, or quits a presentation.
"""

from __future__ import annotations

import contextlib
import sys
from collections.abc import Callable

from qc_tool.focus.discovery import OpenDocument, as_bool, as_count, read_attribute
from qc_tool.focus.protocol import FocusOutcome, path_digest

#: ``ppViewNormal``. Slide-show, presenter, and every other view refuse.
PP_VIEW_NORMAL = 9
#: ``msoTrue``; a slide hidden from the show is refused rather than guessed at.
MSO_TRUE = -1

Foreground = Callable[[int], bool]


def _slide(presentation: object, index: int) -> object | None:
    slides = read_attribute(presentation, "Slides")
    if not 1 <= index <= as_count(read_attribute(slides, "Count")):
        return None
    with contextlib.suppress(Exception):
        return slides.Item(index)  # type: ignore[attr-defined]
    return None


def _slide_is_hidden(slide: object) -> bool:
    transition = read_attribute(slide, "SlideShowTransition")
    if transition is None:
        return False
    return read_attribute(transition, "Hidden") == MSO_TRUE


def navigate_powerpoint(
    window_object: object,
    *,
    slide_index: int,
    expected_path_digest: str,
    salt: bytes,
    window_handle: int,
    set_foreground: Foreground,
) -> str:
    """Move one exact presentation's normal editing view to one exact slide."""
    presentation = read_attribute(window_object, "Presentation")
    application = read_attribute(window_object, "Application")
    if presentation is None:
        presentation = read_attribute(window_object, "Parent")
    if presentation is None or application is None:
        return FocusOutcome.BOUND_DOCUMENT_NOT_FOUND.value
    full_name = read_attribute(presentation, "FullName")
    if not isinstance(full_name, str) or path_digest(full_name, salt) != (
        expected_path_digest
    ):
        return FocusOutcome.BOUND_DOCUMENT_NOT_FOUND.value
    if as_count(read_attribute(read_attribute(presentation, "Windows"), "Count")) != 1:
        return FocusOutcome.TARGET_WINDOW_MISSING.value
    show_windows = read_attribute(application, "SlideShowWindows")
    if show_windows is not None and as_count(read_attribute(show_windows, "Count")) > 0:
        return FocusOutcome.TARGET_WINDOW_MISSING.value
    view = read_attribute(window_object, "View")
    if view is None or read_attribute(window_object, "ViewType") != PP_VIEW_NORMAL:
        return FocusOutcome.TARGET_WINDOW_MISSING.value
    saved_before = as_bool(read_attribute(presentation, "Saved"))
    slide = _slide(presentation, slide_index)
    if slide is None:
        return FocusOutcome.TARGET_SLIDE_MISSING.value
    if _slide_is_hidden(slide):
        # Normal-view behaviour for a hidden slide was never proved live.
        return FocusOutcome.TARGET_SLIDE_MISSING.value
    try:
        window_object.Activate()  # type: ignore[attr-defined]
        view.GotoSlide(slide_index)  # type: ignore[attr-defined]
    except Exception:
        return FocusOutcome.HELPER_FAILED.value
    if (
        saved_before is True
        and as_bool(read_attribute(presentation, "Saved")) is False
    ):
        return FocusOutcome.SIDE_EFFECT_DETECTED.value
    if not set_foreground(window_handle):
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
    slide_index = request.get("slide_index")
    salt_hex = request.get("path_salt")
    expected = request.get("expected_path_digest")
    if (
        not isinstance(slide_index, int)
        or isinstance(slide_index, bool)
        or slide_index < 1
        or not isinstance(salt_hex, str)
        or not isinstance(expected, str)
    ):
        return FocusOutcome.INVALID_REQUEST.value
    window_object = win32_office.object_from_window(handles[0])
    if window_object is None:
        return FocusOutcome.TARGET_WINDOW_MISSING.value
    return navigate_powerpoint(
        window_object,
        slide_index=slide_index,
        expected_path_digest=expected,
        salt=bytes.fromhex(salt_hex),
        window_handle=handles[0],
        set_foreground=win32_office.set_foreground_window,
    )
