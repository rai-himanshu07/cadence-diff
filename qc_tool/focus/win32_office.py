"""Windows-only enumeration behind :mod:`qc_tool.focus.discovery`.

Every ctypes and pywin32 call lives here so the rest of the focus package stays
importable and type-checkable on any platform. Each public entry point starts
with an explicit ``sys.platform`` guard, which is also what narrows
``ctypes.WinDLL`` and ``ctypes.WINFUNCTYPE`` for Pyright on Linux.

Nothing here activates, opens, saves, recalculates, or closes a document.
"""

from __future__ import annotations

import contextlib
import ctypes
import logging
import sys
from collections.abc import Callable
from ctypes import wintypes
from dataclasses import replace
from typing import Any

from qc_tool.focus.discovery import (
    DiscoveryReason,
    DiscoveryResult,
    FocusApplication,
    FrameObservation,
    OpenDocument,
    assess_discovery,
    read_excel_document,
    read_powerpoint_document,
    resolve_frame_visibility,
)

logger = logging.getLogger(__name__)

_OBJID_NATIVEOM = 0xFFFFFFF0
_MAX_CHILD_WINDOWS = 4096
_MAX_OBJECT_MODEL_ATTEMPTS = 64

_EXCEL_FRAME_CLASS = "XLMAIN"
_EXCEL_DOCUMENT_CLASS = "EXCEL7"
_POWERPOINT_FRAME_CLASS = "PPTFrameClass"
_POWERPOINT_DOCUMENT_CLASSES = ("mdiClass", "MDIClient", "paneClassDC")
_FRAME_CLASSES = {
    _EXCEL_FRAME_CLASS: FocusApplication.EXCEL,
    _POWERPOINT_FRAME_CLASS: FocusApplication.POWERPOINT,
}


class _GUID(ctypes.Structure):
    _fields_ = (
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    )


def _iid_idispatch() -> _GUID:
    return _GUID(
        0x00020400,
        0x0000,
        0x0000,
        (ctypes.c_ubyte * 8)(0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46),
    )


def _load_user32() -> tuple[Any, Any] | None:
    """``(user32, EnumProc)`` with explicit signatures, or ``None`` off Windows.

    Unset argtypes truncate a 64-bit ``HWND`` to a C ``int`` and silently lose
    windows, which would make discovery quietly incomplete.
    """
    if sys.platform != "win32":
        return None
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows.argtypes = (enum_proc, wintypes.LPARAM)
    user32.EnumWindows.restype = wintypes.BOOL
    user32.EnumChildWindows.argtypes = (wintypes.HWND, enum_proc, wintypes.LPARAM)
    user32.EnumChildWindows.restype = wintypes.BOOL
    user32.GetClassNameW.argtypes = (wintypes.HWND, wintypes.LPWSTR, ctypes.c_int)
    user32.GetClassNameW.restype = ctypes.c_int
    user32.IsWindowVisible.argtypes = (wintypes.HWND,)
    user32.IsWindowVisible.restype = wintypes.BOOL
    user32.IsWindow.argtypes = (wintypes.HWND,)
    user32.IsWindow.restype = wintypes.BOOL
    user32.SetForegroundWindow.argtypes = (wintypes.HWND,)
    user32.SetForegroundWindow.restype = wintypes.BOOL
    user32.GetWindowThreadProcessId.argtypes = (
        wintypes.HWND,
        ctypes.POINTER(wintypes.DWORD),
    )
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    return user32, enum_proc


def _class_name(user32: Any, hwnd: int) -> str:
    buffer = ctypes.create_unicode_buffer(256)
    length = user32.GetClassNameW(hwnd, buffer, len(buffer))
    return buffer.value if length > 0 else ""


def window_is_visible(hwnd: int) -> bool | None:
    """Live visibility for one window handle, or ``None`` when unprovable."""
    loaded = _load_user32()
    if loaded is None or not hwnd:
        return None
    user32, _enum_proc = loaded
    if not user32.IsWindow(wintypes.HWND(hwnd)):
        return None
    return bool(user32.IsWindowVisible(wintypes.HWND(hwnd)))


def set_foreground_window(hwnd: int) -> bool:
    """Best-effort foreground activation. Denial is never a reason to retarget."""
    loaded = _load_user32()
    if loaded is None or not hwnd:
        return False
    user32, _enum_proc = loaded
    try:
        return bool(user32.SetForegroundWindow(wintypes.HWND(hwnd)))
    except Exception:
        return False


def _enumerate_top_level_windows() -> list[tuple[int, str]]:
    loaded = _load_user32()
    if loaded is None:
        return []
    user32, enum_proc = loaded
    found: list[tuple[int, str]] = []

    def callback(hwnd: int, _param: int) -> bool:
        if hwnd:
            found.append((int(hwnd), _class_name(user32, hwnd)))
        return True

    user32.EnumWindows(enum_proc(callback), 0)
    return found


def _enumerate_child_windows(hwnd: int) -> list[tuple[int, str]]:
    loaded = _load_user32()
    if loaded is None:
        return []
    user32, enum_proc = loaded
    found: list[tuple[int, str]] = []

    def callback(child: int, _param: int) -> bool:
        if len(found) >= _MAX_CHILD_WINDOWS:
            return False
        if child:
            found.append((int(child), _class_name(user32, child)))
        return True

    user32.EnumChildWindows(wintypes.HWND(hwnd), enum_proc(callback), 0)
    return found


def _window_state(hwnd: int) -> tuple[bool, int]:
    """``(visible, process id)`` for one window."""
    loaded = _load_user32()
    if loaded is None:
        return False, 0
    user32, _enum_proc = loaded
    process_id = wintypes.DWORD(0)
    user32.GetWindowThreadProcessId(wintypes.HWND(hwnd), ctypes.byref(process_id))
    return bool(user32.IsWindowVisible(wintypes.HWND(hwnd))), int(process_id.value)


def _current_identity() -> tuple[object | None, int | None]:
    if sys.platform != "win32":
        return None, None
    import win32api  # pyright: ignore[reportMissingModuleSource]
    import win32con  # pyright: ignore[reportMissingModuleSource]
    import win32process  # pyright: ignore[reportMissingModuleSource]
    import win32security  # pyright: ignore[reportMissingModuleSource]
    import win32ts  # pyright: ignore[reportMissingModuleSource]

    token = win32security.OpenProcessToken(
        win32api.GetCurrentProcess(), win32con.TOKEN_QUERY
    )
    try:
        sid = win32security.GetTokenInformation(token, win32security.TokenUser)[0]
    finally:
        win32api.CloseHandle(token)
    session = win32ts.ProcessIdToSessionId(win32process.GetCurrentProcessId())
    return sid, int(session)


def _process_facts(
    process_id: int, own_sid: object | None, own_session: int | None
) -> tuple[bool, bool, bool, int, float]:
    """``(identity known, same user, same session, session id, created)``."""
    if sys.platform != "win32":
        return False, False, False, -1, 0.0
    import win32api  # pyright: ignore[reportMissingModuleSource]
    import win32con  # pyright: ignore[reportMissingModuleSource]
    import win32process  # pyright: ignore[reportMissingModuleSource]
    import win32security  # pyright: ignore[reportMissingModuleSource]
    import win32ts  # pyright: ignore[reportMissingModuleSource]

    try:
        session = int(win32ts.ProcessIdToSessionId(process_id))
    except Exception:
        return False, False, False, -1, 0.0
    query_limited = getattr(win32con, "PROCESS_QUERY_LIMITED_INFORMATION", 0x1000)
    try:
        handle = win32api.OpenProcess(query_limited, False, process_id)
    except Exception:
        return False, False, session == own_session, session, 0.0
    try:
        created = 0.0
        with contextlib.suppress(Exception):
            created = float(
                win32process.GetProcessTimes(handle)["CreationTime"].timestamp()
            )
        token = win32security.OpenProcessToken(handle, win32con.TOKEN_QUERY)
        try:
            sid = win32security.GetTokenInformation(token, win32security.TokenUser)[0]
        finally:
            win32api.CloseHandle(token)
    except Exception:
        return False, False, session == own_session, session, 0.0
    finally:
        win32api.CloseHandle(handle)
    return True, sid == own_sid, session == own_session, session, created


def _dispatch_via_pythoncom(address: int) -> object | None:
    if sys.platform != "win32":
        return None
    import pythoncom  # pyright: ignore[reportMissingModuleSource]
    import win32com.client  # pyright: ignore[reportMissingModuleSource]

    from_address: Callable[..., Any] | None = getattr(
        pythoncom, "ObjectFromAddress", None
    )
    if from_address is None:
        return None
    try:
        return win32com.client.Dispatch(from_address(address, pythoncom.IID_IDispatch))
    except Exception:
        return None


def object_from_window(hwnd: int) -> object | None:
    """The native object model behind one Office window, or ``None``."""
    if sys.platform != "win32":
        return None
    oleacc = ctypes.WinDLL("oleacc", use_last_error=True)
    pointer = ctypes.c_void_p()
    iid = _iid_idispatch()
    try:
        result = oleacc.AccessibleObjectFromWindow(
            wintypes.HWND(hwnd),
            wintypes.DWORD(_OBJID_NATIVEOM),
            ctypes.byref(iid),
            ctypes.byref(pointer),
        )
    except Exception:
        return None
    if result != 0 or not pointer.value:
        return None
    return _dispatch_via_pythoncom(int(pointer.value))


def _document_windows(frame_hwnd: int, application: FocusApplication) -> list[int]:
    children = _enumerate_child_windows(frame_hwnd)
    if application is FocusApplication.EXCEL:
        exact = [child for child, name in children if name == _EXCEL_DOCUMENT_CLASS]
        if exact:
            return exact[:_MAX_OBJECT_MODEL_ATTEMPTS]
        preferred: list[int] = []
    else:
        preferred = [
            child for child, name in children if name in _POWERPOINT_DOCUMENT_CLASSES
        ]
    remaining = [child for child, _name in children if child not in preferred]
    ordered = [frame_hwnd, *preferred, *remaining]
    return ordered[:_MAX_OBJECT_MODEL_ATTEMPTS]


def discover(application: FocusApplication) -> DiscoveryResult:
    """Enumerate every same-user, same-session document of one application."""
    if sys.platform != "win32":
        return DiscoveryResult(
            application=application, reasons=(DiscoveryReason.UNSUPPORTED_PLATFORM,)
        )
    own_sid, own_session = _current_identity()
    frames: list[FrameObservation] = []
    documents: dict[tuple[int, str], OpenDocument] = {}
    visible_frames: dict[tuple[int, str], tuple[int, ...]] = {}
    attempted = False
    native_available = False

    for hwnd, class_name in _enumerate_top_level_windows():
        frame_application = _FRAME_CLASSES.get(class_name)
        if frame_application is not application:
            continue
        visible, process_id = _window_state(hwnd)
        identity_known, same_user, same_session, session_id, created = _process_facts(
            process_id, own_sid, own_session
        )
        reached = False
        protected = False
        for candidate in _document_windows(hwnd, application):
            attempted = True
            window_object = object_from_window(candidate)
            if window_object is None:
                continue
            native_available = True
            reader = (
                read_excel_document
                if application is FocusApplication.EXCEL
                else read_powerpoint_document
            )
            observed = reader(
                window_object,
                process_id=process_id,
                process_created=created,
                windows_session_id=session_id,
                window_visible=window_is_visible,
            )
            if observed is None:
                continue
            document, protected = observed
            document = replace(document, object_model_window_handle=candidate)
            documents[document.instance_key] = document
            reached = True
            if visible:
                key = document.instance_key
                visible_frames[key] = (*visible_frames.get(key, ()), hwnd)
            break
        frames.append(
            FrameObservation(
                application=application,
                process_id=process_id,
                visible=visible,
                identity_known=identity_known,
                same_user=same_user,
                same_session=same_session,
                protected_view=protected,
                reached_document=reached,
            )
        )
    repaired, _unresolved = resolve_frame_visibility(
        list(documents.values()), visible_frames
    )
    return assess_discovery(
        application,
        frames,
        repaired,
        native_object_model_available=native_available,
        object_model_attempted=attempted,
    )
