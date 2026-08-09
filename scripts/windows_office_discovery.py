"""Aggregate-only Windows spike proving complete read-only Office document discovery.

Step 1 of the click-to-focus plan. This probe never activates, opens, saves, or
closes an analyst-owned document; it only enumerates windows and reads identity
metadata. It emits fixed codes and counts. Paths, filenames, window titles, user
SIDs, sheet names, and document content are never written to the result file.
"""

from __future__ import annotations

import argparse
import contextlib
import ctypes
import hashlib
import importlib
import json
import os
import secrets
import sys
import time
from collections.abc import Callable, Sequence
from ctypes import wintypes
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any

_SCHEMA_VERSION = 1
_UNDECLARED = -1

_EXCEL = "excel"
_POWERPOINT = "powerpoint"
_APPLICATIONS = (_EXCEL, _POWERPOINT)

_EXCEL_FRAME_CLASS = "XLMAIN"
_EXCEL_DOCUMENT_CLASS = "EXCEL7"
_POWERPOINT_FRAME_CLASS = "PPTFrameClass"
_FRAME_CLASSES = {
    _EXCEL_FRAME_CLASS: _EXCEL,
    _POWERPOINT_FRAME_CLASS: _POWERPOINT,
}

_OBJID_NATIVEOM = 0xFFFFFFF0
_MAX_CHILD_WINDOWS = 4096
_MAX_ROT_ENTRIES = 4096
_MAX_OBJECT_MODEL_ATTEMPTS = 64
_POWERPOINT_DOCUMENT_CLASSES = ("mdiClass", "MDIClient", "paneClassDC")

_SCENARIOS = (
    "baseline",
    "two-excel-processes",
    "multi-window",
    "minimized",
    "protected-view",
    "hidden-worker",
    "onedrive",
    "unc",
    "sharepoint-url",
    "elevated",
)

_EXTENSION_KINDS = {
    ".xls": "xls",
    ".xlsb": "xlsb",
    ".xlsm": "xlsm",
    ".xlsx": "xlsx",
    ".ppt": "ppt",
    ".pptm": "pptm",
    ".pptx": "pptx",
}
_EXCEL_EXTENSIONS = frozenset({"xls", "xlsb", "xlsm", "xlsx"})
_POWERPOINT_EXTENSIONS = frozenset({"ppt", "pptm", "pptx"})

_MECHANISM_NONE = "none"
_MECHANISM_PYTHONCOM = "native_om_pythoncom"
_MECHANISM_COMTYPES = "native_om_comtypes"
_MECHANISM_ROT = "running_object_table"

_VERDICT_COMPLETE = "complete_and_exact"
_VERDICT_REFUSED = "refused_incomplete"
_VERDICT_UNDERCOUNT = "silent_undercount"
_VERDICT_OVERCOUNT = "silent_overcount"
_VERDICT_UNDECLARED = "expectations_not_declared"
_VERDICT_NOT_APPLICABLE = "not_applicable"
_FATAL_VERDICTS = frozenset({_VERDICT_UNDERCOUNT, _VERDICT_OVERCOUNT})
_PASSING_VERDICTS = frozenset(
    {_VERDICT_COMPLETE, _VERDICT_REFUSED, _VERDICT_NOT_APPLICABLE}
)

# Opaque per-run salt so document keys cannot be correlated across runs.
_KEY_SALT = secrets.token_bytes(16)


@dataclass(frozen=True, slots=True)
class WindowFact:
    """One observed top-level or document window belonging to an Office frame."""

    application: str
    process_id: int
    class_name: str
    visible: bool
    minimized: bool
    same_user: bool
    same_session: bool
    identity_known: bool
    protected_view: bool
    document_key: str | None
    mechanism: str


@dataclass(frozen=True, slots=True)
class DocumentFact:
    """One deduplicated document identity reached through the object model."""

    application: str
    document_key: str
    process_id: int
    window_count: int
    visible_window_count: int
    saved: bool | None
    autosave: bool | None
    path_kind: str
    extension_kind: str
    hidden_instance: bool
    is_addin: bool
    visibility_unproved: bool
    source: str

    @property
    def analyst_candidate(self) -> bool:
        """Add-ins and windowless workbooks are never analyst-owned documents."""
        return (
            not self.hidden_instance
            and not self.is_addin
            and self.visible_window_count > 0
        )


@dataclass(frozen=True, slots=True)
class Expectation:
    """Operator-declared scenario staged manually inside the guest."""

    processes: int = _UNDECLARED
    documents: int = _UNDECLARED
    windows: int = _UNDECLARED

    @property
    def declared(self) -> bool:
        return self.documents != _UNDECLARED


def _document_key(*parts: str) -> str:
    digest = hashlib.blake2s(_KEY_SALT, digest_size=12)
    for part in parts:
        digest.update(b"\x00")
        digest.update(part.casefold().encode("utf-8", "replace"))
    return digest.hexdigest()


def classify_path_kind(full_name: str) -> str:
    """Classify a document location without retaining any part of it."""
    text = full_name.strip()
    if not text:
        return "none"
    if text[:8].casefold().startswith(("http://", "https://")):
        return "url"
    if text.startswith("\\\\"):
        return "unc"
    if len(text) >= 3 and text[1] == ":" and text[2] in "\\/":
        return "local"
    return "unrecognized"


def classify_extension(full_name: str) -> str:
    """Classify a document format without retaining the filename."""
    text = full_name.strip()
    if not text:
        return "none"
    base = text.rsplit("\\", 1)[-1].rsplit("/", 1)[-1]
    if "." not in base:
        return "none"
    return _EXTENSION_KINDS.get(base[base.rfind(".") :].casefold(), "other")


def application_for_extension(extension_kind: str) -> str | None:
    if extension_kind in _EXCEL_EXTENSIONS:
        return _EXCEL
    if extension_kind in _POWERPOINT_EXTENSIONS:
        return _POWERPOINT
    return None


def _counter(values: Sequence[str]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for value in values:
        counts[value] = counts.get(value, 0) + 1
    return dict(sorted(counts.items()))


def incomplete_reasons(
    windows: Sequence[WindowFact],
    documents: Sequence[DocumentFact],
    *,
    rot_only_documents: int,
    native_om_available: bool,
    object_model_attempted: bool = True,
    expectation: Expectation | None = None,
) -> list[str]:
    """Return the fixed codes that force `enumeration_incomplete` for one application."""
    reasons: set[str] = set()
    if object_model_attempted and not native_om_available:
        reasons.add("native_object_model_unavailable")
    if not windows and expectation is not None and expectation.documents > 0:
        reasons.add("office_windows_not_found")
    if rot_only_documents > 0:
        reasons.add("running_object_table_document_not_enumerated")
    for window in windows:
        if not window.identity_known:
            reasons.add("process_identity_unavailable")
            continue
        if not window.same_user:
            reasons.add("cross_user_office_window")
        if not window.same_session:
            reasons.add("cross_session_office_window")
        if window.protected_view:
            reasons.add("protected_view_window")
        if window.visible and window.document_key is None and not window.protected_view:
            reasons.add("document_window_object_model_unreachable")
    for document in documents:
        # An unproved window state must refuse, never silently drop a document.
        if document.visibility_unproved and not document.is_addin:
            reasons.add("window_visibility_unproved")
        if not document.analyst_candidate:
            continue
        if document.window_count == 0:
            reasons.add("document_without_window")
        if document.saved is None:
            reasons.add("saved_state_unproved")
        if document.autosave is None:
            reasons.add("autosave_state_unproved")
    return sorted(reasons)


def _verdict(
    coverage: str,
    *,
    expectation: Expectation,
    observed_anything: bool,
    analyst_documents: int,
    analyst_processes: int,
    analyst_windows: int,
) -> str:
    if not expectation.declared:
        return _VERDICT_UNDECLARED if observed_anything else _VERDICT_NOT_APPLICABLE
    if coverage != "complete":
        return _VERDICT_REFUSED
    pairs = (
        (expectation.documents, analyst_documents),
        (expectation.processes, analyst_processes),
        (expectation.windows, analyst_windows),
    )
    compared = [
        (expected, actual) for expected, actual in pairs if expected != _UNDECLARED
    ]
    # Missing evidence is the dangerous direction, so it wins over a surplus.
    if any(actual < expected for expected, actual in compared):
        return _VERDICT_UNDERCOUNT
    if any(actual > expected for expected, actual in compared):
        return _VERDICT_OVERCOUNT
    return _VERDICT_COMPLETE


def summarize_application(
    application: str,
    windows: Sequence[WindowFact],
    documents: Sequence[DocumentFact],
    *,
    expectation: Expectation,
    rot_only_documents: int,
    native_om_available: bool,
    object_model_attempted: bool = True,
) -> dict[str, object]:
    """Reduce raw observations to the aggregate record written to the result file."""
    reasons = incomplete_reasons(
        windows,
        documents,
        rot_only_documents=rot_only_documents,
        native_om_available=native_om_available,
        object_model_attempted=object_model_attempted,
        expectation=expectation,
    )
    coverage = "complete" if not reasons else "enumeration_incomplete"
    analyst = [document for document in documents if document.analyst_candidate]
    analyst_processes = {document.process_id for document in analyst}
    analyst_windows = sum(document.visible_window_count for document in analyst)
    paths = _counter([document.document_key for document in analyst])
    verdict = _verdict(
        coverage,
        expectation=expectation,
        observed_anything=bool(windows or documents),
        analyst_documents=len(analyst),
        analyst_processes=len(analyst_processes),
        analyst_windows=analyst_windows,
    )
    matches = {
        "documents": expectation.documents in {_UNDECLARED, len(analyst)},
        "processes": expectation.processes in {_UNDECLARED, len(analyst_processes)},
        "windows": expectation.windows in {_UNDECLARED, analyst_windows},
    }
    return {
        "addin_documents": sum(1 for document in documents if document.is_addin),
        "application": application,
        "analyst_documents": len(analyst),
        "analyst_processes": len(analyst_processes),
        "analyst_visible_windows": analyst_windows,
        "autosave_off_documents": sum(
            1 for document in analyst if document.autosave is False
        ),
        "autosave_on_documents": sum(
            1 for document in analyst if document.autosave is True
        ),
        "autosave_unknown_documents": sum(
            1 for document in analyst if document.autosave is None
        ),
        "coverage": coverage,
        "dirty_documents": sum(1 for document in analyst if document.saved is False),
        "documents_discovered": len(documents),
        "duplicate_path_analyst_documents": sum(
            count for count in paths.values() if count > 1
        ),
        "expected_documents": expectation.documents,
        "expected_matches": matches,
        "expected_processes": expectation.processes,
        "expected_visible_windows": expectation.windows,
        "extension_kinds": _counter([document.extension_kind for document in analyst]),
        "fatal": verdict in _FATAL_VERDICTS,
        "hidden_instance_documents": sum(
            1 for document in documents if document.hidden_instance
        ),
        "incomplete_reasons": reasons,
        "mechanisms": _counter([window.mechanism for window in windows]),
        "minimized_windows": sum(1 for window in windows if window.minimized),
        "native_object_model_available": native_om_available,
        "passed": verdict in _PASSING_VERDICTS,
        "path_kinds": _counter([document.path_kind for document in analyst]),
        "protected_view_windows": sum(1 for window in windows if window.protected_view),
        "running_object_table_only_documents": rot_only_documents,
        "saved_documents": sum(1 for document in analyst if document.saved is True),
        "verdict": verdict,
        "window_processes": len({window.process_id for window in windows}),
        "windowless_documents": sum(
            1 for document in documents if document.window_count == 0
        ),
        "windows_observed": len(windows),
        "windows_visible": sum(1 for window in windows if window.visible),
        "windows_without_document": sum(
            1 for window in windows if window.document_key is None
        ),
    }


def evaluate(applications: dict[str, dict[str, object]]) -> tuple[str, bool]:
    """Fold per-application records into one overall outcome."""
    if any(record.get("fatal") is True for record in applications.values()):
        return "fatal_silent_miss", False
    passed = bool(applications) and all(
        record.get("passed") is True for record in applications.values()
    )
    if not passed:
        return "failed", False
    if all(record.get("coverage") == "complete" for record in applications.values()):
        return "complete", True
    return "enumeration_incomplete", True


def _write_json(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=True, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    os.replace(temporary, path)


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
    """Return `(user32, EnumProc)` with explicit signatures, or None off Windows.

    Unset argtypes truncate 64-bit HWNDs to C int, which silently loses windows.
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
    user32.IsIconic.argtypes = (wintypes.HWND,)
    user32.IsIconic.restype = wintypes.BOOL
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


def _enumerate_all_top_level_windows() -> list[tuple[int, str]]:
    """Return `(hwnd, class_name)` for every top-level window on the desktop."""
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


def _enumerate_top_level_windows() -> list[tuple[int, str]]:
    """Return `(hwnd, class_name)` for every top-level Office frame window."""
    return [
        window for window in _enumerate_all_top_level_windows()
        if window[1] in _FRAME_CLASSES
    ]


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

    user32.EnumChildWindows(hwnd, enum_proc(callback), 0)
    return found


def _window_state(hwnd: int) -> tuple[bool, bool, int]:
    """Return `(visible, minimized, process_id)` for one window."""
    loaded = _load_user32()
    if loaded is None:
        return False, False, 0
    user32, _enum_proc = loaded
    process_id = wintypes.DWORD(0)
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
    return (
        bool(user32.IsWindowVisible(hwnd)),
        bool(user32.IsIconic(hwnd)),
        int(process_id.value),
    )


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


def _process_identity(
    process_id: int, own_sid: object | None, own_session: int | None
) -> tuple[bool, bool, bool]:
    """Return `(identity_known, same_user, same_session)` for one process."""
    if sys.platform != "win32":
        return False, False, False
    import win32api  # pyright: ignore[reportMissingModuleSource]
    import win32con  # pyright: ignore[reportMissingModuleSource]
    import win32security  # pyright: ignore[reportMissingModuleSource]
    import win32ts  # pyright: ignore[reportMissingModuleSource]

    try:
        session = int(win32ts.ProcessIdToSessionId(process_id))
    except Exception:
        return False, False, False
    query_limited = getattr(win32con, "PROCESS_QUERY_LIMITED_INFORMATION", 0x1000)
    try:
        handle = win32api.OpenProcess(query_limited, False, process_id)
    except Exception:
        return False, False, session == own_session
    try:
        token = win32security.OpenProcessToken(handle, win32con.TOKEN_QUERY)
        try:
            sid = win32security.GetTokenInformation(token, win32security.TokenUser)[0]
        finally:
            win32api.CloseHandle(token)
    except Exception:
        return False, False, session == own_session
    finally:
        win32api.CloseHandle(handle)
    return True, sid == own_sid, session == own_session


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
        return win32com.client.Dispatch(
            from_address(address, pythoncom.IID_IDispatch)
        )
    except Exception:
        return None


def _object_from_window_comtypes(hwnd: int) -> object | None:
    try:
        automation = importlib.import_module("comtypes.automation")
        client = importlib.import_module("comtypes.client")
        comtypes = importlib.import_module("comtypes")
    except ImportError:
        return None
    pointer = ctypes.POINTER(automation.IDispatch)()
    try:
        comtypes.windll.oleacc.AccessibleObjectFromWindow(
            hwnd,
            _OBJID_NATIVEOM,
            ctypes.byref(automation.IDispatch._iid_),
            ctypes.byref(pointer),
        )
    except Exception:
        return None
    if not pointer:
        return None
    try:
        return client.GetBestInterface(pointer)
    except Exception:
        return None


def _object_from_window(hwnd: int) -> tuple[object | None, str]:
    """Return `(native object model, mechanism)` for one document window."""
    if sys.platform != "win32":
        return None, _MECHANISM_NONE
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
        result = -1
    if result == 0 and pointer.value:
        dispatched = _dispatch_via_pythoncom(int(pointer.value))
        if dispatched is not None:
            return dispatched, _MECHANISM_PYTHONCOM
    fallback = _object_from_window_comtypes(hwnd)
    if fallback is not None:
        return fallback, _MECHANISM_COMTYPES
    return None, _MECHANISM_NONE


def _as_bool(value: object) -> bool | None:
    if value is None:
        return None
    with contextlib.suppress(Exception):
        return bool(value)
    return None


def _read_attribute(source: object, name: str) -> object | None:
    with contextlib.suppress(Exception):
        return getattr(source, name)
    return None


def _window_is_visible(item: object) -> bool | None:
    """Excel windows expose `Visible`; PowerPoint document windows only expose a handle."""
    flag = _as_bool(_read_attribute(item, "Visible"))
    if flag is not None:
        return flag
    for name in ("HWND", "Hwnd"):
        handle = _read_attribute(item, name)
        if isinstance(handle, int) and handle:
            visible, _minimized, _process_id = _window_state(handle)
            return visible
    return None


def _window_counts(collection: object) -> tuple[int, int, bool]:
    """Return `(window count, visible count, visibility unproved)` for a document."""
    total = 0
    with contextlib.suppress(Exception):
        total = int(getattr(collection, "Count", 0))
    visible = 0
    unproved = False
    for index in range(1, total + 1):
        item: object | None = None
        with contextlib.suppress(Exception):
            item = collection.Item(index)  # pyright: ignore[reportAttributeAccessIssue]
        if item is None:
            unproved = True
            continue
        state = _window_is_visible(item)
        if state is None:
            unproved = True
        elif state:
            visible += 1
    return total, visible, unproved


def _excel_document(window_object: object) -> tuple[DocumentFact, bool] | None:
    workbook = _read_attribute(window_object, "Parent")
    application = _read_attribute(window_object, "Application")
    if workbook is None or application is None:
        return None
    full_name = _read_attribute(workbook, "FullName")
    if not isinstance(full_name, str):
        return None
    hidden_instance = _as_bool(_read_attribute(application, "Visible")) is not True
    protected_view_count = _read_attribute(application, "ProtectedViewWindows")
    protected = False
    with contextlib.suppress(Exception):
        protected = int(getattr(protected_view_count, "Count", 0)) > 0
    window_count, visible_windows, visibility_unproved = _window_counts(
        _read_attribute(workbook, "Windows")
    )
    return (
        DocumentFact(
            application=_EXCEL,
            document_key=_document_key(_EXCEL, full_name),
            process_id=0,
            window_count=window_count,
            visible_window_count=0 if hidden_instance else visible_windows,
            saved=_as_bool(_read_attribute(workbook, "Saved")),
            autosave=_as_bool(_read_attribute(workbook, "AutoSaveOn")),
            path_kind=classify_path_kind(full_name),
            extension_kind=classify_extension(full_name),
            hidden_instance=hidden_instance,
            is_addin=_as_bool(_read_attribute(workbook, "IsAddin")) is True,
            visibility_unproved=visibility_unproved,
            source="window",
        ),
        protected,
    )


def _powerpoint_document(window_object: object) -> tuple[DocumentFact, bool] | None:
    presentation = _read_attribute(window_object, "Presentation")
    application = _read_attribute(window_object, "Application")
    if presentation is None:
        presentation = _read_attribute(window_object, "Parent")
    if presentation is None or application is None:
        return None
    full_name = _read_attribute(presentation, "FullName")
    if not isinstance(full_name, str):
        return None
    hidden_instance = _as_bool(_read_attribute(application, "Visible")) is not True
    window_count, visible_windows, visibility_unproved = _window_counts(
        _read_attribute(presentation, "Windows")
    )
    return (
        DocumentFact(
            application=_POWERPOINT,
            document_key=_document_key(_POWERPOINT, full_name),
            process_id=0,
            window_count=window_count,
            visible_window_count=0 if hidden_instance else visible_windows,
            saved=_as_bool(_read_attribute(presentation, "Saved")),
            autosave=_as_bool(_read_attribute(presentation, "AutoSaveOn")),
            path_kind=classify_path_kind(full_name),
            extension_kind=classify_extension(full_name),
            hidden_instance=hidden_instance,
            is_addin=False,
            visibility_unproved=visibility_unproved,
            source="window",
        ),
        False,
    )


def _document_windows(frame_hwnd: int, application: str) -> list[tuple[int, str]]:
    """Return candidate document windows for one Office frame, most likely first."""
    frame_class = (
        _EXCEL_FRAME_CLASS if application == _EXCEL else _POWERPOINT_FRAME_CLASS
    )
    children = _enumerate_child_windows(frame_hwnd)
    if application == _EXCEL:
        candidates = [
            child for child in children if child[1] == _EXCEL_DOCUMENT_CLASS
        ]
        if candidates:
            return candidates
        preferred: list[tuple[int, str]] = []
    else:
        preferred = [
            child for child in children if child[1] in _POWERPOINT_DOCUMENT_CLASSES
        ]
    remaining = [child for child in children if child not in preferred]
    ordered = [(frame_hwnd, frame_class), *preferred, *remaining]
    return ordered[:_MAX_OBJECT_MODEL_ATTEMPTS]


def _running_object_table_documents() -> tuple[dict[str, tuple[str, str, str]], int, str]:
    """Return `(documents, total entries, error code)` from the Running Object Table."""
    if sys.platform != "win32":
        return {}, 0, "unsupported_platform"
    import pythoncom  # pyright: ignore[reportMissingModuleSource]

    documents: dict[str, tuple[str, str, str]] = {}
    monikers: list[Any] = []
    try:
        table = pythoncom.GetRunningObjectTable()
        # The shipped stub declares no argument; older builds require a reserved 0.
        create_bind_ctx: Any = pythoncom.CreateBindCtx
        try:
            context = create_bind_ctx()
        except TypeError:
            context = create_bind_ctx(0)
        # The stub declares one moniker; pywin32 actually returns a tuple.
        enumerator: Any = table.EnumRunning()
        while len(monikers) < _MAX_ROT_ENTRIES:
            batch: Any = enumerator.Next(32)
            if not batch:
                break
            monikers.extend(batch if isinstance(batch, tuple) else (batch,))
    except Exception as exc:
        return {}, 0, type(exc).__name__
    for moniker in monikers:
        try:
            display_name = moniker.GetDisplayName(context, None)
        except Exception:
            continue
        if not isinstance(display_name, str):
            continue
        extension_kind = classify_extension(display_name)
        application = application_for_extension(extension_kind)
        if application is None:
            continue
        documents[_document_key(application, display_name)] = (
            application,
            classify_path_kind(display_name),
            extension_kind,
        )
    return documents, len(monikers), ""


@dataclass(slots=True)
class Diagnostics:
    """Why enumeration produced what it produced. Fixed codes and counts only."""

    top_level_windows_total: int = 0
    office_frame_windows: int = 0
    frame_class_counts: dict[str, int] | None = None
    top_level_class_counts: dict[str, int] | None = None
    window_visibility_from_property: int = 0
    window_visibility_from_frame: int = 0
    window_visibility_unresolved: int = 0
    object_model_attempts: int = 0
    object_model_successes: int = 0
    running_object_table_entries: int = 0
    running_object_table_error: str = ""
    pythoncom_object_from_address: bool = False
    comtypes_importable: bool = False

    def as_payload(self) -> dict[str, object]:
        payload: dict[str, object] = {
            "comtypes_importable": self.comtypes_importable,
            "frame_class_counts": dict(sorted((self.frame_class_counts or {}).items())),
            "object_model_attempts": self.object_model_attempts,
            "object_model_successes": self.object_model_successes,
            "office_frame_windows": self.office_frame_windows,
            "pythoncom_object_from_address": self.pythoncom_object_from_address,
            "running_object_table_entries": self.running_object_table_entries,
            "running_object_table_error": self.running_object_table_error,
            "top_level_windows_total": self.top_level_windows_total,
            "window_visibility_from_frame": self.window_visibility_from_frame,
            "window_visibility_from_property": self.window_visibility_from_property,
            "window_visibility_unresolved": self.window_visibility_unresolved,
        }
        if self.top_level_class_counts is not None:
            payload["top_level_class_counts"] = dict(
                sorted(self.top_level_class_counts.items())
            )
        return payload


def _pythoncom_object_from_address_available() -> bool:
    if sys.platform != "win32":
        return False
    import pythoncom  # pyright: ignore[reportMissingModuleSource]

    return getattr(pythoncom, "ObjectFromAddress", None) is not None


def _comtypes_importable() -> bool:
    try:
        importlib.import_module("comtypes.client")
    except ImportError:
        return False
    return True


def resolve_frame_visibility(
    documents: dict[tuple[int, str], DocumentFact],
    visible_frames: dict[tuple[int, str], int],
) -> tuple[dict[tuple[int, str], DocumentFact], int, int]:
    """Fall back to directly observed visible Office frames when COM cannot say.

    PowerPoint document windows expose no readable visibility, so the enumerated
    top-level frame that reached the document is the evidence of record.
    """
    resolved: dict[tuple[int, str], DocumentFact] = {}
    from_frame = 0
    unresolved = 0
    for key, document in documents.items():
        if not document.visibility_unproved or document.visible_window_count > 0:
            resolved[key] = document
            continue
        observed = visible_frames.get(key, 0)
        if observed > 0:
            from_frame += 1
            resolved[key] = replace(
                document,
                visible_window_count=observed,
                window_count=max(document.window_count, observed),
                visibility_unproved=False,
            )
            continue
        unresolved += 1
        resolved[key] = document
    return resolved, from_frame, unresolved


def _collect(
    diagnostics: Diagnostics,
    *,
    dump_window_classes: bool = False,
) -> tuple[dict[str, list[WindowFact]], dict[str, list[DocumentFact]], bool]:
    own_sid, own_session = _current_identity()
    windows: dict[str, list[WindowFact]] = {name: [] for name in _APPLICATIONS}
    # Keyed per process so the same file open twice is never collapsed into one.
    documents: dict[str, dict[tuple[int, str], DocumentFact]] = {
        name: {} for name in _APPLICATIONS
    }
    native_available = False
    visible_frames: dict[tuple[int, str], int] = {}

    all_windows = _enumerate_all_top_level_windows()
    diagnostics.top_level_windows_total = len(all_windows)
    if dump_window_classes:
        diagnostics.top_level_class_counts = _counter(
            [window[1] or "<unnamed>" for window in all_windows]
        )
    frames = [window for window in all_windows if window[1] in _FRAME_CLASSES]
    diagnostics.office_frame_windows = len(frames)
    diagnostics.frame_class_counts = _counter([window[1] for window in frames])

    for frame_hwnd, frame_class in frames:
        application = _FRAME_CLASSES[frame_class]
        visible, minimized, process_id = _window_state(frame_hwnd)
        identity_known, same_user, same_session = _process_identity(
            process_id, own_sid, own_session
        )
        document_key: str | None = None
        mechanism = _MECHANISM_NONE
        protected = False
        for candidate_hwnd, _candidate_class in _document_windows(
            frame_hwnd, application
        ):
            diagnostics.object_model_attempts += 1
            window_object, mechanism = _object_from_window(candidate_hwnd)
            if window_object is None:
                mechanism = _MECHANISM_NONE
                continue
            diagnostics.object_model_successes += 1
            native_available = True
            reader = _excel_document if application == _EXCEL else _powerpoint_document
            observed = reader(window_object)
            if observed is None:
                continue
            document, protected = observed
            document = replace(document, process_id=process_id)
            documents[application][(process_id, document.document_key)] = document
            document_key = document.document_key
            if visible:
                instance = (process_id, document_key)
                visible_frames[instance] = visible_frames.get(instance, 0) + 1
            if not document.visibility_unproved:
                diagnostics.window_visibility_from_property += 1
            break
        windows[application].append(
            WindowFact(
                application=application,
                process_id=process_id,
                class_name=frame_class,
                visible=visible,
                minimized=minimized,
                same_user=same_user,
                same_session=same_session,
                identity_known=identity_known,
                protected_view=protected,
                document_key=document_key,
                mechanism=mechanism,
            )
        )
    resolved: dict[str, list[DocumentFact]] = {}
    for name, found in documents.items():
        repaired, from_frame, unresolved = resolve_frame_visibility(
            found, visible_frames
        )
        diagnostics.window_visibility_from_frame += from_frame
        diagnostics.window_visibility_unresolved += unresolved
        resolved[name] = list(repaired.values())
    return windows, resolved, native_available


def _spawn_hidden_worker(source: Path | None) -> tuple[Any, Any] | None:
    """Start a hidden QC-owned Excel instance the way the production worker does."""
    if sys.platform != "win32":
        return None
    import win32com.client  # pyright: ignore[reportMissingModuleSource]

    application = win32com.client.DispatchEx("Excel.Application")
    application.Visible = False
    application.AutomationSecurity = 3
    application.EnableEvents = False
    application.DisplayAlerts = False
    application.Interactive = False
    workbook = None
    if source is not None:
        workbook = application.Workbooks.Open(
            str(source), ReadOnly=True, UpdateLinks=0, AddToMru=False
        )
    return application, workbook


def _close_hidden_worker(handle: tuple[Any, Any]) -> None:
    application, workbook = handle
    if workbook is not None:
        with contextlib.suppress(Exception):
            workbook.Close(SaveChanges=False)
    with contextlib.suppress(Exception):
        application.Quit()


def _com_initialize() -> bool:
    """Enter a COM apartment for this thread; report whether we own it."""
    if sys.platform != "win32":
        return False
    import pythoncom  # pyright: ignore[reportMissingModuleSource]

    try:
        pythoncom.CoInitialize()
    except Exception:
        return False
    return True


def _com_uninitialize() -> None:
    if sys.platform != "win32":
        return
    import pythoncom  # pyright: ignore[reportMissingModuleSource]

    with contextlib.suppress(Exception):
        pythoncom.CoUninitialize()


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Prove complete read-only discovery of open Office documents. "
            "Emits aggregate fixed codes and counts only."
        )
    )
    parser.add_argument("--scenario", choices=_SCENARIOS, default="baseline")
    parser.add_argument(
        "--output", type=Path, default=Path("windows-office-discovery.json")
    )
    for application in ("excel", "ppt"):
        for field in ("processes", "documents", "windows"):
            parser.add_argument(
                f"--expect-{application}-{field}", type=int, default=_UNDECLARED
            )
    parser.add_argument("--spawn-hidden-worker", action="store_true")
    parser.add_argument("--hidden-worker-file", type=Path, default=None)
    parser.add_argument("--dump-window-classes", action="store_true")
    return parser


def _expectations(args: argparse.Namespace) -> dict[str, Expectation]:
    return {
        _EXCEL: Expectation(
            processes=args.expect_excel_processes,
            documents=args.expect_excel_documents,
            windows=args.expect_excel_windows,
        ),
        _POWERPOINT: Expectation(
            processes=args.expect_ppt_processes,
            documents=args.expect_ppt_documents,
            windows=args.expect_ppt_windows,
        ),
    }


def _run_discovery(args: argparse.Namespace) -> tuple[int, dict[str, object]]:
    if os.name != "nt":
        return 2, {
            "schema_version": _SCHEMA_VERSION,
            "overall": "unsupported_platform",
            "passed": False,
            "scenario": args.scenario,
        }
    if args.hidden_worker_file is not None and not args.hidden_worker_file.is_file():
        return 2, {
            "schema_version": _SCHEMA_VERSION,
            "overall": "invalid_hidden_worker_source",
            "passed": False,
            "scenario": args.scenario,
        }
    # Excel resolves relative paths against its own directory, not this process's.
    source = (
        args.hidden_worker_file.resolve()
        if args.hidden_worker_file is not None
        else None
    )

    hidden: tuple[Any, Any] | None = None
    diagnostics = Diagnostics()
    started = time.monotonic()
    owns_apartment = _com_initialize()
    try:
        diagnostics.pythoncom_object_from_address = (
            _pythoncom_object_from_address_available()
        )
        diagnostics.comtypes_importable = _comtypes_importable()
        if args.spawn_hidden_worker:
            hidden = _spawn_hidden_worker(source)
        windows, documents, native_available = _collect(
            diagnostics, dump_window_classes=bool(args.dump_window_classes)
        )
        rot_documents, rot_entries, rot_error = _running_object_table_documents()
        diagnostics.running_object_table_entries = rot_entries
        diagnostics.running_object_table_error = rot_error
    finally:
        if hidden is not None:
            _close_hidden_worker(hidden)
        if owns_apartment:
            _com_uninitialize()
    elapsed_seconds = round(time.monotonic() - started, 3)

    expectations = _expectations(args)
    applications: dict[str, dict[str, object]] = {}
    for application in _APPLICATIONS:
        observed_keys = {document.document_key for document in documents[application]}
        rot_only = sum(
            1
            for key, (owner, _path_kind, _extension) in rot_documents.items()
            if owner == application and key not in observed_keys
        )
        applications[application] = summarize_application(
            application,
            windows[application],
            documents[application],
            expectation=expectations[application],
            rot_only_documents=rot_only,
            native_om_available=native_available,
            object_model_attempted=diagnostics.object_model_attempts > 0,
        )

    overall, passed = evaluate(applications)
    payload: dict[str, object] = {
        "schema_version": _SCHEMA_VERSION,
        "applications": applications,
        "diagnostics": diagnostics.as_payload(),
        "elapsed_seconds": elapsed_seconds,
        "hidden_worker_spawned": bool(args.spawn_hidden_worker),
        "overall": overall,
        "passed": passed,
        "running_object_table_documents": len(rot_documents),
        "scenario": args.scenario,
    }
    return (0 if passed else 1), payload


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(list(sys.argv[1:] if argv is None else argv))
    exit_code, payload = _run_discovery(args)
    _write_json(args.output, payload)
    print(f"Windows Office discovery [{payload['scenario']}]: {payload['overall']}")
    print(f"Evidence: {args.output.name}")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
