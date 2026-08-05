"""Complete, read-only discovery of already-open Office documents.

The document readers here are duck-typed over the native object model, so they
run unchanged against mocked COM objects on any platform. Only
``qc_tool.focus.win32_office`` touches ctypes or pywin32, and it is imported
lazily on Windows.

Two rules from the proved Step 1 spike are load-bearing:

* an add-in, hidden-instance, or windowless workbook is never an analyst
  document, or one open workbook presents as four and every action refuses;
* an Office attribute that cannot be read makes discovery incomplete. It never
  silently drops a document.
"""

from __future__ import annotations

import contextlib
import logging
import sys
from collections.abc import Callable, Sequence
from dataclasses import dataclass, replace
from enum import StrEnum

logger = logging.getLogger(__name__)

WindowVisibility = Callable[[int], bool | None]


class FocusApplication(StrEnum):
    EXCEL = "excel"
    POWERPOINT = "powerpoint"


class PathKind(StrEnum):
    LOCAL = "local"
    UNC = "unc"
    URL = "url"
    NONE = "none"
    UNRECOGNIZED = "unrecognized"


class DiscoveryReason(StrEnum):
    """Fixed codes that force ``enumeration_incomplete`` for one application."""

    UNSUPPORTED_PLATFORM = "unsupported_platform"
    NATIVE_OBJECT_MODEL_UNAVAILABLE = "native_object_model_unavailable"
    PROCESS_IDENTITY_UNAVAILABLE = "process_identity_unavailable"
    CROSS_USER_OFFICE_WINDOW = "cross_user_office_window"
    CROSS_SESSION_OFFICE_WINDOW = "cross_session_office_window"
    PROTECTED_VIEW_WINDOW = "protected_view_window"
    DOCUMENT_WINDOW_UNREACHABLE = "document_window_object_model_unreachable"
    WINDOW_VISIBILITY_UNPROVED = "window_visibility_unproved"
    DOCUMENT_WITHOUT_WINDOW = "document_without_window"
    SAVED_STATE_UNPROVED = "saved_state_unproved"
    AUTOSAVE_STATE_UNPROVED = "autosave_state_unproved"
    ROT_DOCUMENT_NOT_ENUMERATED = "running_object_table_document_not_enumerated"


@dataclass(frozen=True, slots=True)
class FrameObservation:
    """One enumerated top-level Office frame window."""

    application: FocusApplication
    process_id: int
    visible: bool
    identity_known: bool
    same_user: bool
    same_session: bool
    protected_view: bool
    reached_document: bool


@dataclass(frozen=True, slots=True)
class OpenDocument:
    """One deduplicated open document, keyed per ``(process, path)``."""

    application: FocusApplication
    process_id: int
    process_created: float
    windows_session_id: int
    full_name: str
    window_count: int
    visible_window_count: int
    visible_window_handles: tuple[int, ...]
    saved: bool | None
    autosave: bool | None
    is_addin: bool = False
    hidden_instance: bool = False
    visibility_unproved: bool = False
    object_model_window_handle: int = 0

    @property
    def instance_key(self) -> tuple[int, str]:
        """Identity is per process and path; path alone hides a duplicate open."""
        return self.process_id, self.full_name.casefold()

    @property
    def path_kind(self) -> PathKind:
        return classify_path_kind(self.full_name)

    @property
    def analyst_candidate(self) -> bool:
        return (
            not self.hidden_instance
            and not self.is_addin
            and self.visible_window_count > 0
        )


@dataclass(frozen=True, slots=True)
class DiscoveryResult:
    application: FocusApplication
    documents: tuple[OpenDocument, ...] = ()
    reasons: tuple[DiscoveryReason, ...] = ()

    @property
    def complete(self) -> bool:
        return not self.reasons

    @property
    def candidates(self) -> tuple[OpenDocument, ...]:
        return tuple(
            document for document in self.documents if document.analyst_candidate
        )

    @classmethod
    def from_payload(cls, payload: object) -> DiscoveryResult | None:
        """Rebuild a helper reply; anything malformed yields ``None``."""
        if not isinstance(payload, dict):
            return None
        try:
            application = FocusApplication(payload.get("application"))
            reasons = tuple(
                DiscoveryReason(item) for item in payload.get("reasons") or ()
            )
            documents = tuple(
                OpenDocument(
                    application=FocusApplication(item["application"]),
                    process_id=int(item["process_id"]),
                    process_created=float(item["process_created"]),
                    windows_session_id=int(item["windows_session_id"]),
                    full_name=str(item["full_name"]),
                    window_count=int(item["window_count"]),
                    visible_window_count=int(item["visible_window_count"]),
                    visible_window_handles=tuple(
                        int(handle) for handle in item["visible_window_handles"]
                    ),
                    saved=item["saved"],
                    autosave=item["autosave"],
                    is_addin=bool(item["is_addin"]),
                    hidden_instance=bool(item["hidden_instance"]),
                    visibility_unproved=bool(item["visibility_unproved"]),
                    object_model_window_handle=int(
                        item.get("object_model_window_handle", 0)
                    ),
                )
                for item in payload.get("documents") or ()
            )
        except (KeyError, TypeError, ValueError):
            return None
        return cls(application=application, documents=documents, reasons=reasons)


def classify_path_kind(full_name: str) -> PathKind:
    text = full_name.strip()
    if not text:
        return PathKind.NONE
    if text[:8].casefold().startswith(("http://", "https://")):
        return PathKind.URL
    if text.startswith("\\\\") or text.startswith("//"):
        return PathKind.UNC
    if len(text) >= 3 and text[1] == ":" and text[2] in "\\/":
        return PathKind.LOCAL
    if text.startswith("/"):
        return PathKind.LOCAL
    return PathKind.UNRECOGNIZED


def read_attribute(source: object, name: str) -> object | None:
    with contextlib.suppress(Exception):
        return getattr(source, name)
    return None


def as_bool(value: object) -> bool | None:
    if value is None:
        return None
    with contextlib.suppress(Exception):
        return bool(value)
    return None


def as_count(value: object) -> int:
    """Late-bound COM counts arrive as ``object``; anything unreadable is zero."""
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    return 0


def _window_handle(window: object) -> int:
    for name in ("HWND", "Hwnd"):
        handle = read_attribute(window, name)
        if isinstance(handle, int) and handle:
            return handle
    return 0


def window_is_visible(window: object, *, window_visible: WindowVisibility) -> bool | None:
    """Excel exposes ``Visible``; a PowerPoint ``DocumentWindow`` does not."""
    flag = as_bool(read_attribute(window, "Visible"))
    if flag is not None:
        return flag
    handle = _window_handle(window)
    if handle:
        return window_visible(handle)
    return None


def read_window_collection(
    collection: object, *, window_visible: WindowVisibility
) -> tuple[int, int, tuple[int, ...], bool]:
    """Return ``(total, visible, visible handles, visibility unproved)``."""
    total = 0
    with contextlib.suppress(Exception):
        total = as_count(read_attribute(collection, "Count"))
    visible = 0
    handles: list[int] = []
    unproved = False
    for index in range(1, total + 1):
        window: object | None = None
        with contextlib.suppress(Exception):
            window = collection.Item(index)  # type: ignore[attr-defined]
        if window is None:
            unproved = True
            continue
        state = window_is_visible(window, window_visible=window_visible)
        if state is None:
            unproved = True
        elif state:
            visible += 1
            handle = _window_handle(window)
            if handle:
                handles.append(handle)
            else:
                unproved = True
    return total, visible, tuple(handles), unproved


def read_excel_document(
    window_object: object,
    *,
    process_id: int,
    process_created: float,
    windows_session_id: int,
    window_visible: WindowVisibility,
) -> tuple[OpenDocument, bool] | None:
    """Read one Excel workbook, plus whether its instance has Protected View."""
    workbook = read_attribute(window_object, "Parent")
    application = read_attribute(window_object, "Application")
    if workbook is None or application is None:
        return None
    full_name = read_attribute(workbook, "FullName")
    if not isinstance(full_name, str) or not full_name.strip():
        return None
    hidden_instance = as_bool(read_attribute(application, "Visible")) is not True
    protected = False
    with contextlib.suppress(Exception):
        collection = read_attribute(application, "ProtectedViewWindows")
        protected = as_count(read_attribute(collection, "Count")) > 0
    total, visible, handles, unproved = read_window_collection(
        read_attribute(workbook, "Windows"), window_visible=window_visible
    )
    document = OpenDocument(
        application=FocusApplication.EXCEL,
        process_id=process_id,
        process_created=process_created,
        windows_session_id=windows_session_id,
        full_name=full_name,
        window_count=total,
        visible_window_count=0 if hidden_instance else visible,
        visible_window_handles=() if hidden_instance else handles,
        saved=as_bool(read_attribute(workbook, "Saved")),
        autosave=as_bool(read_attribute(workbook, "AutoSaveOn")),
        is_addin=as_bool(read_attribute(workbook, "IsAddin")) is True,
        hidden_instance=hidden_instance,
        visibility_unproved=unproved,
    )
    return document, protected


def read_powerpoint_document(
    window_object: object,
    *,
    process_id: int,
    process_created: float,
    windows_session_id: int,
    window_visible: WindowVisibility,
) -> tuple[OpenDocument, bool] | None:
    """Read one presentation; visibility resolves from frames, not properties."""
    presentation = read_attribute(window_object, "Presentation")
    application = read_attribute(window_object, "Application")
    if presentation is None:
        presentation = read_attribute(window_object, "Parent")
    if presentation is None or application is None:
        return None
    full_name = read_attribute(presentation, "FullName")
    if not isinstance(full_name, str) or not full_name.strip():
        return None
    hidden_instance = as_bool(read_attribute(application, "Visible")) is not True
    total, visible, handles, unproved = read_window_collection(
        read_attribute(presentation, "Windows"), window_visible=window_visible
    )
    document = OpenDocument(
        application=FocusApplication.POWERPOINT,
        process_id=process_id,
        process_created=process_created,
        windows_session_id=windows_session_id,
        full_name=full_name,
        window_count=total,
        visible_window_count=0 if hidden_instance else visible,
        visible_window_handles=() if hidden_instance else handles,
        saved=as_bool(read_attribute(presentation, "Saved")),
        autosave=as_bool(read_attribute(presentation, "AutoSaveOn")),
        hidden_instance=hidden_instance,
        visibility_unproved=unproved,
    )
    return document, False


def resolve_frame_visibility(
    documents: Sequence[OpenDocument],
    visible_frames: dict[tuple[int, str], tuple[int, ...]],
) -> tuple[tuple[OpenDocument, ...], int]:
    """Use directly observed visible Office frames when COM cannot prove visibility.

    A PowerPoint ``DocumentWindow`` may expose neither ``Visible`` nor a readable
    handle, so the enumerated top-level frame that reached the document is the
    evidence of record. A document's own window handles always win when it has
    any, because an Excel workbook window is not its ``XLMAIN`` frame.

    Returns the repaired documents and the count still unresolved.
    """
    repaired: list[OpenDocument] = []
    unresolved = 0
    for document in documents:
        proved = document.visible_window_count > 0 and document.visible_window_handles
        if not document.visibility_unproved and proved:
            repaired.append(document)
            continue
        observed = visible_frames.get(document.instance_key, ())
        if not observed:
            if document.visibility_unproved:
                unresolved += 1
            repaired.append(document)
            continue
        handles = document.visible_window_handles or observed
        repaired.append(
            replace(
                document,
                visible_window_count=max(
                    document.visible_window_count, len(set(handles))
                ),
                window_count=max(document.window_count, len(set(handles))),
                visible_window_handles=tuple(sorted(set(handles))),
                visibility_unproved=False,
            )
        )
    return tuple(repaired), unresolved


def assess_discovery(
    application: FocusApplication,
    frames: Sequence[FrameObservation],
    documents: Sequence[OpenDocument],
    *,
    native_object_model_available: bool,
    object_model_attempted: bool = True,
    rot_only_documents: int = 0,
) -> DiscoveryResult:
    """Decide whether discovery for one application is complete."""
    reasons: set[DiscoveryReason] = set()
    if object_model_attempted and not native_object_model_available:
        reasons.add(DiscoveryReason.NATIVE_OBJECT_MODEL_UNAVAILABLE)
    if rot_only_documents > 0:
        reasons.add(DiscoveryReason.ROT_DOCUMENT_NOT_ENUMERATED)
    for frame in frames:
        if frame.application is not application:
            continue
        if not frame.identity_known:
            reasons.add(DiscoveryReason.PROCESS_IDENTITY_UNAVAILABLE)
            continue
        if not frame.same_user:
            reasons.add(DiscoveryReason.CROSS_USER_OFFICE_WINDOW)
        if not frame.same_session:
            reasons.add(DiscoveryReason.CROSS_SESSION_OFFICE_WINDOW)
        if frame.protected_view:
            reasons.add(DiscoveryReason.PROTECTED_VIEW_WINDOW)
        if frame.visible and not frame.reached_document and not frame.protected_view:
            reasons.add(DiscoveryReason.DOCUMENT_WINDOW_UNREACHABLE)
    for document in documents:
        if document.application is not application:
            continue
        if document.visibility_unproved and not document.is_addin:
            reasons.add(DiscoveryReason.WINDOW_VISIBILITY_UNPROVED)
        if not document.analyst_candidate:
            continue
        if document.window_count == 0:
            reasons.add(DiscoveryReason.DOCUMENT_WITHOUT_WINDOW)
        if document.saved is None:
            reasons.add(DiscoveryReason.SAVED_STATE_UNPROVED)
        if document.autosave is None:
            reasons.add(DiscoveryReason.AUTOSAVE_STATE_UNPROVED)
    return DiscoveryResult(
        application=application,
        documents=tuple(
            document for document in documents if document.application is application
        ),
        reasons=tuple(sorted(reasons, key=lambda item: item.value)),
    )


def discover_open_documents(application: FocusApplication) -> DiscoveryResult:
    """Enumerate every same-user, same-session document of one application."""
    if sys.platform != "win32":
        return DiscoveryResult(
            application=application,
            reasons=(DiscoveryReason.UNSUPPORTED_PLATFORM,),
        )
    from qc_tool.focus import win32_office

    return win32_office.discover(application)
