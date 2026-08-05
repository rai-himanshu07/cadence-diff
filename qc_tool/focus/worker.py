"""Short-lived Windows focus helper. One COM apartment, one action, then exit.

The helper is the only place that talks to an analyst-owned Office process. It
never opens, saves, recalculates, refreshes, closes, or quits a document, and it
never mutates a global Office setting. It reads a primitive JSON request from
stdin, records fixed stages in a private file so the parent can tell a
pre-dispatch timeout from a post-dispatch one, and writes a primitive result
file containing fixed codes only.
"""

from __future__ import annotations

import contextlib
import importlib
import json
import os
import sys
from dataclasses import asdict
from pathlib import Path

from qc_tool.focus.binding import (
    BindOutcome,
    FileChangedError,
    HashTimeoutError,
    hash_open_file,
)
from qc_tool.focus.discovery import (
    DiscoveryResult,
    FocusApplication,
    OpenDocument,
    discover_open_documents,
)
from qc_tool.focus.protocol import (
    SCHEMA_VERSION,
    FocusAction,
    FocusOutcome,
    FocusStage,
    fixed_code,
    path_digest,
)


def _write_stage(stage_path: Path | None, stage: FocusStage) -> None:
    if stage_path is None:
        return
    with contextlib.suppress(OSError):
        temporary = stage_path.with_suffix(".tmp")
        temporary.write_text(stage.value, encoding="utf-8")
        os.replace(temporary, stage_path)


def _document_payload(document: OpenDocument) -> dict[str, object]:
    payload = asdict(document)
    payload["application"] = document.application.value
    payload["path_kind"] = document.path_kind.value
    return payload


def _discovery_payload(discovery: DiscoveryResult) -> dict[str, object]:
    return {
        "application": discovery.application.value,
        "documents": [_document_payload(item) for item in discovery.documents],
        "reasons": [reason.value for reason in discovery.reasons],
    }


def _int_field(request: dict[str, object], name: str) -> int | None:
    value = request.get(name)
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    return None


def _str_field(request: dict[str, object], name: str) -> str | None:
    value = request.get(name)
    return value if isinstance(value, str) and value else None


def locate_bound_document(
    discovery: DiscoveryResult, request: dict[str, object]
) -> OpenDocument | None:
    """Find the single document matching every bound identity attribute."""
    salt_hex = _str_field(request, "path_salt")
    expected_digest = _str_field(request, "expected_path_digest")
    process_id = _int_field(request, "process_id")
    window_handle = _int_field(request, "window_handle")
    session_id = _int_field(request, "windows_session_id")
    created = request.get("process_created")
    if (
        salt_hex is None
        or expected_digest is None
        or process_id is None
        or window_handle is None
        or session_id is None
        or not isinstance(created, int | float)
        or isinstance(created, bool)
    ):
        return None
    try:
        salt = bytes.fromhex(salt_hex)
    except ValueError:
        return None
    for document in discovery.candidates:
        if document.process_id != process_id:
            continue
        if document.process_created != float(created):
            continue
        if document.windows_session_id != session_id:
            continue
        if window_handle not in set(document.visible_window_handles):
            continue
        if path_digest(document.full_name, salt) != expected_digest:
            continue
        return document
    return None


def _perform_focus(
    discovery: DiscoveryResult,
    request: dict[str, object],
    stage_path: Path | None,
) -> dict[str, object]:
    document = locate_bound_document(discovery, request)
    if document is None:
        return {"outcome": FocusOutcome.BOUND_DOCUMENT_NOT_FOUND.value}
    expected = _str_field(request, "expected_sha256")
    if expected is None or document.autosave is not False or document.saved is None:
        return {"outcome": BindOutcome.BINDING_IDENTITY_CHANGED.value}
    _write_stage(stage_path, FocusStage.HASHING)
    source = Path(document.full_name)
    try:
        before, identity = hash_open_file(source)
    except HashTimeoutError:
        return {"outcome": BindOutcome.HASH_TIMEOUT.value}
    except (FileChangedError, OSError):
        return {"outcome": BindOutcome.DOCUMENT_CHANGING.value}
    if before != expected:
        return {
            "outcome": (
                BindOutcome.MATCHING_DOCUMENT_DIRTY_AND_CHANGED.value
                if document.saved is False
                else BindOutcome.DOCUMENT_CHANGING.value
            )
        }
    file_id = _int_pair(request.get("file_id"))
    if file_id is not None and file_id != identity.file_id:
        return {"outcome": BindOutcome.BINDING_IDENTITY_CHANGED.value}
    _write_stage(stage_path, FocusStage.VALIDATED)
    action = _focus_action(document.application)
    if action is None:
        return {"outcome": FocusOutcome.ACTION_UNAVAILABLE.value}
    _write_stage(stage_path, FocusStage.ACTION_STARTED)
    try:
        outcome = action(document, request)
    except Exception:
        outcome = FocusOutcome.HELPER_FAILED.value
    _write_stage(stage_path, FocusStage.ACTION_FINISHED)
    reply: dict[str, object] = {
        "outcome": fixed_code(outcome),
        "unsaved_changes": document.saved is False,
    }
    try:
        after, _identity = hash_open_file(source)
    except (HashTimeoutError, FileChangedError, OSError):
        reply["outcome"] = FocusOutcome.SIDE_EFFECT_DETECTED.value
        return reply
    if after != before:
        reply["outcome"] = FocusOutcome.SIDE_EFFECT_DETECTED.value
    return reply


def _int_pair(value: object) -> tuple[int, int] | None:
    if not isinstance(value, list) or len(value) != 2:
        return None
    first, second = value
    if isinstance(first, int) and isinstance(second, int):
        return int(first), int(second)
    return None


def _focus_action(application: FocusApplication):
    """Resolve the navigation implementation for one application, if any."""
    module = (
        "qc_tool.focus.excel_focus"
        if application is FocusApplication.EXCEL
        else "qc_tool.focus.ppt_focus"
    )
    try:
        implementation = importlib.import_module(module)
    except ImportError:
        return None
    return getattr(implementation, "focus_document", None)


def handle_request(
    request: dict[str, object], *, stage_path: Path | None = None
) -> dict[str, object]:
    """Execute one primitive request and return its primitive reply."""
    _write_stage(stage_path, FocusStage.STARTING)
    if request.get("schema_version") != SCHEMA_VERSION:
        return {"outcome": FocusOutcome.INVALID_REQUEST.value}
    raw_action = request.get("action")
    try:
        action = FocusAction(raw_action)
    except ValueError:
        return {"outcome": FocusOutcome.INVALID_REQUEST.value}
    try:
        application = FocusApplication(request.get("application"))
    except ValueError:
        return {"outcome": FocusOutcome.INVALID_REQUEST.value}
    _write_stage(stage_path, FocusStage.DISCOVERING)
    discovery = discover_open_documents(application)
    if action is FocusAction.HEALTH_CHECK:
        return {
            "outcome": FocusOutcome.HEALTHY.value,
            "complete": discovery.complete,
            "reasons": [reason.value for reason in discovery.reasons],
        }
    if action is FocusAction.DISCOVER:
        return {
            "outcome": FocusOutcome.DISCOVERED.value,
            "discovery": _discovery_payload(discovery),
        }
    return _perform_focus(discovery, request, stage_path)


def main(argv: list[str] | None = None) -> int:
    del argv
    try:
        request = json.loads(sys.stdin.read() or "{}")
    except json.JSONDecodeError:
        request = {}
    if not isinstance(request, dict):
        request = {}
    raw_stage = request.get("stage_path")
    stage_path = Path(raw_stage) if isinstance(raw_stage, str) and raw_stage else None
    raw_result = request.get("result_path")
    result_path = (
        Path(raw_result) if isinstance(raw_result, str) and raw_result else None
    )
    owns_apartment = _com_initialize()
    try:
        reply: dict[str, object] = handle_request(request, stage_path=stage_path)
    except Exception:
        reply = {"outcome": FocusOutcome.HELPER_FAILED.value}
    finally:
        if owns_apartment:
            _com_uninitialize()
    reply["schema_version"] = SCHEMA_VERSION
    payload = json.dumps(reply)
    if result_path is not None:
        with contextlib.suppress(OSError):
            temporary = result_path.with_suffix(".tmp")
            temporary.write_text(payload, encoding="utf-8")
            os.replace(temporary, result_path)
    sys.stdout.write(payload)
    return 0


def _com_initialize() -> bool:
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


# A spawn child re-imports the parent's __main__ as __mp_main__; guarding on
# "__main__" alone keeps this helper from ever booting a second server.
if __name__ == "__main__":
    raise SystemExit(main())
