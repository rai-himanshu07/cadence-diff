"""Primitive request/result vocabulary shared by the focus helper and its parent.

Only fixed codes, counts, and timings cross this boundary in either direction.
No COM object, server session key, or free-form message is ever carried.
"""

from __future__ import annotations

import hashlib
import os
import re
from enum import StrEnum

SCHEMA_VERSION = 1

_FIXED_CODE_RE = re.compile(r"[^a-z0-9_]+")
_MAX_CODE_LENGTH = 64


class FocusAction(StrEnum):
    DISCOVER = "discover"
    FOCUS = "focus"
    HEALTH_CHECK = "health_check"


class FocusStage(StrEnum):
    """Fixed helper stages recorded in a private stage file."""

    STARTING = "starting"
    DISCOVERING = "discovering"
    HASHING = "hashing"
    VALIDATED = "validated"
    ACTION_STARTED = "action_started"
    ACTION_FINISHED = "action_finished"


#: Reaching one of these means an Office call may already be in flight.
DISPATCHED_STAGES = frozenset({FocusStage.ACTION_STARTED, FocusStage.ACTION_FINISHED})


class FocusOutcome(StrEnum):
    """Fixed helper and navigator results."""

    DISCOVERED = "discovered"
    HEALTHY = "healthy"
    FOCUSED = "focused"
    FOCUSED_WITHOUT_FOREGROUND = "focused_without_foreground"
    BOUND_DOCUMENT_NOT_FOUND = "bound_document_not_found"
    TARGET_SHEET_MISSING = "target_sheet_missing"
    TARGET_SHEET_HIDDEN = "target_sheet_hidden"
    TARGET_SLIDE_MISSING = "target_slide_missing"
    TARGET_SHAPE_MISSING = "target_shape_missing"
    TARGET_ADDRESS_INVALID = "target_address_invalid"
    TARGET_WINDOW_MISSING = "target_window_missing"
    SIDE_EFFECT_DETECTED = "side_effect_detected"
    ACTION_UNAVAILABLE = "action_unavailable"
    HELPER_FAILED = "helper_failed"
    HELPER_CRASHED = "helper_crashed"
    INVALID_REQUEST = "invalid_request"
    TIMEOUT_NO_ACTION_DISPATCHED = "timeout_no_action_dispatched"
    TIMEOUT_ACTION_MAY_HAVE_COMPLETED = "timeout_action_may_have_completed"
    FOCUS_DISABLED = "focus_disabled"
    COOLING_DOWN = "cooling_down"
    UNSUPPORTED_PLATFORM = "unsupported_platform"


def fixed_code(value: object, *, fallback: str = "unknown") -> str:
    """Reduce anything to a fixed lowercase code so no path or text can leak."""
    text = value.value if isinstance(value, StrEnum) else str(value)
    cleaned = _FIXED_CODE_RE.sub("_", text.strip().casefold()).strip("_")
    return (cleaned or fallback)[:_MAX_CODE_LENGTH]


def canonical_path(full_name: str) -> str:
    """Case- and separator-normalised path used for every identity digest.

    Backslashes collapse first so a Windows path digests identically on either
    platform; the tests that pin binding continuity then run anywhere.
    """
    return os.path.normcase(os.path.normpath(full_name.strip().replace("\\", "/")))


def path_digest(full_name: str, salt: bytes) -> str:
    """Salted digest so the helper can recognise a path it never receives."""
    digest = hashlib.sha256()
    digest.update(salt)
    digest.update(b"\x00")
    digest.update(canonical_path(full_name).encode("utf-8", "replace"))
    return digest.hexdigest()
