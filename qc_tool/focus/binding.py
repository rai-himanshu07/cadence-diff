"""Exact saved-byte matching and the explicit per-client document binding.

Nothing here activates Office. Binding proves that exactly one eligible, open,
non-managed document is byte-identical to the run's role hash, and records the
deterministic identity attributes a later action revalidates. The binding lives
in server memory only: it is never written to SQLite, reports, or exports.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import hmac
import logging
import os
import secrets
import time
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path, PureWindowsPath

from qc_tool.focus.discovery import (
    DiscoveryReason,
    DiscoveryResult,
    OpenDocument,
    PathKind,
)
from qc_tool.focus.model import FocusRole
from qc_tool.focus.package_risk import PackageScanError, scan_focus_package
from qc_tool.focus.protocol import canonical_path, path_digest

logger = logging.getLogger(__name__)

#: A confirmed binding expires on its own, independently of every other trigger.
BINDING_TTL = dt.timedelta(minutes=10)

#: Bounded read budget for hashing one candidate, including slow UNC sources.
HASH_DEADLINE_SECONDS = 20.0

_HASH_CHUNK = 1 << 20


class BindOutcome(StrEnum):
    """Fixed binding results. Never a path, hash, value, or COM object name."""

    MATCHED = "matched"
    NO_EXACT_MATCH = "no_exact_match"
    AMBIGUOUS_EXACT_MATCHES = "ambiguous_exact_matches"
    MATCHING_DOCUMENT_DIRTY_AND_CHANGED = "matching_document_dirty_and_changed"
    UNSUPPORTED_ACTIVE_CONTENT = "unsupported_active_content"
    UNSUPPORTED_FORMAT = "unsupported_format"
    UNSUPPORTED_LOCATION = "unsupported_location"
    ENCRYPTED_DOCUMENT_UNSUPPORTED = "encrypted_document_unsupported"
    PROTECTED_VIEW = "protected_view"
    ENUMERATION_INCOMPLETE = "enumeration_incomplete"
    DOCUMENT_CHANGING = "document_changing"
    HASH_TIMEOUT = "hash_timeout"
    PACKAGE_UNREADABLE = "package_unreadable"
    MANAGED_COPY_REFUSED = "managed_copy_refused"
    WINDOW_AMBIGUOUS = "window_ambiguous"
    AUTOSAVE_ENABLED = "autosave_enabled"
    AUTOSAVE_STATE_UNPROVED = "autosave_state_unproved"
    SAVED_STATE_UNPROVED = "saved_state_unproved"
    BINDING_EXPIRED = "binding_expired"
    BINDING_IDENTITY_CHANGED = "binding_identity_changed"


#: Reported when every hash-matching candidate was refused. Most specific first.
_REFUSAL_PRECEDENCE: tuple[BindOutcome, ...] = (
    BindOutcome.PROTECTED_VIEW,
    BindOutcome.ENCRYPTED_DOCUMENT_UNSUPPORTED,
    BindOutcome.UNSUPPORTED_ACTIVE_CONTENT,
    BindOutcome.MANAGED_COPY_REFUSED,
    BindOutcome.AUTOSAVE_ENABLED,
    BindOutcome.AUTOSAVE_STATE_UNPROVED,
    BindOutcome.SAVED_STATE_UNPROVED,
    BindOutcome.DOCUMENT_CHANGING,
    BindOutcome.HASH_TIMEOUT,
    BindOutcome.PACKAGE_UNREADABLE,
    BindOutcome.UNSUPPORTED_FORMAT,
    BindOutcome.UNSUPPORTED_LOCATION,
)

_SCAN_ERROR_OUTCOMES = {
    PackageScanError.ACTIVE_CONTENT_FORMAT: BindOutcome.UNSUPPORTED_ACTIVE_CONTENT,
    PackageScanError.UNSUPPORTED_FORMAT: BindOutcome.UNSUPPORTED_FORMAT,
    PackageScanError.ENCRYPTED_PACKAGE: BindOutcome.ENCRYPTED_DOCUMENT_UNSUPPORTED,
    PackageScanError.MALFORMED_PACKAGE: BindOutcome.PACKAGE_UNREADABLE,
    PackageScanError.UNREADABLE_PACKAGE: BindOutcome.PACKAGE_UNREADABLE,
    PackageScanError.PACKAGE_TOO_LARGE: BindOutcome.PACKAGE_UNREADABLE,
}


class HashTimeoutError(RuntimeError):
    """Hashing a candidate exceeded its bounded read budget."""


class FileChangedError(RuntimeError):
    """The file identity or stat changed while it was being hashed."""


@dataclass(frozen=True, slots=True)
class FileIdentity:
    """Stable filesystem identity captured from one open handle."""

    device: int
    index: int
    size: int
    modified_ns: int

    @property
    def file_id(self) -> tuple[int, int]:
        return self.device, self.index


@dataclass(frozen=True, slots=True)
class DocumentBinding:
    """Deterministic identity of one explicitly confirmed open document."""

    run_id: int
    role: FocusRole
    expected_sha256: str
    process_id: int
    process_created: float
    windows_session_id: int
    window_handle: int
    canonical_path_hmac: str
    file_id: tuple[int, int]
    bound_at: dt.datetime
    unsaved_changes: bool = False


@dataclass(frozen=True, slots=True)
class BindingRequest:
    run_id: int
    role: FocusRole
    expected_sha256: str
    managed_root: Path
    managed_path: Path | None = None


@dataclass(frozen=True, slots=True)
class BindingResult:
    outcome: BindOutcome
    binding: DocumentBinding | None = None

    @property
    def bound(self) -> bool:
        return self.outcome is BindOutcome.MATCHED and self.binding is not None


def _identity_from_stat(status: os.stat_result) -> FileIdentity:
    return FileIdentity(
        device=int(status.st_dev),
        index=int(status.st_ino),
        size=int(status.st_size),
        modified_ns=int(status.st_mtime_ns),
    )


def hash_open_file(
    path: Path, *, deadline_seconds: float = HASH_DEADLINE_SECONDS
) -> tuple[str, FileIdentity]:
    """Hash a saved document from one handle and prove it did not change.

    Raises ``HashTimeoutError`` when the bounded read budget is exceeded and
    ``FileChangedError`` when the file identity or stat moved underneath us.
    """
    started = time.monotonic()
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        before = _identity_from_stat(os.fstat(handle.fileno()))
        while True:
            if time.monotonic() - started > deadline_seconds:
                raise HashTimeoutError("focus hash budget exceeded")
            chunk = handle.read(_HASH_CHUNK)
            if not chunk:
                break
            digest.update(chunk)
        after = _identity_from_stat(os.fstat(handle.fileno()))
    if before != after:
        raise FileChangedError("the document changed while it was being hashed")
    return digest.hexdigest(), after


def is_managed_copy(
    path: Path, managed_root: Path, *, managed_path: Path | None = None
) -> bool:
    """Strict resolved containment, plus file identity for junction aliases."""
    try:
        resolved = path.resolve(strict=True)
    except OSError:
        return False
    try:
        resolved.relative_to(managed_root.resolve())
        return True
    except (OSError, ValueError):
        pass
    if managed_path is None:
        return False
    try:
        candidate = resolved.stat()
        managed = managed_path.resolve(strict=True).stat()
    except OSError:
        return False
    return (candidate.st_dev, candidate.st_ino) == (managed.st_dev, managed.st_ino)


def _discovery_outcome(discovery: DiscoveryResult) -> BindOutcome:
    if DiscoveryReason.PROTECTED_VIEW_WINDOW in discovery.reasons:
        return BindOutcome.PROTECTED_VIEW
    return BindOutcome.ENUMERATION_INCOMPLETE


def _eligibility_refusal(
    document: OpenDocument, request: BindingRequest
) -> BindOutcome | None:
    """Policy refusal for one candidate, evaluated before any hashing."""
    path_kind = document.path_kind
    if path_kind not in {PathKind.LOCAL, PathKind.UNC, PathKind.URL}:
        return BindOutcome.UNSUPPORTED_LOCATION
    if document.autosave is None:
        return BindOutcome.AUTOSAVE_STATE_UNPROVED
    if document.autosave:
        return BindOutcome.AUTOSAVE_ENABLED
    if path_kind is PathKind.URL:
        return BindOutcome.UNSUPPORTED_LOCATION
    if document.saved is None:
        return BindOutcome.SAVED_STATE_UNPROVED
    path = Path(document.full_name)
    if is_managed_copy(path, request.managed_root, managed_path=request.managed_path):
        return BindOutcome.MANAGED_COPY_REFUSED
    scan = scan_focus_package(path)
    if scan.error is not None:
        return _SCAN_ERROR_OUTCOMES[scan.error]
    if scan.risks:
        return BindOutcome.UNSUPPORTED_ACTIVE_CONTENT
    return None


def _highest_refusal(refusals: set[BindOutcome]) -> BindOutcome:
    for outcome in _REFUSAL_PRECEDENCE:
        if outcome in refusals:
            return outcome
    return BindOutcome.NO_EXACT_MATCH


def resolve_binding(
    request: BindingRequest,
    discovery: DiscoveryResult,
    *,
    hash_deadline_seconds: float = HASH_DEADLINE_SECONDS,
) -> tuple[BindOutcome, OpenDocument | None, FileIdentity | None]:
    """Apply the fixed resolution precedence to one discovery snapshot.

    1. complete discovery, 2. deduplicated identities, 3. eligibility and exact
    hash across every process, 4. zero or many matches refuse before any window
    is inspected, 5. only the sole match needs one visible window.
    """
    if not discovery.complete:
        return _discovery_outcome(discovery), None, None
    refusals: set[BindOutcome] = set()
    matches: list[tuple[OpenDocument, FileIdentity]] = []
    for document in discovery.candidates:
        refusal = _eligibility_refusal(document, request)
        if refusal is not None:
            refusals.add(refusal)
            continue
        path = Path(document.full_name)
        try:
            digest, identity = hash_open_file(
                path, deadline_seconds=hash_deadline_seconds
            )
        except HashTimeoutError:
            refusals.add(BindOutcome.HASH_TIMEOUT)
            continue
        except FileChangedError:
            refusals.add(BindOutcome.DOCUMENT_CHANGING)
            continue
        except OSError:
            refusals.add(BindOutcome.PACKAGE_UNREADABLE)
            continue
        if not hmac.compare_digest(digest, request.expected_sha256):
            continue
        matches.append((document, identity))
    if len(matches) > 1:
        return BindOutcome.AMBIGUOUS_EXACT_MATCHES, None, None
    if not matches:
        if refusals:
            return _highest_refusal(refusals), None, None
        return BindOutcome.NO_EXACT_MATCH, None, None
    document, identity = matches[0]
    handles = set(document.visible_window_handles)
    if document.visible_window_count != 1 or len(handles) != 1:
        return BindOutcome.WINDOW_AMBIGUOUS, None, None
    return BindOutcome.MATCHED, document, identity


class BindingRegistry:
    """Per-client, in-memory bindings. Never persisted and never exported."""

    def __init__(self, *, ttl: dt.timedelta = BINDING_TTL) -> None:
        self._ttl = ttl
        self._key = secrets.token_bytes(32)
        self._bindings: dict[tuple[str, int, FocusRole], DocumentBinding] = {}
        # Private server state. Never returned to a helper, a log, or the UI.
        self._paths: dict[tuple[str, int, FocusRole], str] = {}

    def path_hmac(self, full_name: str) -> str:
        """Server-keyed digest of a canonical path; helpers never see the key."""
        canonical = canonical_path(full_name)
        return hmac.new(self._key, canonical.encode("utf-8"), hashlib.sha256).hexdigest()

    def bind(
        self,
        client_id: str,
        request: BindingRequest,
        document: OpenDocument,
        identity: FileIdentity,
        *,
        now: dt.datetime | None = None,
    ) -> DocumentBinding:
        handles = sorted(set(document.visible_window_handles))
        if len(handles) != 1:
            raise ValueError("a binding requires exactly one visible window")
        binding = DocumentBinding(
            run_id=request.run_id,
            role=request.role,
            expected_sha256=request.expected_sha256,
            process_id=document.process_id,
            process_created=document.process_created,
            windows_session_id=document.windows_session_id,
            window_handle=handles[0],
            canonical_path_hmac=self.path_hmac(document.full_name),
            file_id=identity.file_id,
            bound_at=now or dt.datetime.now(dt.UTC),
            unsaved_changes=document.saved is False,
        )
        key = (client_id, request.run_id, request.role)
        self._bindings[key] = binding
        self._paths[key] = document.full_name
        return binding

    def path_digest(
        self, client_id: str, run_id: int, role: FocusRole, salt: bytes
    ) -> str | None:
        """Salted digest of the bound path, so a helper never receives the path."""
        full_name = self._paths.get((client_id, run_id, role))
        if full_name is None:
            return None
        return path_digest(full_name, salt)

    def folder_label(
        self, client_id: str, run_id: int, role: FocusRole
    ) -> str:
        """Privacy-minimised parent-folder name shown at confirmation time."""
        full_name = self._paths.get((client_id, run_id, role))
        if not full_name:
            return ""
        return PureWindowsPath(full_name.replace("/", "\\")).parent.name

    def get(
        self,
        client_id: str,
        run_id: int,
        role: FocusRole,
        *,
        now: dt.datetime | None = None,
    ) -> DocumentBinding | None:
        key = (client_id, run_id, role)
        binding = self._bindings.get(key)
        if binding is None:
            return None
        moment = now or dt.datetime.now(dt.UTC)
        if moment - binding.bound_at >= self._ttl:
            del self._bindings[key]
            self._paths.pop(key, None)
            return None
        return binding

    def invalidate_client(self, client_id: str) -> None:
        for key in [key for key in self._bindings if key[0] == client_id]:
            del self._bindings[key]
            self._paths.pop(key, None)

    def invalidate_run(self, run_id: int) -> None:
        for key in [key for key in self._bindings if key[1] == run_id]:
            del self._bindings[key]
            self._paths.pop(key, None)

    def clear(self) -> None:
        self._bindings.clear()
        self._paths.clear()

    def __len__(self) -> int:
        return len(self._bindings)


def revalidate_binding(
    binding: DocumentBinding,
    discovery: DiscoveryResult,
    registry: BindingRegistry,
    *,
    hash_deadline_seconds: float = HASH_DEADLINE_SECONDS,
) -> BindingResult:
    """Re-prove a confirmed binding immediately before any Office action."""
    if not discovery.complete:
        return BindingResult(_discovery_outcome(discovery))
    expected_hmac = binding.canonical_path_hmac
    for document in discovery.candidates:
        if document.process_id != binding.process_id:
            continue
        if registry.path_hmac(document.full_name) != expected_hmac:
            continue
        if (
            document.process_created != binding.process_created
            or document.windows_session_id != binding.windows_session_id
            or binding.window_handle not in set(document.visible_window_handles)
        ):
            return BindingResult(BindOutcome.BINDING_IDENTITY_CHANGED)
        if document.autosave is not False or document.saved is None:
            return BindingResult(BindOutcome.BINDING_IDENTITY_CHANGED)
        path = Path(document.full_name)
        try:
            digest, identity = hash_open_file(
                path, deadline_seconds=hash_deadline_seconds
            )
        except HashTimeoutError:
            return BindingResult(BindOutcome.HASH_TIMEOUT)
        except FileChangedError:
            return BindingResult(BindOutcome.DOCUMENT_CHANGING)
        except OSError:
            return BindingResult(BindOutcome.PACKAGE_UNREADABLE)
        if identity.file_id != binding.file_id:
            return BindingResult(BindOutcome.BINDING_IDENTITY_CHANGED)
        if not hmac.compare_digest(digest, binding.expected_sha256):
            if document.saved is False:
                return BindingResult(BindOutcome.MATCHING_DOCUMENT_DIRTY_AND_CHANGED)
            return BindingResult(BindOutcome.DOCUMENT_CHANGING)
        return BindingResult(BindOutcome.MATCHED, binding)
    return BindingResult(BindOutcome.BINDING_IDENTITY_CHANGED)
