"""Server-side focus service: enablement, action tokens, binding, dispatch.

This layer is free of UI imports so every guard is executable without a browser.
A click sends only a single-use token; the run, finding, role, and private target
sidecar are always reloaded here and never trusted from the browser.
"""

from __future__ import annotations

import datetime as dt
import logging
import secrets
import sys
from dataclasses import dataclass, field
from enum import StrEnum
from pathlib import Path

from qc_tool.focus.binding import (
    BindingRegistry,
    BindingRequest,
    BindOutcome,
    FileIdentity,
    resolve_binding,
    revalidate_binding,
)
from qc_tool.focus.discovery import DiscoveryResult, FocusApplication, OpenDocument
from qc_tool.focus.model import FocusRole, FocusTargetSeed, FocusTargetSidecar
from qc_tool.focus.navigator import FocusNavigator, FocusReply
from qc_tool.focus.protocol import SCHEMA_VERSION, FocusAction, FocusOutcome
from qc_tool.history.store import RunRecord
from qc_tool.server_config import NetworkMode

logger = logging.getLogger(__name__)

#: A rendered action token is single use and short lived.
TOKEN_TTL = dt.timedelta(minutes=5)

_APPLICATIONS = {
    FocusRole.BASELINE_EXCEL: FocusApplication.EXCEL,
    FocusRole.CURRENT_EXCEL: FocusApplication.EXCEL,
    FocusRole.BASELINE_PPT: FocusApplication.POWERPOINT,
    FocusRole.CURRENT_PPT: FocusApplication.POWERPOINT,
}
_PAIRED_ROLES = (
    (FocusRole.BASELINE_EXCEL, FocusRole.CURRENT_EXCEL),
    (FocusRole.BASELINE_PPT, FocusRole.CURRENT_PPT),
)

ROLE_LABELS = {
    FocusRole.CURRENT_EXCEL: "current Excel",
    FocusRole.BASELINE_EXCEL: "baseline Excel",
    FocusRole.CURRENT_PPT: "current PowerPoint",
    FocusRole.BASELINE_PPT: "baseline PowerPoint",
}


class FocusUnavailable(StrEnum):
    """Fixed reasons the feature renders no action at all."""

    AVAILABLE = "available"
    FEATURE_OFF = "feature_off"
    UNSUPPORTED_PLATFORM = "unsupported_platform"
    NETWORK_NOT_LOOPBACK = "network_not_loopback"


class TokenRejection(StrEnum):
    UNKNOWN_TOKEN = "unknown_token"
    EXPIRED_TOKEN = "expired_token"
    STALE_RENDER = "stale_render"


@dataclass(frozen=True, slots=True)
class ActionClaim:
    """What one consumed token authorises. Never carries a locator."""

    run_id: int
    finding_id: str
    role: FocusRole
    revision: int
    issued_at: dt.datetime


@dataclass(frozen=True, slots=True)
class BindReport:
    outcome: BindOutcome
    role: FocusRole
    folder_label: str = ""
    unsaved_changes: bool = False

    @property
    def offered(self) -> bool:
        return self.outcome is BindOutcome.MATCHED


@dataclass(slots=True)
class _ClientState:
    revision: int = 0
    acknowledged: bool = False
    tokens: dict[str, ActionClaim] = field(default_factory=dict)
    pending: dict[tuple[int, FocusRole], tuple[OpenDocument, FileIdentity]] = field(
        default_factory=dict
    )


class FocusService:
    """Everything a focus click needs, with no browser-supplied authority."""

    def __init__(
        self,
        work_dir: Path,
        *,
        enabled: bool,
        network_mode: NetworkMode = NetworkMode.LOCAL,
        navigator: FocusNavigator | None = None,
        registry: BindingRegistry | None = None,
        platform: str = sys.platform,
    ) -> None:
        self._work_dir = work_dir
        self._enabled = enabled
        self._network_mode = network_mode
        self._platform = platform
        self._navigator = navigator or FocusNavigator()
        self._registry = registry or BindingRegistry()
        self._clients: dict[str, _ClientState] = {}

    # -- availability ------------------------------------------------------

    def availability(self) -> FocusUnavailable:
        if not self._enabled:
            return FocusUnavailable.FEATURE_OFF
        if self._platform != "win32":
            return FocusUnavailable.UNSUPPORTED_PLATFORM
        if self._network_mode is not NetworkMode.LOCAL:
            return FocusUnavailable.NETWORK_NOT_LOOPBACK
        return FocusUnavailable.AVAILABLE

    @property
    def available(self) -> bool:
        return self.availability() is FocusUnavailable.AVAILABLE

    # -- per-client render and token state ---------------------------------

    def _state(self, client_id: str) -> _ClientState:
        return self._clients.setdefault(client_id, _ClientState())

    def revision(self, client_id: str) -> int:
        return self._state(client_id).revision

    def new_revision(self, client_id: str) -> int:
        """Rendering new content retires every token issued under the old one."""
        state = self._state(client_id)
        state.revision += 1
        state.tokens.clear()
        return state.revision

    def acknowledged(self, client_id: str) -> bool:
        return self._state(client_id).acknowledged

    def acknowledge(self, client_id: str) -> None:
        """One-time acceptance that selection can trigger add-ins or handlers."""
        self._state(client_id).acknowledged = True

    def forget_client(self, client_id: str) -> None:
        self._clients.pop(client_id, None)
        self._registry.invalidate_client(client_id)

    def forget_run(self, run_id: int) -> None:
        for state in self._clients.values():
            for token, claim in list(state.tokens.items()):
                if claim.run_id == run_id:
                    del state.tokens[token]
            for key in [key for key in state.pending if key[0] == run_id]:
                del state.pending[key]
        self._registry.invalidate_run(run_id)

    def issue_token(
        self,
        client_id: str,
        run_id: int,
        finding_id: str,
        role: FocusRole,
        *,
        now: dt.datetime | None = None,
    ) -> str:
        state = self._state(client_id)
        token = secrets.token_urlsafe(32)
        state.tokens[token] = ActionClaim(
            run_id=run_id,
            finding_id=finding_id,
            role=role,
            revision=state.revision,
            issued_at=now or dt.datetime.now(dt.UTC),
        )
        return token

    def consume_token(
        self, client_id: str, token: str, *, now: dt.datetime | None = None
    ) -> ActionClaim | TokenRejection:
        state = self._state(client_id)
        claim = state.tokens.pop(token, None)
        if claim is None:
            return TokenRejection.UNKNOWN_TOKEN
        if claim.revision != state.revision:
            return TokenRejection.STALE_RENDER
        moment = now or dt.datetime.now(dt.UTC)
        if moment - claim.issued_at >= TOKEN_TTL:
            return TokenRejection.EXPIRED_TOKEN
        return claim

    # -- targets -----------------------------------------------------------

    def seeds(self, record: RunRecord, finding_id: str) -> tuple[FocusTargetSeed, ...]:
        """Role seeds that this run can actually act on, ambiguity suppressed."""
        sidecar: FocusTargetSidecar = record.focus_targets
        usable: list[FocusTargetSeed] = []
        for seed in sidecar.seeds(finding_id):
            if not record.file_hashes.get(seed.role.value):
                continue
            usable.append(seed)
        return tuple(seed for seed in usable if seed.role not in self._ambiguous(record))

    @staticmethod
    def _ambiguous(record: RunRecord) -> set[FocusRole]:
        """A role pair whose two inputs hash identically cannot be told apart."""
        ambiguous: set[FocusRole] = set()
        for baseline, current in _PAIRED_ROLES:
            first = record.file_hashes.get(baseline.value)
            second = record.file_hashes.get(current.value)
            if first and second and first == second:
                ambiguous.update({baseline, current})
        return ambiguous

    def seed(
        self, record: RunRecord, finding_id: str, role: FocusRole
    ) -> FocusTargetSeed | None:
        matches = [seed for seed in self.seeds(record, finding_id) if seed.role is role]
        return matches[0] if len(matches) == 1 else None

    def binding(self, client_id: str, run_id: int, role: FocusRole):
        return self._registry.get(client_id, run_id, role)

    # -- dispatch ----------------------------------------------------------

    async def _discover(self, application: FocusApplication) -> DiscoveryResult | None:
        reply = await self._navigator.submit(
            {
                "schema_version": SCHEMA_VERSION,
                "action": FocusAction.DISCOVER.value,
                "application": application.value,
            }
        )
        if reply.outcome is not FocusOutcome.DISCOVERED:
            return None
        return DiscoveryResult.from_payload(reply.discovery)

    async def bind(
        self, client_id: str, record: RunRecord, role: FocusRole
    ) -> BindReport:
        """Discover, match exact saved bytes, and offer one document to confirm."""
        if not self.available:
            return BindReport(BindOutcome.ENUMERATION_INCOMPLETE, role)
        expected = record.file_hashes.get(role.value)
        if not expected or role in self._ambiguous(record):
            return BindReport(BindOutcome.NO_EXACT_MATCH, role)
        discovery = await self._discover(_APPLICATIONS[role])
        if discovery is None:
            return BindReport(BindOutcome.ENUMERATION_INCOMPLETE, role)
        managed = record.file_paths.get(role.value)
        request = BindingRequest(
            run_id=record.run_id,
            role=role,
            expected_sha256=expected,
            managed_root=self._work_dir,
            managed_path=Path(managed) if managed else None,
        )
        outcome, document, identity = resolve_binding(request, discovery)
        if outcome is not BindOutcome.MATCHED or document is None or identity is None:
            return BindReport(outcome, role)
        state = self._state(client_id)
        state.pending[(record.run_id, role)] = (document, identity)
        return BindReport(
            outcome,
            role,
            folder_label=_folder_label(document.full_name),
            unsaved_changes=document.saved is False,
        )

    def confirm(
        self, client_id: str, record: RunRecord, role: FocusRole
    ) -> BindOutcome:
        """Promote an offered candidate into a live binding after confirmation."""
        state = self._state(client_id)
        offered = state.pending.pop((record.run_id, role), None)
        if offered is None:
            return BindOutcome.NO_EXACT_MATCH
        expected = record.file_hashes.get(role.value)
        if not expected:
            return BindOutcome.NO_EXACT_MATCH
        document, identity = offered
        self._registry.bind(
            client_id,
            BindingRequest(
                run_id=record.run_id,
                role=role,
                expected_sha256=expected,
                managed_root=self._work_dir,
            ),
            document,
            identity,
        )
        return BindOutcome.MATCHED

    async def focus(
        self, client_id: str, record: RunRecord, claim: ActionClaim
    ) -> FocusReply:
        """Revalidate the confirmed binding, then dispatch exactly one action."""
        if not self.available:
            return FocusReply(FocusOutcome.UNSUPPORTED_PLATFORM)
        if not self.acknowledged(client_id):
            return FocusReply(FocusOutcome.ACTION_UNAVAILABLE)
        seed = self.seed(record, claim.finding_id, claim.role)
        if seed is None:
            return FocusReply(FocusOutcome.ACTION_UNAVAILABLE)
        binding = self._registry.get(client_id, record.run_id, claim.role)
        if binding is None:
            return FocusReply(BindOutcome.BINDING_EXPIRED.value)
        discovery = await self._discover(_APPLICATIONS[claim.role])
        if discovery is None:
            return FocusReply(BindOutcome.ENUMERATION_INCOMPLETE.value)
        revalidated = revalidate_binding(binding, discovery, self._registry)
        if not revalidated.bound:
            return FocusReply(revalidated.outcome.value)
        salt = secrets.token_bytes(16)
        digest = self._registry.path_digest(
            client_id, record.run_id, claim.role, salt
        )
        if digest is None:
            return FocusReply(BindOutcome.BINDING_IDENTITY_CHANGED.value)
        return await self._navigator.submit(
            {
                "schema_version": SCHEMA_VERSION,
                "action": FocusAction.FOCUS.value,
                "application": _APPLICATIONS[claim.role].value,
                "expected_sha256": binding.expected_sha256,
                "path_salt": salt.hex(),
                "expected_path_digest": digest,
                "process_id": binding.process_id,
                "process_created": binding.process_created,
                "windows_session_id": binding.windows_session_id,
                "window_handle": binding.window_handle,
                "file_id": list(binding.file_id),
                "sheet": seed.sheet,
                "address": seed.address,
                "slide_index": seed.slide_index,
            }
        )


def _folder_label(full_name: str) -> str:
    """Parent folder name only; a full path never reaches the browser."""
    from pathlib import PureWindowsPath

    return PureWindowsPath(full_name.replace("/", "\\")).parent.name
