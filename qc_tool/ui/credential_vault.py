"""Ephemeral credentials for one configuration-workspace server process."""

from __future__ import annotations

import threading
import uuid
from dataclasses import dataclass, field


@dataclass(frozen=True, slots=True)
class CredentialLease:
    """A revocable credential copy for one imminent worker dispatch."""

    token: str
    session_id: str
    input_generation: int
    credentials: dict[str, str] = field(repr=False, compare=False)


class CredentialVault:
    """Keep passwords in memory and bind them to exact source generations."""

    def __init__(self) -> None:
        self._entries: dict[tuple[str, int, str, str], str] = {}
        self._leases: dict[str, tuple[str, int]] = {}
        self._lock = threading.RLock()

    def __repr__(self) -> str:
        with self._lock:
            sessions = {key[0] for key in self._entries}
            return (
                f"CredentialVault(sessions={len(sessions)}, "
                f"entries={len(self._entries)}, leases={len(self._leases)})"
            )

    def store_bundle(
        self,
        session_id: str,
        *,
        input_generation: int,
        source_hashes: dict[str, str],
        credentials: dict[str, str],
    ) -> None:
        """Replace one session's credentials with a source-bound bundle."""
        with self._lock:
            self._clear_session_locked(session_id)
            for role, password in credentials.items():
                source_hash = source_hashes.get(role)
                if password and source_hash:
                    self._entries[
                        (session_id, input_generation, role, source_hash)
                    ] = password

    def set(
        self,
        session_id: str,
        *,
        input_generation: int,
        role: str,
        source_sha256: str,
        password: str,
    ) -> None:
        """Set or clear one role after an inline credential prompt."""
        with self._lock:
            self._clear_role_locked(session_id, role)
            if password:
                self._entries[
                    (session_id, input_generation, role, source_sha256)
                ] = password

    def snapshot(
        self,
        session_id: str,
        input_generation: int,
        source_hashes: dict[str, str],
    ) -> dict[str, str]:
        """Return a copy containing only credentials for this exact generation."""
        with self._lock:
            return {
                role: password
                for role, source_hash in source_hashes.items()
                if (
                    password := self._entries.get(
                        (session_id, input_generation, role, source_hash)
                    )
                )
            }

    def claim_snapshot(
        self,
        session_id: str,
        input_generation: int,
        source_hashes: dict[str, str],
    ) -> dict[str, str]:
        """Transfer credentials to one page and remove the handoff copy."""
        with self._lock:
            credentials = self.snapshot(session_id, input_generation, source_hashes)
            self._clear_session_locked(session_id)
            return credentials

    def acquire_lease(
        self,
        session_id: str,
        input_generation: int,
        source_hashes: dict[str, str],
    ) -> CredentialLease | None:
        """Copy matching credentials into a revocable pre-dispatch lease."""
        with self._lock:
            credentials = self.snapshot(session_id, input_generation, source_hashes)
            if not credentials:
                return None
            token = uuid.uuid4().hex
            self._leases[token] = (session_id, input_generation)
            return CredentialLease(
                token=token,
                session_id=session_id,
                input_generation=input_generation,
                credentials=credentials,
            )

    def lease_is_current(self, lease: CredentialLease) -> bool:
        with self._lock:
            return self._leases.get(lease.token) == (
                lease.session_id,
                lease.input_generation,
            )

    def release_lease(self, token: str) -> None:
        with self._lock:
            self._leases.pop(token, None)

    def clear_role(self, session_id: str, role: str) -> None:
        with self._lock:
            self._clear_role_locked(session_id, role)
            self._revoke_session_leases_locked(session_id)

    def clear_session(self, session_id: str) -> None:
        with self._lock:
            self._clear_session_locked(session_id)

    def _clear_role_locked(self, session_id: str, role: str) -> None:
        for key in tuple(self._entries):
            if key[0] == session_id and key[2] == role:
                self._entries.pop(key, None)

    def _clear_session_locked(self, session_id: str) -> None:
        for key in tuple(self._entries):
            if key[0] == session_id:
                self._entries.pop(key, None)
        self._revoke_session_leases_locked(session_id)

    def _revoke_session_leases_locked(self, session_id: str) -> None:
        for token, lease_key in tuple(self._leases.items()):
            if lease_key[0] == session_id:
                self._leases.pop(token, None)
