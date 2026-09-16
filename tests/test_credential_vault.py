"""Contracts for ephemeral configuration-workspace credentials."""

from __future__ import annotations

from qc_tool.ui.credential_vault import CredentialVault


def test_credentials_are_bound_to_session_generation_role_and_source_hash() -> None:
    vault = CredentialVault()
    source_hashes = {"current_excel": "a" * 64}
    vault.store_bundle(
        "session-a",
        input_generation=3,
        source_hashes=source_hashes,
        credentials={"current_excel": "secret"},
    )

    assert vault.snapshot("session-a", 3, source_hashes) == {
        "current_excel": "secret"
    }
    assert vault.snapshot("session-a", 2, source_hashes) == {}
    assert vault.snapshot("session-b", 3, source_hashes) == {}
    assert vault.snapshot("session-a", 3, {"current_excel": "b" * 64}) == {}


def test_snapshot_is_a_copy_and_never_exposes_secrets_in_repr() -> None:
    vault = CredentialVault()
    source_hashes = {"current_excel": "a" * 64}
    vault.store_bundle(
        "session-a",
        input_generation=1,
        source_hashes=source_hashes,
        credentials={"current_excel": "private-password"},
    )

    snapshot = vault.snapshot("session-a", 1, source_hashes)
    snapshot["current_excel"] = "changed"

    assert vault.snapshot("session-a", 1, source_hashes)["current_excel"] == (
        "private-password"
    )
    assert "private-password" not in repr(vault)


def test_claim_snapshot_transfers_credentials_only_once() -> None:
    vault = CredentialVault()
    source_hashes = {"current_excel": "a" * 64}
    vault.store_bundle(
        "session-a",
        input_generation=1,
        source_hashes=source_hashes,
        credentials={"current_excel": "secret"},
    )

    assert vault.claim_snapshot("session-a", 1, source_hashes) == {
        "current_excel": "secret"
    }
    assert vault.claim_snapshot("session-a", 1, source_hashes) == {}


def test_clear_role_and_session_remove_only_the_requested_credentials() -> None:
    vault = CredentialVault()
    source_hashes = {
        "baseline_excel": "a" * 64,
        "current_excel": "b" * 64,
    }
    for session_id in ("session-a", "session-b"):
        vault.store_bundle(
            session_id,
            input_generation=1,
            source_hashes=source_hashes,
            credentials={
                "baseline_excel": "baseline-secret",
                "current_excel": "current-secret",
            },
        )

    vault.clear_role("session-a", "baseline_excel")
    assert vault.snapshot("session-a", 1, source_hashes) == {
        "current_excel": "current-secret"
    }
    assert len(vault.snapshot("session-b", 1, source_hashes)) == 2

    vault.clear_session("session-a")
    assert vault.snapshot("session-a", 1, source_hashes) == {}
    assert len(vault.snapshot("session-b", 1, source_hashes)) == 2


def test_dispatch_lease_must_match_generation_and_can_be_revoked() -> None:
    vault = CredentialVault()
    source_hashes = {"current_excel": "a" * 64}
    vault.store_bundle(
        "session-a",
        input_generation=4,
        source_hashes=source_hashes,
        credentials={"current_excel": "secret"},
    )

    lease = vault.acquire_lease("session-a", 4, source_hashes)

    assert lease is not None
    assert vault.lease_is_current(lease)
    assert lease.credentials == {"current_excel": "secret"}
    assert "secret" not in repr(lease)
    vault.release_lease(lease.token)
    assert not vault.lease_is_current(lease)


def test_clear_session_revokes_its_dispatch_leases() -> None:
    vault = CredentialVault()
    source_hashes = {"current_excel": "a" * 64}
    vault.store_bundle(
        "session-a",
        input_generation=1,
        source_hashes=source_hashes,
        credentials={"current_excel": "secret"},
    )
    lease = vault.acquire_lease("session-a", 1, source_hashes)
    assert lease is not None

    vault.clear_session("session-a")

    assert not vault.lease_is_current(lease)
    assert vault.snapshot("session-a", 1, source_hashes) == {}
