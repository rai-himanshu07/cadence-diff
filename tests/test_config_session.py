"""Tests for the Step 5 configuration-session store
(``qc_tool.history.config_session``).

Plan: docs/plans/plan-20260913-mode-aware-configuration-wizard.md, Step 5.
"""

from __future__ import annotations

import datetime as dt
import json
import sqlite3
from pathlib import Path

import pytest

from qc_tool.history.config_session import (
    ConfigSessionConflictError,
    ConfigSessionStore,
    session_key_for,
    source_set_digest_for,
)


def test_session_key_is_stable_and_order_independent() -> None:
    a = session_key_for({"baseline_excel": "a" * 64, "current_excel": "b" * 64})
    b = session_key_for({"current_excel": "b" * 64, "baseline_excel": "a" * 64})
    assert a == b


def test_session_key_differs_for_different_file_sets() -> None:
    first = session_key_for({"current_excel": "a" * 64})
    second = session_key_for({"current_excel": "b" * 64})
    assert first != second


def test_session_key_never_contains_the_input_hashes_verbatim() -> None:
    key = session_key_for({"current_excel": "a" * 64})
    assert "a" * 64 not in key


def test_create_session_uses_random_ids_and_a_stable_source_set_digest(
    tmp_path: Path,
) -> None:
    store = ConfigSessionStore(tmp_path / "sessions.sqlite3")
    file_hashes = {"current_excel": "a" * 64}

    first = store.create_session(
        file_hashes=file_hashes,
        profile_name="default",
        choices={"mode": "current_file_preflight", "file_hashes": file_hashes},
    )
    second = store.create_session(
        file_hashes=file_hashes,
        profile_name="default",
        choices={"mode": "current_file_preflight", "file_hashes": file_hashes},
    )

    assert first.session_id != second.session_id
    assert first.source_set_digest == second.source_set_digest
    assert first.source_set_digest == source_set_digest_for(file_hashes)
    assert first.session_id != first.source_set_digest
    assert "a" * 64 not in first.session_id


def test_legacy_session_row_is_backfilled_and_resolves_by_alias_or_canonical_id(
    tmp_path: Path,
) -> None:
    db_path = tmp_path / "sessions.sqlite3"
    legacy_key = session_key_for({"current_excel": "a" * 64})
    now = dt.datetime.now(dt.UTC).isoformat(timespec="milliseconds")
    with sqlite3.connect(db_path) as connection:
        connection.executescript(
            """
            CREATE TABLE config_sessions (
                session_key TEXT PRIMARY KEY,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                profile_name TEXT NOT NULL DEFAULT '',
                choices TEXT NOT NULL DEFAULT '{}'
            );
            """
        )
        connection.execute(
            "INSERT INTO config_sessions VALUES (?, ?, ?, ?, ?)",
            (
                legacy_key,
                now,
                now,
                "default",
                json.dumps(
                    {
                        "mode": "current_file_preflight",
                        "file_hashes": {"current_excel": "a" * 64},
                    }
                ),
            ),
        )

    store = ConfigSessionStore(db_path)
    migrated = store.get(legacy_key)

    assert migrated is not None
    assert migrated.session_key == legacy_key
    assert migrated.session_id != legacy_key
    assert migrated.source_set_digest == legacy_key
    assert store.get(migrated.session_id) == migrated


def test_revision_checked_save_rejects_a_stale_writer(tmp_path: Path) -> None:
    store = ConfigSessionStore(tmp_path / "sessions.sqlite3")
    created = store.create_session(
        file_hashes={"current_excel": "a" * 64},
        choices={"mode": "current_file_preflight"},
    )
    updated = store.save_choices(
        created.session_id,
        choices={"mode": "current_file_preflight", "step": 2},
        expected_revision=created.revision,
    )

    assert updated.revision == created.revision + 1
    with pytest.raises(ConfigSessionConflictError):
        store.save_choices(
            created.session_id,
            choices={"mode": "current_file_preflight", "step": 3},
            expected_revision=created.revision,
        )


def test_input_generation_changes_only_when_run_inputs_change(tmp_path: Path) -> None:
    store = ConfigSessionStore(tmp_path / "sessions.sqlite3")
    first_hashes = {"current_excel": "a" * 64}
    created = store.create_session(
        file_hashes=first_hashes,
        choices={"mode": "current_file_preflight", "file_hashes": first_hashes},
    )
    decisions_only = store.save_choices(
        created.session_id,
        choices={
            "mode": "current_file_preflight",
            "file_hashes": first_hashes,
            "region_decisions": {"r1": {"confirmed": True}},
        },
        expected_revision=created.revision,
    )
    changed_hashes = {"current_excel": "b" * 64}
    changed_input = store.save_choices(
        created.session_id,
        choices={"mode": "current_file_preflight", "file_hashes": changed_hashes},
        expected_revision=decisions_only.revision,
    )

    assert decisions_only.input_generation == created.input_generation
    assert changed_input.input_generation == created.input_generation + 1


def test_save_and_get_round_trip(tmp_path: Path) -> None:
    store = ConfigSessionStore(tmp_path / "sessions.sqlite3")
    key = session_key_for({"current_excel": "a" * 64})
    assert store.get(key) is None

    saved = store.save_choices(
        key, profile_name="monthly", choices={"member_map": {"primary": "ops"}}
    )
    assert saved.session_key == key
    assert saved.profile_name == "monthly"
    assert saved.choices == {"member_map": {"primary": "ops"}}

    fetched = store.get(key)
    assert fetched is not None
    assert fetched.choices == {"member_map": {"primary": "ops"}}


def test_save_choices_replaces_wholesale_and_updates_timestamp(tmp_path: Path) -> None:
    store = ConfigSessionStore(tmp_path / "sessions.sqlite3")
    key = session_key_for({"current_excel": "a" * 64})
    first = store.save_choices(key, choices={"step": 1})
    second = store.save_choices(key, choices={"step": 2})

    assert second.choices == {"step": 2}
    assert second.created_at == first.created_at
    assert second.updated_at >= first.updated_at


def test_delete_removes_the_session(tmp_path: Path) -> None:
    store = ConfigSessionStore(tmp_path / "sessions.sqlite3")
    key = session_key_for({"current_excel": "a" * 64})
    store.save_choices(key, choices={"step": 1})
    store.delete(key)
    assert store.get(key) is None


def test_delete_stale_removes_only_old_sessions(tmp_path: Path) -> None:
    store = ConfigSessionStore(tmp_path / "sessions.sqlite3")
    old_key = session_key_for({"current_excel": "a" * 64})
    new_key = session_key_for({"current_excel": "b" * 64})
    store.save_choices(old_key, choices={"step": 1})
    store.save_choices(new_key, choices={"step": 1})

    future_cutoff = dt.datetime.now(dt.UTC) + dt.timedelta(days=1)
    removed = store.delete_stale(future_cutoff)

    assert removed == 2
    assert store.get(old_key) is None
    assert store.get(new_key) is None


def test_delete_stale_keeps_recent_sessions(tmp_path: Path) -> None:
    store = ConfigSessionStore(tmp_path / "sessions.sqlite3")
    key = session_key_for({"current_excel": "a" * 64})
    store.save_choices(key, choices={"step": 1})

    past_cutoff = dt.datetime.now(dt.UTC) - dt.timedelta(days=1)
    removed = store.delete_stale(past_cutoff)

    assert removed == 0
    assert store.get(key) is not None
