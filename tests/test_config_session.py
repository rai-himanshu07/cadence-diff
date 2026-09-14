"""Tests for the Step 5 configuration-session store
(``qc_tool.history.config_session``).

Plan: docs/plans/plan-20260913-mode-aware-configuration-wizard.md, Step 5.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

from qc_tool.history.config_session import ConfigSessionStore, session_key_for


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
