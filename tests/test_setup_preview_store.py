"""Tests for the Step 6 setup-scan sidecar store
(``qc_tool.setup.preview_store``).

Plan: docs/plans/plan-20260913-mode-aware-configuration-wizard.md, Step 6.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

from qc_tool.setup.preview_store import SetupScanStore

_RESULT: dict[str, object] = {
    "member_id": "primary",
    "baseline_hash": "a" * 64,
    "current_hash": "b" * 64,
    "baseline_sheets": [{"sheet_name": "Data", "hidden": False, "regions": []}],
    "current_sheets": [
        {"sheet_name": "Data", "hidden": False, "regions": []},
        {"sheet_name": "Summary", "hidden": True, "regions": []},
    ],
}


def test_save_and_get_result_round_trip(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    assert store.get_result("session-1", "primary") is None

    saved = store.save_result("session-1", "primary", _RESULT)
    assert saved.session_key == "session-1"
    assert saved.member_id == "primary"
    assert saved.blob_bytes > 0

    fetched = store.get_result("session-1", "primary")
    assert fetched == _RESULT


def test_result_is_compressed_at_rest(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    store.save_result("session-1", "primary", _RESULT)
    with store._connect() as conn:
        row = conn.execute(
            "SELECT blob FROM setup_scan_blocks WHERE session_key = ?", ("session-1",)
        ).fetchone()
    # A real zlib stream never decodes as plain UTF-8 JSON text.
    assert not bytes(row["blob"]).startswith(b"{")


def test_get_sheet_profile_returns_one_sheet_only(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    store.save_result("session-1", "primary", _RESULT)

    profile = store.get_sheet_profile("session-1", "primary", "Summary")
    assert profile == {"sheet_name": "Summary", "hidden": True, "regions": []}

    baseline_profile = store.get_sheet_profile(
        "session-1", "primary", "Data", side="baseline"
    )
    assert baseline_profile == {"sheet_name": "Data", "hidden": False, "regions": []}


def test_get_sheet_profile_returns_none_for_an_unknown_sheet(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    store.save_result("session-1", "primary", _RESULT)

    assert store.get_sheet_profile("session-1", "primary", "NoSuchSheet") is None


def test_save_result_replaces_the_prior_result_for_the_same_key(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    store.save_result("session-1", "primary", _RESULT)
    updated = dict(_RESULT, current_hash="c" * 64)
    store.save_result("session-1", "primary", updated)

    fetched = store.get_result("session-1", "primary")
    assert fetched is not None
    assert fetched["current_hash"] == "c" * 64


def test_delete_removes_the_result(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    store.save_result("session-1", "primary", _RESULT)
    store.delete("session-1")
    assert store.get_result("session-1", "primary") is None


def test_delete_stale_removes_only_old_results(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    store.save_result("session-1", "primary", _RESULT)
    store.save_result("session-2", "primary", _RESULT)

    future_cutoff = dt.datetime.now(dt.UTC) + dt.timedelta(days=1)
    removed = store.delete_stale(future_cutoff)

    assert removed == 2
    assert store.get_result("session-1", "primary") is None
    assert store.get_result("session-2", "primary") is None


def test_delete_stale_keeps_recent_results(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    store.save_result("session-1", "primary", _RESULT)

    past_cutoff = dt.datetime.now(dt.UTC) - dt.timedelta(days=1)
    removed = store.delete_stale(past_cutoff)

    assert removed == 0
    assert store.get_result("session-1", "primary") is not None


def test_different_members_of_the_same_session_are_independent(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    store.save_result("session-1", "primary", _RESULT)
    store.save_result("session-1", "secondary", dict(_RESULT, member_id="secondary"))

    assert store.get_result("session-1", "primary") is not None
    secondary = store.get_result("session-1", "secondary")
    assert secondary is not None
    assert secondary["member_id"] == "secondary"
