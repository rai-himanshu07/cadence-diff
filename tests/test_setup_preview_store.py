"""Tests for the Step 6 setup-scan sidecar store
(``qc_tool.setup.preview_store``).

Plan: docs/plans/plan-20260913-mode-aware-configuration-wizard.md, Step 6.
"""

from __future__ import annotations

import datetime as dt
import sqlite3
from pathlib import Path

import pytest

from qc_tool.io.model import CellRecord, SheetSnapshot
from qc_tool.setup.preview_store import (
    SetupScanStore,
    SetupSidecarCorruptError,
    SetupSidecarStaleError,
)

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


def test_delete_stale_preserves_a_newer_generation_in_the_same_session(
    tmp_path: Path,
) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    store.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=1,
        source_hash="a" * 64,
        sheet=_typed_sheet(),
    )
    old_time = (dt.datetime.now(dt.UTC) - dt.timedelta(days=30)).isoformat(
        timespec="milliseconds"
    )
    with store._connect() as connection:
        connection.execute(
            """
            UPDATE setup_sheet_inventory SET created_at = ?
            WHERE session_key = ? AND input_generation = 1
            """,
            (old_time, "session-1"),
        )
    new_sheet = SheetSnapshot(
        name="Data",
        visibility="visible",
        max_row=1,
        max_column=1,
        cells={(1, 1): CellRecord(1, 1, "new")},
    )
    store.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=2,
        source_hash="b" * 64,
        sheet=new_sheet,
    )

    removed = store.delete_stale(dt.datetime.now(dt.UTC) - dt.timedelta(days=7))

    assert removed == 1
    restored = store.load_sheet(
        "session-1",
        "primary",
        "current",
        "Data",
        expected_generation=2,
        expected_source_hash="b" * 64,
    )
    assert restored.cells[(1, 1)].value == "new"


def test_different_members_of_the_same_session_are_independent(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    store.save_result("session-1", "primary", _RESULT)
    store.save_result("session-1", "secondary", dict(_RESULT, member_id="secondary"))

    assert store.get_result("session-1", "primary") is not None
    secondary = store.get_result("session-1", "secondary")
    assert secondary is not None
    assert secondary["member_id"] == "secondary"


def _typed_sheet() -> SheetSnapshot:
    return SheetSnapshot(
        name="Data",
        visibility="hidden",
        max_row=5,
        max_column=4,
        hidden_rows=frozenset({4}),
        hidden_columns=frozenset({3}),
        cells={
            (1, 1): CellRecord(1, 1, "Header"),
            (2, 1): CellRecord(2, 1, 7),
            (2, 2): CellRecord(2, 2, 2.5),
            (2, 3): CellRecord(2, 3, True),
            (3, 1): CellRecord(3, 1, dt.date(2026, 9, 15)),
            (3, 2): CellRecord(3, 2, dt.datetime(2026, 9, 15, 10, 30)),
            (5, 4): CellRecord(
                5,
                4,
                None,
                formula="=SUM(A1:A2)",
                is_formula=True,
            ),
        },
    )


def test_sheet_blocks_round_trip_typed_cells_and_metadata(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3", block_rows=2)
    store.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=3,
        source_hash="a" * 64,
        sheet=_typed_sheet(),
    )

    restored = store.load_sheet(
        "session-1",
        "primary",
        "current",
        "Data",
        expected_generation=3,
        expected_source_hash="a" * 64,
    )

    assert restored.visibility == "hidden"
    assert restored.hidden_rows == frozenset({4})
    assert restored.hidden_columns == frozenset({3})
    assert restored.cells[(2, 1)].value == 7
    assert isinstance(restored.cells[(2, 1)].value, int)
    assert restored.cells[(2, 2)].value == 2.5
    assert isinstance(restored.cells[(2, 2)].value, float)
    assert restored.cells[(2, 3)].value is True
    assert restored.cells[(3, 1)].value == dt.date(2026, 9, 15)
    assert not isinstance(restored.cells[(3, 1)].value, dt.datetime)
    assert restored.cells[(3, 2)].value == dt.datetime(2026, 9, 15, 10, 30)
    assert restored.cells[(5, 4)].has_formula
    assert restored.cells[(5, 4)].formula == "=SUM(A1:A2)"


def test_window_query_decodes_only_overlapping_row_blocks(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3", block_rows=2)
    store.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=1,
        source_hash="a" * 64,
        sheet=_typed_sheet(),
    )

    window = store.query_window(
        "session-1",
        "primary",
        "current",
        "Data",
        expected_generation=1,
        expected_source_hash="a" * 64,
        min_row=2,
        max_row=3,
        min_col=1,
        max_col=2,
    )

    assert set(window.cells) == {(2, 1), (2, 2), (3, 1), (3, 2)}
    assert window.decoded_blocks == 2
    assert window.total_blocks == 3


def test_column_query_is_bounded_to_requested_rows_and_columns(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3", block_rows=2)
    store.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=1,
        source_hash="a" * 64,
        sheet=_typed_sheet(),
    )

    cells = store.query_cells(
        "session-1",
        "primary",
        "current",
        "Data",
        expected_generation=1,
        expected_source_hash="a" * 64,
        min_row=2,
        max_row=3,
        columns=frozenset({1}),
    )

    assert set(cells) == {(2, 1), (3, 1)}


@pytest.mark.parametrize(
    ("generation", "source_hash"),
    [(2, "a" * 64), (1, "b" * 64)],
)
def test_sheet_queries_reject_stale_generation_or_source_hash(
    tmp_path: Path, generation: int, source_hash: str
) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    store.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=1,
        source_hash="a" * 64,
        sheet=_typed_sheet(),
    )

    with pytest.raises(SetupSidecarStaleError):
        store.load_sheet(
            "session-1",
            "primary",
            "current",
            "Data",
            expected_generation=generation,
            expected_source_hash=source_hash,
        )


def test_late_old_generation_cannot_overwrite_new_generation(
    tmp_path: Path,
) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    old_sheet = _typed_sheet()
    new_sheet = SheetSnapshot(
        name="Data",
        visibility="visible",
        max_row=1,
        max_column=1,
        cells={(1, 1): CellRecord(1, 1, "new-generation")},
    )

    store.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=2,
        source_hash="b" * 64,
        sheet=new_sheet,
    )
    store.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=1,
        source_hash="a" * 64,
        sheet=old_sheet,
    )

    current = store.load_sheet(
        "session-1",
        "primary",
        "current",
        "Data",
        expected_generation=2,
        expected_source_hash="b" * 64,
    )
    old = store.load_sheet(
        "session-1",
        "primary",
        "current",
        "Data",
        expected_generation=1,
        expected_source_hash="a" * 64,
    )

    assert current.cells[(1, 1)].value == "new-generation"
    assert old.cells[(1, 1)].value == "Header"


def test_delete_generation_reclaims_only_the_superseded_generation(
    tmp_path: Path,
) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    old_sheet = _typed_sheet()
    new_sheet = SheetSnapshot(
        name="Data",
        visibility="visible",
        max_row=1,
        max_column=1,
        cells={(1, 1): CellRecord(1, 1, "new")},
    )
    store.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=1,
        source_hash="a" * 64,
        sheet=old_sheet,
    )
    store.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=2,
        source_hash="b" * 64,
        sheet=new_sheet,
    )

    assert store.delete_generation("session-1", 1) == 1

    with pytest.raises(SetupSidecarStaleError):
        store.load_sheet(
            "session-1",
            "primary",
            "current",
            "Data",
            expected_generation=1,
            expected_source_hash="a" * 64,
        )
    current = store.load_sheet(
        "session-1",
        "primary",
        "current",
        "Data",
        expected_generation=2,
        expected_source_hash="b" * 64,
    )
    assert current.cells[(1, 1)].value == "new"


def test_malformed_sheet_block_fails_closed_without_exposing_content(
    tmp_path: Path,
) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    store.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=1,
        source_hash="a" * 64,
        sheet=_typed_sheet(),
    )
    with sqlite3.connect(tmp_path / "setup_scans.sqlite3") as connection:
        connection.execute(
            "UPDATE setup_cell_blocks SET blob = ? WHERE session_key = ?",
            (b"not-a-zlib-stream", "session-1"),
        )

    with pytest.raises(SetupSidecarCorruptError, match="setup sidecar block is unreadable"):
        store.load_sheet(
            "session-1",
            "primary",
            "current",
            "Data",
            expected_generation=1,
            expected_source_hash="a" * 64,
        )


def test_delete_removes_sheet_inventory_and_cell_blocks(tmp_path: Path) -> None:
    store = SetupScanStore(tmp_path / "setup_scans.sqlite3")
    store.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=1,
        source_hash="a" * 64,
        sheet=_typed_sheet(),
    )

    store.delete("session-1")

    assert store.list_sheets("session-1", "primary", "current") == ()
