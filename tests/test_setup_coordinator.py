"""Lifecycle contracts for the progressive setup coordinator."""

from __future__ import annotations

import threading
from pathlib import Path

import pytest

from qc_tool.io.model import CellRecord, SheetSnapshot
from qc_tool.setup.coordinator import (
    SetupCoordinator,
    SetupCoordinatorStore,
    SetupMemberInput,
    SetupStatus,
)
from qc_tool.setup.preview_store import SetupScanStore, SetupSidecarStaleError
from qc_tool.setup.preview_worker import SetupScanOutcome, SetupScanProgress


def _member(tmp_path: Path) -> SetupMemberInput:
    return SetupMemberInput(
        member_id="primary",
        baseline_path=str(tmp_path / "baseline.xlsx"),
        current_path=str(tmp_path / "current.xlsx"),
        baseline_hash="a" * 64,
        current_hash="b" * 64,
    )


def _result_payload() -> dict[str, object]:
    return {
        "member_id": "primary",
        "baseline_hash": "a" * 64,
        "current_hash": "b" * 64,
        "baseline_sheets": [],
        "current_sheets": [],
    }


def test_coordinator_persists_aggregate_progress_and_completion(tmp_path: Path) -> None:
    def runner(request, *, cancel_event, on_progress):
        del request, cancel_event
        on_progress(SetupScanProgress("inventory_ready", "current", 0, 2))
        on_progress(SetupScanProgress("sheet_ready", "current", 1, 2))
        on_progress(SetupScanProgress("sheet_ready", "current", 2, 2))
        return SetupScanOutcome(result_payload=_result_payload())

    coordinator = SetupCoordinator(tmp_path, runner=runner)
    coordinator.attach("session-1", 1, "client-1")
    job = coordinator.start(
        "session-1",
        1,
        {"primary": _member(tmp_path)},
        passwords_by_member={},
    )

    assert job.wait_terminal(2.0)
    snapshot = job.snapshot()
    assert snapshot.status is SetupStatus.COMPLETE
    assert snapshot.processed == 2
    assert snapshot.total == 2
    assert snapshot.member_status == {"primary": "done"}
    persisted = SetupCoordinatorStore(tmp_path / "history.sqlite3").get(
        "session-1", 1
    )
    assert persisted is not None
    assert persisted.status is SetupStatus.COMPLETE
    assert persisted.processed == 2
    assert persisted.total == 2


def test_completed_sheet_profile_is_available_before_member_completion(
    tmp_path: Path,
) -> None:
    emitted = threading.Event()
    release = threading.Event()

    def runner(request, *, cancel_event, on_progress):
        del request, cancel_event
        on_progress(SetupScanProgress("inventory_ready", "current", 0, 2))
        on_progress(
            SetupScanProgress(
                "sheet_ready",
                "current",
                1,
                2,
                profile_payload={
                    "sheet_name": "Synthetic",
                    "hidden": False,
                    "very_hidden": False,
                    "regions": [],
                    "failure_detail": "",
                },
            )
        )
        emitted.set()
        release.wait(2.0)
        return SetupScanOutcome(result_payload=_result_payload())

    coordinator = SetupCoordinator(tmp_path, runner=runner)
    coordinator.attach("session-1", 1, "client-1")
    job = coordinator.start(
        "session-1",
        1,
        {"primary": _member(tmp_path)},
        passwords_by_member={},
    )

    assert emitted.wait(1.0)
    partial = job.result_payloads_snapshot()["primary"]
    assert job.snapshot().status is SetupStatus.PARTIAL_READY
    assert partial["current_sheets"] == [
        {
            "sheet_name": "Synthetic",
            "hidden": False,
            "very_hidden": False,
            "regions": [],
            "failure_detail": "",
        }
    ]

    release.set()
    assert job.wait_terminal(2.0)


def test_credential_bound_slot_wait_demotes_when_last_page_detaches(
    tmp_path: Path,
) -> None:
    attempted = threading.Event()

    def busy_exclusive(_work_dir, _holder, _body):
        attempted.set()
        from qc_tool.runqueue import QueueBusyError

        raise QueueBusyError("busy")

    coordinator = SetupCoordinator(tmp_path, exclusive_runner=busy_exclusive)
    coordinator.attach("session-1", 1, "client-1")
    job = coordinator.start(
        "session-1",
        1,
        {"primary": _member(tmp_path)},
        passwords_by_member={"primary": {"current": "secret"}},
    )
    assert attempted.wait(1.0)

    coordinator.detach("session-1", 1, "client-1")

    assert job.wait_for_status(SetupStatus.AWAITING_CREDENTIALS, 2.0)
    assert "secret" not in repr(job)


def test_credential_bound_inflight_job_restores_as_awaiting_credentials(
    tmp_path: Path,
) -> None:
    entered = threading.Event()
    release = threading.Event()

    def runner(request, *, cancel_event, on_progress):
        del request, cancel_event, on_progress
        entered.set()
        release.wait(2.0)
        return SetupScanOutcome(result_payload=_result_payload())

    coordinator = SetupCoordinator(tmp_path, runner=runner)
    coordinator.attach("session-1", 1, "client-1")
    coordinator.start(
        "session-1",
        1,
        {"primary": _member(tmp_path)},
        passwords_by_member={"primary": {"current": "secret"}},
    )
    assert entered.wait(1.0)

    reconstructed = SetupCoordinator(tmp_path)
    restored = reconstructed.get_or_create("session-1", 1)

    assert restored.snapshot().status is SetupStatus.AWAITING_CREDENTIALS
    release.set()


def test_detach_demotes_later_credential_member_after_an_earlier_dispatch(
    tmp_path: Path,
) -> None:
    second_waiting = threading.Event()
    calls = 0

    def exclusive(_work_dir, _holder, body):
        nonlocal calls
        calls += 1
        if calls == 1:
            return body()
        second_waiting.set()
        from qc_tool.runqueue import QueueBusyError

        raise QueueBusyError("busy")

    def runner(request, *, cancel_event, on_progress):
        del cancel_event, on_progress
        return SetupScanOutcome(
            result_payload={
                **_result_payload(),
                "member_id": request.member_id,
            }
        )

    coordinator = SetupCoordinator(
        tmp_path,
        runner=runner,
        exclusive_runner=exclusive,
    )
    coordinator.attach("session-1", 1, "client-1")
    job = coordinator.start(
        "session-1",
        1,
        {
            "first": _member(tmp_path),
            "second": SetupMemberInput(
                member_id="second",
                baseline_path=str(tmp_path / "baseline-2.xlsx"),
                current_path=str(tmp_path / "current-2.xlsx"),
                baseline_hash="c" * 64,
                current_hash="d" * 64,
            ),
        },
        passwords_by_member={
            "first": {"current": "secret-one"},
            "second": {"current": "secret-two"},
        },
    )
    assert second_waiting.wait(1.0)

    coordinator.detach("session-1", 1, "client-1")

    assert job.wait_for_status(SetupStatus.AWAITING_CREDENTIALS, 2.0)
    assert "secret-two" not in repr(job)
    assert calls == 2


def test_cancel_wins_over_a_late_successful_result(tmp_path: Path) -> None:
    entered = threading.Event()
    release = threading.Event()

    def runner(request, *, cancel_event, on_progress):
        del request, cancel_event, on_progress
        entered.set()
        release.wait(2.0)
        return SetupScanOutcome(result_payload=_result_payload())

    coordinator = SetupCoordinator(tmp_path, runner=runner)
    coordinator.attach("session-1", 1, "client-1")
    job = coordinator.start(
        "session-1",
        1,
        {"primary": _member(tmp_path)},
        passwords_by_member={},
    )
    assert entered.wait(1.0)

    coordinator.cancel("session-1", 1)
    release.set()

    assert job.wait_terminal(2.0)
    assert job.snapshot().status is SetupStatus.CANCELLED
    assert job.result_payloads == {}


def test_restart_returns_a_fresh_job_handle_for_the_same_generation(
    tmp_path: Path,
) -> None:
    coordinator = SetupCoordinator(tmp_path)
    first = coordinator.get_or_create("session-1", 1)
    coordinator.cancel("session-1", 1)

    replacement = coordinator.restart("session-1", 1)

    assert replacement is not first
    assert replacement.snapshot().status is SetupStatus.AWAITING_INPUTS


def test_terminal_job_evicts_after_last_subscriber_detaches(tmp_path: Path) -> None:
    def runner(request, *, cancel_event, on_progress):
        del request, cancel_event, on_progress
        return SetupScanOutcome(result_payload=_result_payload())

    coordinator = SetupCoordinator(tmp_path, runner=runner)
    coordinator.attach("session-1", 1, "client-1")
    job = coordinator.start(
        "session-1",
        1,
        {"primary": _member(tmp_path)},
        passwords_by_member={},
    )
    assert job.wait_terminal(2.0)

    coordinator.detach("session-1", 1, "client-1")
    assert coordinator.evict_terminal(max_age_seconds=0) == 1
    assert coordinator.find("session-1", 1) is None


def test_discard_removes_in_memory_and_persisted_job_state(tmp_path: Path) -> None:
    coordinator = SetupCoordinator(tmp_path)
    coordinator.get_or_create("session-1", 1)
    sidecar = SetupScanStore(tmp_path / "setup-inspection.sqlite3")
    sidecar.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=1,
        source_hash="a" * 64,
        sheet=SheetSnapshot(
            name="Synthetic",
            visibility="visible",
            max_row=1,
            max_column=1,
            cells={(1, 1): CellRecord(1, 1, "value")},
        ),
    )

    coordinator.discard("session-1", 1)

    assert coordinator.find("session-1", 1) is None
    assert SetupCoordinatorStore(tmp_path / "history.sqlite3").get(
        "session-1", 1
    ) is None
    assert sidecar.list_sheets("session-1", "primary", "current") == ()


def test_discard_reclaims_late_old_write_and_preserves_new_generation(
    tmp_path: Path,
) -> None:
    entered = threading.Event()
    release = threading.Event()
    sidecar = SetupScanStore(tmp_path / "setup-inspection.sqlite3")

    def sheet(value: str) -> SheetSnapshot:
        return SheetSnapshot(
            name="Synthetic",
            visibility="visible",
            max_row=1,
            max_column=1,
            cells={(1, 1): CellRecord(1, 1, value)},
        )

    def runner(request, *, cancel_event, on_progress):
        del cancel_event, on_progress
        entered.set()
        release.wait(2.0)
        sidecar.save_sheet(
            request.session_key,
            request.member_id,
            "current",
            input_generation=request.input_generation,
            source_hash=request.current_hash,
            sheet=sheet("late-old"),
        )
        return SetupScanOutcome(result_payload=_result_payload())

    coordinator = SetupCoordinator(tmp_path, runner=runner)
    coordinator.attach("session-1", 1, "client-1")
    coordinator.start(
        "session-1",
        1,
        {"primary": _member(tmp_path)},
        passwords_by_member={},
    )
    assert entered.wait(1.0)
    sidecar.save_sheet(
        "session-1",
        "primary",
        "current",
        input_generation=2,
        source_hash="d" * 64,
        sheet=sheet("new"),
    )

    cleanup_thread = coordinator.discard("session-1", 1)
    assert cleanup_thread is not None
    release.set()
    cleanup_thread.join(timeout=2.0)
    assert not cleanup_thread.is_alive()

    with pytest.raises(SetupSidecarStaleError):
        sidecar.load_sheet(
            "session-1",
            "primary",
            "current",
            "Synthetic",
            expected_generation=1,
            expected_source_hash="b" * 64,
        )
    current = sidecar.load_sheet(
        "session-1",
        "primary",
        "current",
        "Synthetic",
        expected_generation=2,
        expected_source_hash="d" * 64,
    )
    assert current.cells[(1, 1)].value == "new"
