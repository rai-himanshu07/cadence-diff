"""Step 11: persistent, process-global, single-worker QC run queue."""

from __future__ import annotations

import ast
import contextlib
import json
import multiprocessing as mp
import os
import queue as queue_module
import threading
import time
from collections.abc import Callable, Iterator
from dataclasses import asdict
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast

import pytest

if TYPE_CHECKING:
    from multiprocessing.context import SpawnContext

import qc_tool
from qc_tool.config.profile import DeliverableProfile
from qc_tool.coverage import QCRunMode
from qc_tool.history.run_state import RunStateStore, RunStatus
from qc_tool.history.store import RunHistory
from qc_tool.package import PackageManifest
from qc_tool.runqueue import (
    QueueBusyError,
    RunQueueManager,
    RunRequest,
    new_request_id,
)
from qc_tool.worker import (
    ENVELOPE_VERSION,
    IPC_QUEUE_CAPACITY,
    OwnedCancellationFlag,
    _deliver,
    cancelled_message,
    error_message,
    progress_message,
    result_message,
    sanitize_error,
    worker_main,
)
from tests.conftest import fixture_profile

_FAKE_RUN_ID = 4242
_ORACLE = Path(__file__).parent / "oracles" / "step1_scenarios.json"


def _scenario(identifier: str) -> dict[str, Any]:
    manifest = json.loads(_ORACLE.read_text(encoding="utf-8"))
    for scenario in manifest["scenarios"]:
        if scenario["id"] == identifier:
            fixture = scenario["fixture"]
            assert isinstance(fixture, dict)
            return fixture
    raise AssertionError(f"missing frozen scenario {identifier}")


# --- module-level workers (spawn pickles them by qualified name) -------------


def _sequenced_worker(payload, credentials, events, cancel_flag) -> None:
    log = Path(payload["work_dir"]) / "execution.log"
    with log.open("a", encoding="utf-8") as handle:
        handle.write(f"start:{payload['request_id']}\n")
    time.sleep(0.2)
    with log.open("a", encoding="utf-8") as handle:
        handle.write(f"end:{payload['request_id']}\n")
    events.put(result_message(_FAKE_RUN_ID, []))


def _cooperative_worker(payload, credentials, events, cancel_flag) -> None:
    events.put(progress_message("analyzing_excel", 0, 0, ""))
    for _ in range(600):
        if cancel_flag.is_set():
            events.put(cancelled_message([]))
            return
        time.sleep(0.05)
    events.put(result_message(_FAKE_RUN_ID, []))


def _stubborn_worker(payload, credentials, events, cancel_flag) -> None:
    import signal

    with contextlib.suppress(ValueError, OSError, AttributeError):
        signal.signal(signal.SIGTERM, signal.SIG_IGN)
    events.put(progress_message("analyzing_excel", 0, 0, ""))
    while True:
        time.sleep(0.05)


def _crash_worker(payload, credentials, events, cancel_flag) -> None:
    events.put(progress_message("analyzing_excel", 0, 0, ""))
    time.sleep(0.2)
    os._exit(3)


def _silent_worker(payload, credentials, events, cancel_flag) -> None:
    events.put(progress_message("analyzing_excel", 0, 0, ""))
    time.sleep(0.2)


def _failing_worker(payload, credentials, events, cancel_flag) -> None:
    events.put(
        error_message(
            sanitize_error(RuntimeError("cannot read /private/uploads/quarter.xlsx")),
            [],
        )
    )


def _credential_probe_worker(payload, credentials, events, cancel_flag) -> None:
    marker = Path(payload["work_dir"]) / "credential-roles.txt"
    marker.write_text(",".join(sorted(credentials)), encoding="utf-8")
    events.put(result_message(_FAKE_RUN_ID, []))


def _flood_worker(payload, credentials, events, cancel_flag) -> None:
    total = IPC_QUEUE_CAPACITY * 4
    for index in range(total):
        with contextlib.suppress(queue_module.Full):
            events.put_nowait(progress_message("comparing_formulas", index, total, ""))
    events.put(result_message(_FAKE_RUN_ID, []))


def _telemetry_worker(payload, credentials, events, cancel_flag) -> None:
    events.put(progress_message("comparing_formulas", 3, 7, "block 3 of 7"))
    time.sleep(0.15)
    events.put(
        result_message(
            _FAKE_RUN_ID,
            [
                {
                    "phase": "comparing_formulas",
                    "elapsed_seconds": 0.1,
                    "processed": 7,
                    "total": 7,
                    "peak_rss_bytes": 1024,
                    "error": "",
                }
            ],
        )
    )


# --- helpers -----------------------------------------------------------------


ManagerFactory = Callable[..., RunQueueManager]


@pytest.fixture
def make_manager(tmp_path: Path) -> Iterator[ManagerFactory]:
    managers: list[RunQueueManager] = []

    def factory(worker: Callable[..., None], **kwargs: Any) -> RunQueueManager:
        work_dir = tmp_path / f"work{len(managers)}"
        work_dir.mkdir()
        manager = RunQueueManager(work_dir, worker=worker, **kwargs)
        managers.append(manager)
        return manager

    yield factory
    for manager in managers:
        manager.shutdown()


def _request(manager: RunQueueManager, **overrides: Any) -> RunRequest:
    fields: dict[str, Any] = {
        "request_id": new_request_id(),
        "work_dir": str(manager.work_dir),
        "mode": QCRunMode.CYCLE_COMPARISON.value,
        "profile_name": "fixture",
        "profile": {},
        "files": {},
        "display_files": {"current_excel": "current.xlsx"},
    }
    fields.update(overrides)
    return RunRequest(**fields)


def _wait_until(condition: Callable[[], bool], timeout: float = 30.0) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if condition():
            return
        time.sleep(0.02)
    raise AssertionError("queue condition was not reached in time")


class _SlowStartContext:
    """Spawn context that widens the window between dequeue and a live child."""

    def __init__(self, context: Any, *, delay: float) -> None:
        self._context = context
        self._delay = delay
        self.started = False

    # Capitalized names mirror the multiprocessing context API.
    def Queue(self, maxsize: int = 0) -> Any:
        return self._context.Queue(maxsize)

    def Event(self) -> Any:
        return self._context.Event()

    def Process(self, **kwargs: Any) -> Any:
        process = self._context.Process(**kwargs)
        outer = self

        class _DelayedStart:
            def __getattr__(self, name: str) -> Any:
                return getattr(process, name)

            def start(self) -> None:
                outer.started = True
                time.sleep(outer._delay)
                process.start()

        return _DelayedStart()


def _has_status(
    manager: RunQueueManager, request_id: str, status: RunStatus
) -> Callable[[], bool]:
    def condition() -> bool:
        record = manager.store.get(request_id)
        return record is not None and record.status is status

    return condition


# --- IPC contract ------------------------------------------------------------


def test_ipc_envelope_round_trips_versioned_primitives_only() -> None:
    messages = [
        progress_message("comparing_formulas", 2, 9, "block 2 of 9"),
        result_message(7, [{"phase": "comparing_formulas", "peak_rss_bytes": 1}]),
        cancelled_message([]),
        error_message("RuntimeError: workbook is unreadable", []),
    ]

    for message in messages:
        assert message["v"] == ENVELOPE_VERSION
        assert json.loads(json.dumps(message)) == message
    assert {message["kind"] for message in messages} == {
        "progress",
        "result",
        "cancelled",
        "error",
    }


def test_sanitize_error_keeps_one_bounded_line_without_paths() -> None:
    posix = sanitize_error(RuntimeError("cannot open /srv/private/uploads/q1.xlsx"))
    windows = sanitize_error(r"failed on C:\Users\analyst\uploads\q1.xlsx")
    multiline = sanitize_error("first line\nsecond line")

    assert posix == "RuntimeError: cannot open q1.xlsx"
    assert windows == "failed on q1.xlsx"
    assert multiline == "first line"
    assert len(sanitize_error("x" * 500)) <= 240


def test_terminal_delivery_never_raises_on_a_broken_channel() -> None:
    class _BrokenQueue:
        def put(self, message: object, timeout: float | None = None) -> None:
            raise queue_module.Full

    _deliver(_BrokenQueue(), result_message(1, []))  # pyright: ignore[reportArgumentType]


def test_owned_worker_cancels_itself_when_its_owner_disappears() -> None:
    class _Parent:
        def __init__(self) -> None:
            self.alive = True

        def is_alive(self) -> bool:
            return self.alive

    parent = _Parent()
    flag = OwnedCancellationFlag(threading.Event(), parent, interval=0.0)

    assert flag.is_set() is False
    assert flag.orphaned is False

    parent.alive = False

    assert flag.is_set() is True
    assert flag.orphaned is True
    # Without an owner reference the flag only follows the manager's request.
    detached = OwnedCancellationFlag(threading.Event(), None, interval=0.0)
    assert detached.is_set() is False
    detached.set()
    assert detached.is_set() is True


@pytest.mark.parametrize(
    "entry_point",
    [Path(__file__).parents[1] / "main.py", Path(qc_tool.__file__).parent / "__main__.py"],
)
def test_entry_points_never_serve_when_re_imported_by_a_spawn_worker(
    entry_point: Path,
) -> None:
    """A spawn child re-imports the parent main module as `__mp_main__`."""
    guards = [
        node.test
        for node in ast.parse(entry_point.read_text(encoding="utf-8")).body
        if isinstance(node, ast.If)
    ]

    assert guards, f"{entry_point.name} has no __name__ guard"
    for guard in guards:
        names = {
            literal.value
            for literal in ast.walk(guard)
            if isinstance(literal, ast.Constant) and isinstance(literal.value, str)
        }
        assert names == {"__main__"}, f"{entry_point.name} also runs as {names}"


# --- FIFO, reconnection, cancellation ----------------------------------------


def test_one_worker_runs_queued_requests_in_submission_order(
    make_manager: ManagerFactory,
) -> None:
    scenario = _scenario("queue-refresh")
    manager = make_manager(_sequenced_worker)
    requests = [_request(manager) for _ in scenario["expected_execution_order"]]

    for request in requests:
        manager.submit(request)
    for request in requests:
        assert manager.wait(request.request_id).status is RunStatus.SUCCEEDED

    log = (manager.work_dir / "execution.log").read_text(encoding="utf-8").split()
    assert log == [
        entry
        for request in requests
        for entry in (f"start:{request.request_id}", f"end:{request.request_id}")
    ]


def test_a_second_tab_reconnects_to_active_and_queued_work_from_sqlite(
    make_manager: ManagerFactory,
) -> None:
    manager = make_manager(_cooperative_worker)
    requests = [_request(manager) for _ in range(3)]
    for request in requests:
        manager.submit(request)

    # A refreshed page opens its own store; nothing is resubmitted.
    reconnected = RunStateStore(manager.work_dir / "history.sqlite3")
    _wait_until(
        lambda: any(
            record.status is RunStatus.RUNNING for record in reconnected.pending()
        )
    )
    pending = reconnected.pending()

    assert [record.request_id for record in pending] == [
        request.request_id for request in requests
    ]
    assert [record.queue_position for record in pending] == [0, 1, 2]
    assert pending[0].status is RunStatus.RUNNING
    assert {record.status for record in pending[1:]} == {RunStatus.QUEUED}
    assert pending[0].files == {"current_excel": "current.xlsx"}
    manager.cancel(requests[0].request_id)


def test_queued_cancellation_prevents_the_request_from_ever_starting(
    make_manager: ManagerFactory,
) -> None:
    manager = make_manager(_sequenced_worker)
    active = _request(manager)
    queued = _request(manager)
    manager.submit(active)
    manager.submit(queued)

    assert manager.cancel(queued.request_id) is True
    assert manager.wait(active.request_id).status is RunStatus.SUCCEEDED
    cancelled = manager.store.get(queued.request_id)

    assert cancelled is not None
    assert cancelled.status is RunStatus.CANCELLED
    assert cancelled.started_at is None
    log = (manager.work_dir / "execution.log").read_text(encoding="utf-8")
    assert queued.request_id not in log


def test_active_cancellation_stops_a_cooperative_worker(
    make_manager: ManagerFactory,
) -> None:
    manager = make_manager(_cooperative_worker)
    request = _request(manager)
    manager.submit(request)
    _wait_until(_has_status(manager, request.request_id, RunStatus.RUNNING))

    assert manager.cancel(request.request_id) is True
    record = manager.wait(request.request_id)

    assert record.status is RunStatus.CANCELLED
    assert record.cancel_requested is True
    assert record.run_id is None


def test_cancelling_during_the_spawn_window_still_stops_the_request(
    tmp_path: Path,
) -> None:
    """The slot is reserved before the child exists; a cancel there must hold."""
    work_dir = tmp_path / "spawn-window"
    work_dir.mkdir()
    context = _SlowStartContext(mp.get_context("spawn"), delay=1.0)
    manager = RunQueueManager(
        work_dir,
        context=cast("SpawnContext", context),
        worker=_cooperative_worker,
        cancel_grace_seconds=0.2,
        terminate_join_seconds=0.5,
    )
    try:
        request = _request(manager)
        manager.submit(request)
        _wait_until(lambda: context.started)
        assert manager.cancel(request.request_id) is True

        record = manager.wait(request.request_id, timeout=60.0)

        assert record.status is RunStatus.CANCELLED
        assert record.cancel_requested is True
        assert record.run_id is None
    finally:
        manager.shutdown()


def test_worker_ignoring_cancellation_is_terminated_then_killed(
    make_manager: ManagerFactory,
) -> None:
    manager = make_manager(
        _stubborn_worker, cancel_grace_seconds=0.3, terminate_join_seconds=0.3
    )
    stubborn = _request(manager)
    manager.submit(stubborn)
    _wait_until(_has_status(manager, stubborn.request_id, RunStatus.RUNNING))
    manager.cancel(stubborn.request_id)

    assert manager.wait(stubborn.request_id, timeout=60.0).status is RunStatus.CANCELLED

    # The slot is released, so the next request still reaches the worker.
    follow_up = _request(manager)
    manager.submit(follow_up)
    _wait_until(_has_status(manager, follow_up.request_id, RunStatus.RUNNING))
    manager.cancel(follow_up.request_id)
    assert manager.wait(follow_up.request_id, timeout=60.0).status is RunStatus.CANCELLED


# --- crash, silence, failure -------------------------------------------------


def test_worker_crash_records_a_deterministic_failed_state(
    make_manager: ManagerFactory,
) -> None:
    scenario = _scenario("worker-crash")
    manager = make_manager(_crash_worker)
    request = _request(manager)
    manager.submit(request)

    record = manager.wait(request.request_id)

    assert record.status.value == scenario["expected_terminal_status"]
    assert "exit code 3" in record.error
    assert record.run_id is None
    follow_up = _request(manager)
    manager.submit(follow_up)
    assert manager.wait(follow_up.request_id).status.value == "failed"


def test_worker_exit_without_a_terminal_message_fails(
    make_manager: ManagerFactory,
) -> None:
    manager = make_manager(_silent_worker)
    request = _request(manager)
    manager.submit(request)

    record = manager.wait(request.request_id)

    assert record.status is RunStatus.FAILED
    assert "without a result" in record.error


def test_reported_failures_are_sanitized_before_persistence(
    make_manager: ManagerFactory,
) -> None:
    manager = make_manager(_failing_worker)
    request = _request(manager)
    manager.submit(request)

    record = manager.wait(request.request_id)

    assert record.status is RunStatus.FAILED
    assert record.error == "RuntimeError: cannot read quarter.xlsx"
    assert "/private/" not in record.error


# --- privacy and restart -----------------------------------------------------


def test_package_manifest_request_is_primitive_json_round_trip(
    make_manager: ManagerFactory,
) -> None:
    manager = make_manager(_sequenced_worker)
    manifest = PackageManifest.from_role_files(
        {"current_excel:ops": Path("ops.xlsx")}
    )
    request = _request(
        manager,
        files={"current_excel:ops": "/managed/ops.xlsx"},
        display_files={"current_excel:ops": "ops.xlsx"},
        package_manifest=manifest.model_dump(mode="json"),
        compare_member_sheets={"ops": ("Data", "Summary")},
    )

    payload = json.loads(json.dumps(asdict(request)))

    assert payload["package_manifest"] == manifest.model_dump(mode="json")
    assert payload["compare_member_sheets"] == {
        "ops": ["Data", "Summary"]
    }
    assert payload["files"] == {
        "current_excel:ops": "/managed/ops.xlsx"
    }
    assert "password" not in json.dumps(payload).casefold()


def test_member_password_reaches_worker_only_and_never_queue_state(
    make_manager: ManagerFactory,
) -> None:
    secret = "member-secret-not-persisted"
    manager = make_manager(_credential_probe_worker)
    request = _request(
        manager,
        files={"current_excel:ops": "/managed/ops.xlsx"},
        display_files={"current_excel:ops": "ops.xlsx"},
        package_manifest=PackageManifest.from_role_files(
            {"current_excel:ops": Path("ops.xlsx")}
        ).model_dump(mode="json"),
    )
    manager.submit(request, {"current_excel:ops": secret})

    assert manager.wait(request.request_id).status is RunStatus.SUCCEEDED
    assert (manager.work_dir / "credential-roles.txt").read_text(
        encoding="utf-8"
    ) == "current_excel:ops"
    record = manager.store.get(request.request_id)
    assert record is not None
    assert secret not in json.dumps(record.files)
    assert secret.encode() not in (manager.work_dir / "history.sqlite3").read_bytes()


def test_worker_rejects_malformed_package_manifest_before_running(
    tmp_path: Path,
) -> None:
    events: queue_module.Queue[dict[str, Any]] = queue_module.Queue()
    credentials = {"current_excel:Bad ID": "secret"}
    payload = {
        "request_id": "bad-manifest",
        "work_dir": str(tmp_path),
        "mode": QCRunMode.CURRENT_FILE_PREFLIGHT.value,
        "profile_name": "fixture",
        "profile": {"name": "fixture"},
        "files": {"current_excel:Bad ID": str(tmp_path / "missing.xlsx")},
        "display_files": {"current_excel:Bad ID": "missing.xlsx"},
        "allow_large_workbooks": False,
        "acceptance_absolute": 0.0,
        "acceptance_relative": 0.0,
        "compare_sheets": (),
        "compare_slides": (),
        "rerun_of": None,
        "package_manifest": {
            "version": 1,
            "members": [
                {
                    "member_id": "Bad ID",
                    "side": "current",
                    "artifact": "excel",
                    "display_name": "missing.xlsx",
                }
            ],
        },
        "compare_member_sheets": {},
    }

    worker_main(
        payload,
        credentials,
        cast(Any, events),
        cast(Any, threading.Event()),
    )
    message = events.get_nowait()

    assert message["kind"] == "error"
    assert "ValidationError" in message["error"]
    assert credentials == {}
    assert RunHistory(tmp_path / "history.sqlite3").list_runs() == []


def test_passwords_never_reach_persisted_or_serialized_queue_state(
    make_manager: ManagerFactory,
) -> None:
    secret = "n0t-in-sqlite-please"
    manager = make_manager(_credential_probe_worker)
    request = _request(manager)
    manager.submit(request, {"current_excel": secret})

    assert manager.wait(request.request_id).status is RunStatus.SUCCEEDED
    roles = (manager.work_dir / "credential-roles.txt").read_text(encoding="utf-8")
    record = manager.store.get(request.request_id)

    assert roles == "current_excel"  # the worker received the credential bundle
    assert record is not None
    assert secret not in json.dumps(record.files)
    assert secret not in record.error + record.detail + record.phase
    database = (manager.work_dir / "history.sqlite3").read_bytes()
    assert secret.encode() not in database


def test_password_protected_runs_are_refused_rather_than_queued(
    make_manager: ManagerFactory,
) -> None:
    manager = make_manager(_cooperative_worker)
    active = _request(manager)
    manager.submit(active)

    with pytest.raises(QueueBusyError, match="never queued"):
        manager.submit(_request(manager), {"current_excel": "secret"})

    manager.cancel(active.request_id)
    assert manager.wait(active.request_id).status is RunStatus.CANCELLED


def test_restart_orphans_incomplete_requests_and_never_resumes_them(
    tmp_path: Path,
) -> None:
    work_dir = tmp_path / "restarted"
    work_dir.mkdir()
    store = RunStateStore(work_dir / "history.sqlite3")
    interrupted = store.enqueue(
        "stale-request",
        mode=QCRunMode.CYCLE_COMPARISON.value,
        profile="fixture",
        files={"current_excel": "current.xlsx"},
        queue_position=0,
    )
    store.mark_starting(interrupted.request_id)

    manager = RunQueueManager(work_dir, worker=_sequenced_worker)
    try:
        orphaned = manager.store.get("stale-request")
        assert orphaned is not None
        assert orphaned.status is RunStatus.ORPHANED
        assert orphaned.finished_at is not None

        fresh = _request(manager)
        manager.submit(fresh)
        assert manager.wait(fresh.request_id).status is RunStatus.SUCCEEDED

        unchanged = manager.store.get("stale-request")
        assert unchanged is not None
        assert unchanged.status is RunStatus.ORPHANED
        assert unchanged.run_id is None
        log = (work_dir / "execution.log").read_text(encoding="utf-8")
        assert "stale-request" not in log
    finally:
        manager.shutdown()


# --- progress, telemetry, backpressure ---------------------------------------


def test_progress_and_phase_telemetry_persist_for_reconnected_pages(
    make_manager: ManagerFactory,
) -> None:
    manager = make_manager(_telemetry_worker)
    request = _request(manager)
    manager.submit(request)

    record = manager.wait(request.request_id)

    assert record.status is RunStatus.SUCCEEDED
    assert record.run_id == _FAKE_RUN_ID
    assert (record.phase, record.processed, record.total) == ("comparing_formulas", 3, 7)
    assert record.phases[0]["phase"] == "comparing_formulas"
    assert record.phases[0]["peak_rss_bytes"] == 1024
    assert record.elapsed_seconds() > 0


def test_bounded_channel_drops_progress_without_stalling_the_run(
    make_manager: ManagerFactory,
) -> None:
    manager = make_manager(_flood_worker)
    request = _request(manager)
    manager.submit(request)

    record = manager.wait(request.request_id)

    assert IPC_QUEUE_CAPACITY > 0
    assert record.status is RunStatus.SUCCEEDED
    assert record.phase == "comparing_formulas"


# --- real spawn integration --------------------------------------------------


def test_real_qc_runs_in_an_owned_spawned_worker(
    make_manager: ManagerFactory, fixture_dir: Path
) -> None:
    manager = make_manager(worker_main)
    profile: DeliverableProfile = fixture_profile()
    files = {
        "baseline_excel": str(fixture_dir / "baseline.xlsx"),
        "current_excel": str(fixture_dir / "current.xlsx"),
    }
    request = _request(
        manager,
        mode=QCRunMode.CYCLE_COMPARISON.value,
        profile_name=profile.name,
        profile=profile.model_dump(mode="json"),
        files=files,
        display_files={role: Path(path).name for role, path in files.items()},
    )
    manager.submit(request)

    record = manager.wait(request.request_id, timeout=300.0)

    assert record.status is RunStatus.SUCCEEDED, record.error
    assert record.run_id is not None
    stored = RunHistory(manager.work_dir / "history.sqlite3").get_run(record.run_id)
    assert stored.findings
    assert all(Path(path).exists() for path in stored.report_paths.values())
    assert record.phases, "phase telemetry is captured inside the worker process"
