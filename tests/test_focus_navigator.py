"""Step 4 contracts: the short-lived helper protocol and FocusNavigator."""

import ast
import json
import subprocess
import sys
from pathlib import Path

import pytest

from qc_tool.focus.discovery import (
    DiscoveryResult,
    FocusApplication,
    OpenDocument,
)
from qc_tool.focus.navigator import (
    DEFAULT_HELPER_TIMEOUT_SECONDS,
    FOCUS_COOLDOWN_SECONDS,
    MEASURED_FOCUS_P99_SECONDS,
    FocusNavigator,
    FocusReply,
    HelperRun,
)
from qc_tool.focus.protocol import (
    DISPATCHED_STAGES,
    SCHEMA_VERSION,
    FocusAction,
    FocusOutcome,
    FocusStage,
    canonical_path,
    fixed_code,
    path_digest,
)
from qc_tool.focus.worker import handle_request, locate_bound_document

REPO_ROOT = Path(__file__).resolve().parents[1]


def _excel(**kwargs: object) -> OpenDocument:
    defaults: dict[str, object] = {
        "application": FocusApplication.EXCEL,
        "process_id": 4242,
        "process_created": 1.0,
        "windows_session_id": 1,
        "full_name": "/work/current.xlsx",
        "window_count": 1,
        "visible_window_count": 1,
        "visible_window_handles": (101,),
        "saved": True,
        "autosave": False,
    }
    defaults.update(kwargs)
    return OpenDocument(**defaults)  # type: ignore[arg-type]


def _discovery(*documents: OpenDocument) -> DiscoveryResult:
    return DiscoveryResult(application=FocusApplication.EXCEL, documents=documents)


# --------------------------------------------------------------------------
# protocol
# --------------------------------------------------------------------------


def test_fixed_code_cannot_carry_a_path_or_message() -> None:
    assert fixed_code(r"C:\work\Q3 model.xlsx") == "c_work_q3_model_xlsx"
    assert fixed_code("Focus FAILED: /home/a/b") == "focus_failed_home_a_b"
    assert fixed_code("") == "unknown"
    assert len(fixed_code("x" * 500)) == 64
    assert fixed_code(FocusOutcome.FOCUSED) == "focused"


def test_path_digest_is_salted_and_reveals_nothing() -> None:
    first = path_digest(r"C:\work\a.xlsx", b"salt-one")
    second = path_digest(r"C:\work\a.xlsx", b"salt-two")
    assert first != second
    assert first == path_digest("C:/work/a.xlsx", b"salt-one")
    assert "work" not in first


def test_canonical_path_normalises_separators_and_case() -> None:
    assert canonical_path(" /tmp/./a/../a.xlsx ") == canonical_path("/tmp/a.xlsx")


def test_dispatched_stages_are_only_the_post_action_ones() -> None:
    assert set(DISPATCHED_STAGES) == {
        FocusStage.ACTION_STARTED,
        FocusStage.ACTION_FINISHED,
    }


# --------------------------------------------------------------------------
# worker
# --------------------------------------------------------------------------


def test_worker_guards_its_entry_point_on_main_only() -> None:
    """A spawn child re-imports __main__ as __mp_main__; never match that."""
    source = (REPO_ROOT / "qc_tool" / "focus" / "worker.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    guards = [node for node in tree.body if isinstance(node, ast.If)]
    assert len(guards) == 1
    assert ast.unparse(guards[0].test) == "__name__ == '__main__'"


def test_worker_refuses_an_unknown_schema_or_action() -> None:
    assert handle_request({"schema_version": 99})["outcome"] == (
        FocusOutcome.INVALID_REQUEST.value
    )
    assert handle_request(
        {"schema_version": SCHEMA_VERSION, "action": "delete_everything"}
    )["outcome"] == FocusOutcome.INVALID_REQUEST.value
    assert handle_request(
        {
            "schema_version": SCHEMA_VERSION,
            "action": FocusAction.DISCOVER.value,
            "application": "word",
        }
    )["outcome"] == FocusOutcome.INVALID_REQUEST.value


def test_worker_discovery_refuses_off_windows(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from qc_tool.focus import discovery

    monkeypatch.setattr(discovery.sys, "platform", "linux")
    reply = handle_request(
        {
            "schema_version": SCHEMA_VERSION,
            "action": FocusAction.DISCOVER.value,
            "application": "excel",
        }
    )
    assert reply["outcome"] == FocusOutcome.DISCOVERED.value
    discovery = reply["discovery"]
    assert isinstance(discovery, dict)
    assert discovery["reasons"] == ["unsupported_platform"]
    assert discovery["documents"] == []


def test_worker_records_stages_in_a_private_file(tmp_path: Path) -> None:
    stage_path = tmp_path / "stage.txt"
    handle_request(
        {
            "schema_version": SCHEMA_VERSION,
            "action": FocusAction.DISCOVER.value,
            "application": "excel",
        },
        stage_path=stage_path,
    )
    assert stage_path.read_text(encoding="utf-8") == FocusStage.DISCOVERING.value


def test_locate_bound_document_requires_every_identity_attribute() -> None:
    salt = b"\x01\x02"
    document = _excel(full_name="/work/a.xlsx")
    request = {
        "path_salt": salt.hex(),
        "expected_path_digest": path_digest(document.full_name, salt),
        "process_id": document.process_id,
        "process_created": document.process_created,
        "windows_session_id": document.windows_session_id,
        "window_handle": document.visible_window_handles[0],
    }
    discovery = _discovery(document)
    assert locate_bound_document(discovery, request) is document
    for field, wrong in (
        ("process_id", 1),
        ("process_created", 99.0),
        ("windows_session_id", 7),
        ("window_handle", 999),
        ("expected_path_digest", "0" * 64),
    ):
        assert locate_bound_document(discovery, {**request, field: wrong}) is None


def test_locate_bound_document_refuses_a_malformed_request() -> None:
    assert locate_bound_document(_discovery(_excel()), {}) is None
    assert (
        locate_bound_document(_discovery(_excel()), {"path_salt": "not-hex"}) is None
    )


def test_helper_module_never_calls_a_mutating_office_method() -> None:
    source = (REPO_ROOT / "qc_tool" / "focus").rglob("*.py")
    forbidden = (
        ".Save(",
        ".SaveAs(",
        ".Calculate(",
        ".Refresh(",
        ".Close(",
        ".Quit(",
        "EnableEvents",
        "ScreenUpdating",
        "DisplayAlerts",
        "Calculation =",
    )
    for path in source:
        text = path.read_text(encoding="utf-8")
        for token in forbidden:
            assert token not in text, f"{path.name} contains {token}"


# --------------------------------------------------------------------------
# navigator
# --------------------------------------------------------------------------


def _request(action: FocusAction = FocusAction.FOCUS) -> dict[str, object]:
    return {
        "schema_version": SCHEMA_VERSION,
        "action": action.value,
        "application": "excel",
    }


def _runner(run: HelperRun):
    calls: list[dict[str, object]] = []

    def runner(request: dict[str, object], timeout: float) -> HelperRun:
        calls.append({"request": request, "timeout": timeout})
        return run

    runner.calls = calls  # type: ignore[attr-defined]
    return runner


async def test_successful_reply_is_passed_through() -> None:
    runner = _runner(
        HelperRun(
            payload={"outcome": "focused", "unsaved_changes": True},
            stage=FocusStage.ACTION_FINISHED,
            duration_seconds=0.4,
        )
    )
    navigator = FocusNavigator(runner=runner)
    reply = await navigator.submit(_request())
    assert reply.outcome is FocusOutcome.FOCUSED
    assert reply.unsaved_changes is True
    assert reply.dispatched


async def test_unknown_helper_outcome_is_reduced_to_a_fixed_code() -> None:
    navigator = FocusNavigator(
        runner=_runner(HelperRun(payload={"outcome": r"boom C:\secret.xlsx"}))
    )
    reply = await navigator.submit(_request())
    assert reply.code == "boom_c_secret_xlsx"


async def test_helper_crash_is_reported_without_a_result_file() -> None:
    navigator = FocusNavigator(runner=_runner(HelperRun(payload=None, exit_code=3)))
    reply = await navigator.submit(_request())
    assert reply.outcome is FocusOutcome.HELPER_CRASHED


async def test_silent_helper_exit_is_reported_as_failed() -> None:
    navigator = FocusNavigator(runner=_runner(HelperRun(payload=None, exit_code=0)))
    reply = await navigator.submit(_request())
    assert reply.outcome is FocusOutcome.HELPER_FAILED


async def test_stale_or_unversioned_result_is_not_trusted(tmp_path: Path) -> None:
    from qc_tool.focus.navigator import _read_result

    path = tmp_path / "result.json"
    path.write_text(json.dumps({"outcome": "focused"}), encoding="utf-8")
    assert _read_result(path) is None
    path.write_text("{not json", encoding="utf-8")
    assert _read_result(path) is None


async def test_timeout_before_dispatch_reports_no_action_and_cools_down() -> None:
    clock = {"now": 100.0}
    navigator = FocusNavigator(
        runner=_runner(HelperRun(timed_out=True, stage=FocusStage.HASHING)),
        clock=lambda: clock["now"],
    )
    reply = await navigator.submit(_request())
    assert reply.outcome is FocusOutcome.TIMEOUT_NO_ACTION_DISPATCHED
    assert not reply.dispatched
    assert not navigator.disabled
    assert navigator.cooling_down()
    cooling = await navigator.submit(_request())
    assert cooling.outcome is FocusOutcome.COOLING_DOWN
    clock["now"] += FOCUS_COOLDOWN_SECONDS
    assert not navigator.cooling_down()


async def test_first_request_after_cooldown_runs_a_no_action_health_check() -> None:
    clock = {"now": 0.0}
    replies = [
        HelperRun(timed_out=True, stage=FocusStage.DISCOVERING),
        HelperRun(payload={"outcome": "healthy", "complete": True, "reasons": []}),
        HelperRun(payload={"outcome": "focused"}),
    ]
    seen: list[dict[str, object]] = []

    def runner(request: dict[str, object], _timeout: float) -> HelperRun:
        seen.append(request)
        return replies[len(seen) - 1]

    navigator = FocusNavigator(runner=runner, clock=lambda: clock["now"])
    await navigator.submit(_request())
    clock["now"] += FOCUS_COOLDOWN_SECONDS
    reply = await navigator.submit(_request())
    assert reply.outcome is FocusOutcome.FOCUSED
    assert seen[1]["action"] == FocusAction.HEALTH_CHECK.value
    assert seen[2]["action"] == FocusAction.FOCUS.value


async def test_unhealthy_check_stops_the_replacement_request() -> None:
    clock = {"now": 0.0}
    replies = [
        HelperRun(timed_out=True, stage=FocusStage.DISCOVERING),
        HelperRun(payload={"outcome": "helper_failed"}),
    ]
    seen: list[dict[str, object]] = []

    def runner(request: dict[str, object], _timeout: float) -> HelperRun:
        seen.append(request)
        return replies[len(seen) - 1]

    navigator = FocusNavigator(runner=runner, clock=lambda: clock["now"])
    await navigator.submit(_request())
    clock["now"] += FOCUS_COOLDOWN_SECONDS
    reply = await navigator.submit(_request())
    assert reply.outcome is FocusOutcome.HELPER_FAILED
    assert len(seen) == 2


@pytest.mark.parametrize("stage", sorted(DISPATCHED_STAGES))
async def test_timeout_after_dispatch_cools_down_then_health_checks(
    stage: FocusStage,
) -> None:
    clock = {"now": 0.0}
    replies = [
        HelperRun(timed_out=True, stage=stage),
        HelperRun(payload={"outcome": "healthy", "complete": True, "reasons": []}),
        HelperRun(payload={"outcome": "focused"}),
    ]
    seen: list[dict[str, object]] = []

    def runner(request: dict[str, object], _timeout: float) -> HelperRun:
        seen.append(request)
        return replies[len(seen) - 1]

    navigator = FocusNavigator(runner=runner, clock=lambda: clock["now"])
    reply = await navigator.submit(_request())
    assert reply.outcome is FocusOutcome.TIMEOUT_ACTION_MAY_HAVE_COMPLETED
    assert reply.dispatched
    assert not navigator.disabled
    assert navigator.cooling_down()
    later = await navigator.submit(_request())
    assert later.outcome is FocusOutcome.COOLING_DOWN
    clock["now"] += FOCUS_COOLDOWN_SECONDS
    recovered = await navigator.submit(_request())
    assert recovered.outcome is FocusOutcome.FOCUSED
    assert seen[1]["action"] == FocusAction.HEALTH_CHECK.value


async def test_every_navigator_shares_one_process_global_lock() -> None:
    first = FocusNavigator(runner=_runner(HelperRun(payload={"outcome": "focused"})))
    second = FocusNavigator(runner=_runner(HelperRun(payload={"outcome": "focused"})))
    assert first._lock is second._lock


async def test_requests_are_serialised_server_wide() -> None:
    import asyncio

    order: list[str] = []

    def runner(request: dict[str, object], _timeout: float) -> HelperRun:
        tag = str(request.get("tag"))
        order.append(f"start-{tag}")
        import time as time_module

        time_module.sleep(0.02)
        order.append(f"end-{tag}")
        return HelperRun(payload={"outcome": "focused"})

    navigator = FocusNavigator(runner=runner)
    await asyncio.gather(
        navigator.submit({**_request(), "tag": "a"}),
        navigator.submit({**_request(), "tag": "b"}),
    )
    assert order[0].startswith("start-")
    assert order[1] == order[0].replace("start-", "end-")


async def test_cooldown_floor_cannot_be_lowered() -> None:
    navigator = FocusNavigator(runner=_runner(HelperRun()), cooldown_seconds=1.0)
    assert navigator._cooldown == FOCUS_COOLDOWN_SECONDS


def test_live_latency_measurement_keeps_the_thirty_second_floor() -> None:
    assert MEASURED_FOCUS_P99_SECONDS == 4.141
    assert 2 * MEASURED_FOCUS_P99_SECONDS < FOCUS_COOLDOWN_SECONDS == 30.0


def test_default_timeout_is_bounded() -> None:
    assert 0 < DEFAULT_HELPER_TIMEOUT_SECONDS < FOCUS_COOLDOWN_SECONDS


def test_reply_code_is_always_a_fixed_code() -> None:
    assert FocusReply(outcome=r"C:\a\b.xlsx").code == "c_a_b_xlsx"


# --------------------------------------------------------------------------
# real subprocess boundary
# --------------------------------------------------------------------------


def test_helper_runs_as_a_module_and_returns_a_primitive_result() -> None:
    request = {
        "schema_version": SCHEMA_VERSION,
        "action": FocusAction.DISCOVER.value,
        "application": "excel",
    }
    completed = subprocess.run(
        [sys.executable, "-m", "qc_tool.focus.worker"],
        input=json.dumps(request),
        capture_output=True,
        text=True,
        cwd=REPO_ROOT,
        timeout=60,
        check=False,
    )
    assert completed.returncode == 0
    payload = json.loads(completed.stdout)
    assert payload["schema_version"] == SCHEMA_VERSION
    assert payload["outcome"] == FocusOutcome.DISCOVERED.value
