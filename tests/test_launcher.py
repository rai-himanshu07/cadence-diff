"""Authenticated local-instance and quiet-launcher contracts."""

from __future__ import annotations

import json
import logging
from pathlib import Path
from types import SimpleNamespace

import pytest

import qc_tool.launcher as launcher_module
from qc_tool.launcher import (
    InstanceAlreadyRunningError,
    InstanceMarker,
    LaunchOutcome,
    active_instance_health,
    claim_local_instance,
    configure_launcher_logging,
    instance_proof,
    launch_local_app,
    launch_request_lock,
    probe_instance,
    read_instance_marker,
    release_local_instance,
    write_instance_marker,
)


class _Response:
    def __init__(self, payload: dict[str, object]) -> None:
        self._payload = json.dumps(payload).encode("utf-8")

    def __enter__(self) -> _Response:
        return self

    def __exit__(self, *_args: object) -> None:
        return None

    def read(self, _limit: int = -1) -> bytes:
        return self._payload


def _marker(secret: str = "s" * 64) -> InstanceMarker:
    return InstanceMarker(
        schema_version=1,
        pid=123,
        port=8080,
        secret=secret,
    )


def test_instance_marker_round_trips_without_path_data(tmp_path: Path) -> None:
    marker = _marker()

    write_instance_marker(tmp_path, marker)

    assert read_instance_marker(tmp_path) == marker
    payload = (tmp_path / ".instance.json").read_text(encoding="utf-8")
    assert "path" not in payload.casefold()


def test_quiet_log_keeps_only_fixed_launcher_diagnostics(tmp_path: Path) -> None:
    path = configure_launcher_logging(tmp_path)
    root = logging.getLogger()
    created = [
        handler
        for handler in root.handlers
        if getattr(handler, "baseFilename", "") == str(path.resolve())
    ]
    try:
        logging.getLogger("qc_tool.excel.formulas").warning(
            "unparseable formula: =PRIVATE_CLIENT_PATH"
        )
        logging.getLogger("qc_tool.launcher").error("launch-start-timeout")
        for handler in created:
            handler.flush()

        content = path.read_text(encoding="utf-8")
        assert "launch-start-timeout" in content
        assert "PRIVATE_CLIENT_PATH" not in content
    finally:
        for handler in created:
            root.removeHandler(handler)
            handler.close()


def test_probe_requires_exact_challenge_response() -> None:
    marker = _marker()

    def valid_open(request, *, timeout: float):
        assert timeout > 0
        challenge = request.full_url.partition("challenge=")[2]
        return _Response(
            {
                "schema_version": 1,
                "proof": instance_proof(marker.secret, challenge),
            }
        )

    def lookalike_open(_request, *, timeout: float):
        assert timeout > 0
        return _Response({"schema_version": 1, "proof": "0" * 64})

    assert probe_instance(marker, opener=valid_open)
    assert not probe_instance(marker, opener=lookalike_open)


def test_instance_authority_proves_health_and_cleans_its_marker(
    tmp_path: Path,
) -> None:
    authority = claim_local_instance(tmp_path, 8080)
    challenge = "ab" * 32
    try:
        payload = active_instance_health(challenge)
        assert payload == {
            "schema_version": 1,
            "proof": instance_proof(authority.marker.secret, challenge),
        }
        with pytest.raises(InstanceAlreadyRunningError):
            claim_local_instance(tmp_path, 8080)
    finally:
        release_local_instance(authority)

    assert active_instance_health(challenge) is None
    assert read_instance_marker(tmp_path) is None


def test_launcher_reuses_only_authenticated_instance(tmp_path: Path) -> None:
    marker = _marker()
    opened: list[str] = []

    result = launch_local_app(
        tmp_path,
        port=8080,
        read_marker=lambda _root: marker,
        probe=lambda _marker: True,
        port_open=lambda _port: True,
        spawn=lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("must not spawn")
        ),
        open_browser=opened.append,
    )

    assert result.outcome is LaunchOutcome.REUSED
    assert opened == ["http://127.0.0.1:8080/"]


def test_launcher_refuses_unrelated_port_listener(tmp_path: Path) -> None:
    opened: list[str] = []

    result = launch_local_app(
        tmp_path,
        port=8080,
        read_marker=lambda _root: None,
        probe=lambda _marker: False,
        port_open=lambda _port: True,
        spawn=lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("must not spawn")
        ),
        open_browser=opened.append,
    )

    assert result.outcome is LaunchOutcome.PORT_OCCUPIED
    assert opened == []


def test_launcher_reports_spawn_failure_without_private_detail(tmp_path: Path) -> None:
    result = launch_local_app(
        tmp_path,
        port_open=lambda _port: False,
        spawn=lambda *_args: (_ for _ in ()).throw(
            FileNotFoundError("/private/python/path")
        ),
        open_browser=lambda _url: None,
    )

    assert result.outcome is LaunchOutcome.START_FAILED
    assert "FileNotFoundError" in result.detail
    assert "/private" not in result.detail


def test_spawned_server_is_forced_to_local_mode(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    commands: list[list[str]] = []

    def fake_popen(command: list[str], **_kwargs):
        commands.append(command)
        return SimpleNamespace(poll=lambda: None)

    monkeypatch.setattr(launcher_module.subprocess, "Popen", fake_popen)
    monkeypatch.setattr(launcher_module, "server_interpreter", lambda: Path("python.exe"))

    launcher_module.spawn_local_server(tmp_path, 9123)

    assert commands[0][0] == "python.exe"
    assert "--network" in commands[0]
    assert commands[0][commands[0].index("--network") + 1] == "local"


def test_concurrent_launch_request_does_not_spawn_or_open(tmp_path: Path) -> None:
    opened: list[str] = []
    with launch_request_lock(tmp_path) as acquired:
        assert acquired
        result = launch_local_app(
            tmp_path,
            spawn=lambda *_args, **_kwargs: (_ for _ in ()).throw(
                AssertionError("must not spawn")
            ),
            open_browser=opened.append,
        )

    assert result.outcome is LaunchOutcome.IN_PROGRESS
    assert opened == []


def test_launcher_starts_current_environment_and_opens_after_proof(
    tmp_path: Path,
) -> None:
    marker = InstanceMarker(
        schema_version=1,
        pid=123,
        port=9123,
        secret="s" * 64,
    )
    reads = iter([None, marker])
    spawned: list[tuple[Path, int]] = []
    opened: list[str] = []

    result = launch_local_app(
        tmp_path,
        port=9123,
        read_marker=lambda _root: next(reads, marker),
        probe=lambda observed: observed == marker,
        port_open=lambda _port: False,
        spawn=lambda root, port: spawned.append((root, port))
        or SimpleNamespace(poll=lambda: None),
        open_browser=opened.append,
        deadline_seconds=1.0,
        monotonic=iter([0.0, 0.1, 0.2]).__next__,
        wait=lambda _seconds: None,
    )

    assert result.outcome is LaunchOutcome.STARTED
    assert spawned == [(tmp_path, 9123)]
    assert opened == ["http://127.0.0.1:9123/"]
