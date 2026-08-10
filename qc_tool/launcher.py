"""Authenticated local-instance discovery and quiet application launching."""

from __future__ import annotations

import hashlib
import hmac
import json
import logging
import os
import secrets
import socket
import subprocess
import sys
import tempfile
import time
import urllib.request
import webbrowser
from collections.abc import Callable
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from enum import StrEnum
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any

from qc_tool.security import private_directory, private_file

INSTANCE_SCHEMA_VERSION = 1
INSTANCE_FILENAME = ".instance.json"
HEALTH_PATH = "/_qc/instance"
MAX_HEALTH_BYTES = 4096
LOCK_FILENAME = ".instance.lock"
LAUNCH_LOCK_FILENAME = ".launch.lock"


@dataclass(frozen=True, slots=True)
class InstanceMarker:
    schema_version: int
    pid: int
    port: int
    secret: str


class LaunchOutcome(StrEnum):
    STARTED = "started"
    REUSED = "reused"
    PORT_OCCUPIED = "port_occupied"
    START_FAILED = "start_failed"
    START_TIMEOUT = "start_timeout"
    IN_PROGRESS = "in_progress"


@dataclass(frozen=True, slots=True)
class LaunchResult:
    outcome: LaunchOutcome
    url: str
    detail: str = ""


class InstanceAlreadyRunningError(ValueError):
    """Another server process owns this data directory."""


@dataclass(slots=True)
class InstanceAuthority:
    data_dir: Path
    marker: InstanceMarker
    lock_handle: Any


_active_authority: InstanceAuthority | None = None


class _LauncherLogFilter(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        return record.name in {"qc_tool.launcher", "qc_tool.cli"}


def instance_marker_path(data_dir: Path) -> Path:
    return data_dir / INSTANCE_FILENAME


def _lock_instance_file(handle: Any) -> None:
    handle.seek(0)
    if os.name == "nt":
        import msvcrt

        msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
    else:
        import fcntl

        fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)


def _unlock_instance_file(handle: Any) -> None:
    handle.seek(0)
    if os.name == "nt":
        import msvcrt

        msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
    else:
        import fcntl

        fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


@contextmanager
def launch_request_lock(data_dir: Path):
    """Yield whether this process owns the one active launch request."""
    private_directory(data_dir)
    path = data_dir / LAUNCH_LOCK_FILENAME
    handle = path.open("a+b")
    if path.stat().st_size == 0:
        handle.write(b"0")
        handle.flush()
    private_file(path)
    try:
        _lock_instance_file(handle)
    except OSError:
        handle.close()
        yield False
        return
    try:
        yield True
    finally:
        _unlock_instance_file(handle)
        handle.close()


def claim_local_instance(data_dir: Path, port: int) -> InstanceAuthority:
    """Claim one local server identity until the returned authority is released."""
    global _active_authority
    if _active_authority is not None:
        raise InstanceAlreadyRunningError("QC Tool is already running")
    private_directory(data_dir)
    lock_path = data_dir / LOCK_FILENAME
    lock_handle = lock_path.open("a+b")
    if lock_path.stat().st_size == 0:
        lock_handle.write(b"0")
        lock_handle.flush()
    private_file(lock_path)
    try:
        _lock_instance_file(lock_handle)
    except OSError as exc:
        lock_handle.close()
        raise InstanceAlreadyRunningError(
            "QC Tool is already running for this data directory"
        ) from exc
    marker = InstanceMarker(
        schema_version=INSTANCE_SCHEMA_VERSION,
        pid=os.getpid(),
        port=port,
        secret=secrets.token_urlsafe(48),
    )
    authority = InstanceAuthority(data_dir, marker, lock_handle)
    try:
        write_instance_marker(data_dir, marker)
    except BaseException:
        _unlock_instance_file(lock_handle)
        lock_handle.close()
        raise
    _active_authority = authority
    return authority


def release_local_instance(authority: InstanceAuthority) -> None:
    """Release one authority without deleting a newer process's marker."""
    global _active_authority
    if _active_authority is authority:
        _active_authority = None
    current = read_instance_marker(authority.data_dir)
    if current is not None and hmac.compare_digest(
        current.secret,
        authority.marker.secret,
    ):
        instance_marker_path(authority.data_dir).unlink(missing_ok=True)
    try:
        _unlock_instance_file(authority.lock_handle)
    finally:
        authority.lock_handle.close()


def active_instance_health(challenge: str) -> dict[str, object] | None:
    """Return a challenge proof only while a local server owns the marker."""
    authority = _active_authority
    if authority is None or len(challenge) != 64:
        return None
    try:
        bytes.fromhex(challenge)
    except ValueError:
        return None
    return {
        "schema_version": INSTANCE_SCHEMA_VERSION,
        "proof": instance_proof(authority.marker.secret, challenge),
    }


def write_instance_marker(data_dir: Path, marker: InstanceMarker) -> Path:
    """Atomically write the private local-instance claim."""
    private_directory(data_dir)
    destination = instance_marker_path(data_dir)
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=data_dir,
            prefix=f".{INSTANCE_FILENAME}.",
            suffix=".tmp",
            delete=False,
        ) as handle:
            json.dump(asdict(marker), handle, sort_keys=True, separators=(",", ":"))
            temporary = Path(handle.name)
        private_file(temporary)
        os.replace(temporary, destination)
        temporary = None
        private_file(destination)
        return destination
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def read_instance_marker(data_dir: Path) -> InstanceMarker | None:
    """Read a marker fail-closed; malformed or stale shape is not authority."""
    path = instance_marker_path(data_dir)
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        marker = InstanceMarker(
            schema_version=payload["schema_version"],
            pid=payload["pid"],
            port=payload["port"],
            secret=payload["secret"],
        )
    except (KeyError, OSError, TypeError, ValueError):
        return None
    if (
        marker.schema_version != INSTANCE_SCHEMA_VERSION
        or isinstance(marker.pid, bool)
        or not isinstance(marker.pid, int)
        or marker.pid <= 0
        or isinstance(marker.port, bool)
        or not isinstance(marker.port, int)
        or not 1 <= marker.port <= 65535
        or not isinstance(marker.secret, str)
        or len(marker.secret) < 32
    ):
        return None
    return marker


def instance_proof(secret: str, challenge: str) -> str:
    return hmac.new(
        secret.encode("utf-8"),
        challenge.encode("ascii"),
        hashlib.sha256,
    ).hexdigest()


def probe_instance(
    marker: InstanceMarker,
    *,
    timeout: float = 0.5,
    opener: Callable[..., Any] = urllib.request.urlopen,
) -> bool:
    """Prove a listener knows the private marker secret without sending it."""
    challenge = secrets.token_hex(32)
    request = urllib.request.Request(
        f"http://127.0.0.1:{marker.port}{HEALTH_PATH}?challenge={challenge}",
        headers={"Accept": "application/json"},
    )
    try:
        with opener(request, timeout=timeout) as response:
            payload = json.loads(response.read(MAX_HEALTH_BYTES))
        proof = payload["proof"]
        version = payload["schema_version"]
    except (KeyError, OSError, TypeError, ValueError):
        return False
    return (
        version == INSTANCE_SCHEMA_VERSION
        and isinstance(proof, str)
        and hmac.compare_digest(proof, instance_proof(marker.secret, challenge))
    )


def local_port_open(port: int, *, timeout: float = 0.2) -> bool:
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=timeout):
            return True
    except OSError:
        return False


def launcher_log_path(data_dir: Path) -> Path:
    return data_dir / "logs" / "server.log"


def configure_launcher_logging(data_dir: Path) -> Path:
    """Install one private rotating file handler before server imports."""
    log_path = launcher_log_path(data_dir)
    private_directory(log_path.parent)
    resolved = str(log_path.resolve())
    root = logging.getLogger()
    if not any(
        isinstance(handler, RotatingFileHandler)
        and getattr(handler, "baseFilename", "") == resolved
        for handler in root.handlers
    ):
        handler = RotatingFileHandler(
            log_path,
            maxBytes=1_000_000,
            backupCount=3,
            encoding="utf-8",
        )
        handler.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s")
        )
        handler.addFilter(_LauncherLogFilter())
        root.addHandler(handler)
    private_file(log_path)
    return log_path


def quiet_interpreter() -> Path:
    executable = Path(sys.executable)
    if sys.platform == "win32":
        candidate = executable.with_name("pythonw.exe")
        if candidate.exists():
            return candidate
    return executable


def server_interpreter() -> Path:
    executable = Path(sys.executable)
    if sys.platform == "win32":
        candidate = executable.with_name("python.exe")
        if candidate.exists():
            return candidate
    return executable


def spawn_local_server(data_dir: Path, port: int) -> subprocess.Popen[Any]:
    command = [
        str(server_interpreter()),
        "-m",
        "qc_tool",
        "serve",
        "--data-dir",
        str(data_dir),
        "--port",
        str(port),
        "--network",
        "local",
        "--no-browser",
        "--launcher-child",
    ]
    kwargs: dict[str, Any] = {
        "stdin": subprocess.DEVNULL,
        "stdout": subprocess.DEVNULL,
        "stderr": subprocess.DEVNULL,
        "close_fds": True,
    }
    if sys.platform == "win32":
        kwargs["creationflags"] = (
            subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NO_WINDOW
        )
    else:
        kwargs["start_new_session"] = True
    return subprocess.Popen(command, **kwargs)


def launch_local_app(
    data_dir: Path,
    *,
    port: int = 8080,
    deadline_seconds: float = 15.0,
    read_marker: Callable[[Path], InstanceMarker | None] = read_instance_marker,
    probe: Callable[[InstanceMarker], bool] = probe_instance,
    port_open: Callable[[int], bool] = local_port_open,
    spawn: Callable[[Path, int], Any] = spawn_local_server,
    open_browser: Callable[[str], Any] = webbrowser.open,
    monotonic: Callable[[], float] = time.monotonic,
    wait: Callable[[float], None] = time.sleep,
) -> LaunchResult:
    """Reuse an authenticated server or start this environment and await proof."""
    private_directory(data_dir)
    url = f"http://127.0.0.1:{port}/"
    with launch_request_lock(data_dir) as owns_request:
        if not owns_request:
            return LaunchResult(
                LaunchOutcome.IN_PROGRESS,
                url,
                "another launch request is already in progress",
            )
        return _launch_local_app_owned(
            data_dir,
            port=port,
            url=url,
            deadline_seconds=deadline_seconds,
            read_marker=read_marker,
            probe=probe,
            port_open=port_open,
            spawn=spawn,
            open_browser=open_browser,
            monotonic=monotonic,
            wait=wait,
        )


def _launch_local_app_owned(
    data_dir: Path,
    *,
    port: int,
    url: str,
    deadline_seconds: float,
    read_marker: Callable[[Path], InstanceMarker | None],
    probe: Callable[[InstanceMarker], bool],
    port_open: Callable[[int], bool],
    spawn: Callable[[Path, int], Any],
    open_browser: Callable[[str], Any],
    monotonic: Callable[[], float],
    wait: Callable[[float], None],
) -> LaunchResult:
    marker = read_marker(data_dir)
    if marker is not None and marker.port == port and probe(marker):
        open_browser(url)
        return LaunchResult(LaunchOutcome.REUSED, url)
    if port_open(port):
        return LaunchResult(
            LaunchOutcome.PORT_OCCUPIED,
            url,
            "another process is using the configured port",
        )
    try:
        process = spawn(data_dir, port)
    except OSError as exc:
        return LaunchResult(
            LaunchOutcome.START_FAILED,
            url,
            f"could not start the local server ({type(exc).__name__})",
        )
    deadline = monotonic() + deadline_seconds
    while monotonic() < deadline:
        marker = read_marker(data_dir)
        if marker is not None and marker.port == port and probe(marker):
            open_browser(url)
            return LaunchResult(LaunchOutcome.STARTED, url)
        exit_code = process.poll()
        if exit_code is not None:
            return LaunchResult(
                LaunchOutcome.START_FAILED,
                url,
                f"server exited with code {exit_code}",
            )
        wait(0.05)
    return LaunchResult(
        LaunchOutcome.START_TIMEOUT,
        url,
        "server did not become ready before the launch deadline",
    )
