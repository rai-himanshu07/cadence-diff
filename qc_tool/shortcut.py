"""Explicit per-user Windows desktop shortcut lifecycle."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from dataclasses import asdict, dataclass
from enum import StrEnum
from importlib.resources import files
from pathlib import Path
from typing import Protocol

from qc_tool.launcher import quiet_interpreter
from qc_tool.security import private_directory, private_file

SHORTCUT_SCHEMA_VERSION = 1
SHORTCUT_FILENAME = "QC Tool.lnk"
SHORTCUT_METADATA_FILENAME = "shortcut.json"


@dataclass(frozen=True, slots=True)
class ShortcutSpec:
    path: Path
    target: Path
    arguments: str
    working_directory: Path
    icon_location: str


class ShortcutState(StrEnum):
    INSTALLED = "installed"
    MISSING = "missing"
    STALE = "stale"
    CONFLICT = "conflict"
    UNSUPPORTED = "unsupported"


@dataclass(frozen=True, slots=True)
class ShortcutResult:
    state: ShortcutState
    spec: ShortcutSpec | None = None
    detail: str = ""


class ShortcutUnavailableError(ValueError):
    """Windows shell shortcut support is unavailable."""


class ShortcutBackend(Protocol):
    def write(self, spec: ShortcutSpec) -> None: ...

    def read(self, path: Path) -> ShortcutSpec | None: ...

    def remove(self, path: Path) -> None: ...


class WindowsShellShortcutBackend:
    """Thin pywin32 adapter; imported only after a Windows capability check."""

    @staticmethod
    def _shell():
        try:
            import win32com.client  # pyright: ignore[reportMissingModuleSource]
        except ImportError as exc:
            raise ShortcutUnavailableError(
                "pywin32 is required for Windows shortcut support"
            ) from exc

        return win32com.client.Dispatch("WScript.Shell")

    def write(self, spec: ShortcutSpec) -> None:
        shortcut = self._shell().CreateShortcut(str(spec.path))
        shortcut.TargetPath = str(spec.target)
        shortcut.Arguments = spec.arguments
        shortcut.WorkingDirectory = str(spec.working_directory)
        shortcut.IconLocation = spec.icon_location
        shortcut.Description = "Launch the local cadence-diff QC Tool"
        shortcut.Save()

    def read(self, path: Path) -> ShortcutSpec | None:
        if not path.exists():
            return None
        shortcut = self._shell().CreateShortcut(str(path))
        return ShortcutSpec(
            path=path,
            target=Path(str(shortcut.TargetPath)),
            arguments=str(shortcut.Arguments),
            working_directory=Path(str(shortcut.WorkingDirectory)),
            icon_location=str(shortcut.IconLocation),
        )

    def remove(self, path: Path) -> None:
        path.unlink(missing_ok=True)


def _windows_desktop_directory() -> Path:
    try:
        from win32com.shell import shell, shellcon  # pyright: ignore[reportMissingModuleSource]
    except ImportError as exc:
        raise ShortcutUnavailableError(
            "pywin32 is required for Windows shortcut support"
        ) from exc

    return Path(
        shell.SHGetFolderPath(
            0,
            shellcon.CSIDL_DESKTOPDIRECTORY,
            0,
            0,
        )
    )


def shortcut_spec(
    data_dir: Path,
    *,
    port: int = 8080,
    desktop_dir: Path | None = None,
    interpreter: Path | None = None,
) -> ShortcutSpec:
    desktop = (desktop_dir or _windows_desktop_directory()).resolve()
    target = (interpreter or quiet_interpreter()).resolve()
    working_directory = data_dir.resolve()
    arguments = subprocess.list2cmdline(
        [
            "-m",
            "qc_tool",
            "launch",
            "--data-dir",
            str(working_directory),
            "--port",
            str(port),
        ]
    )
    icon = Path(str(files("qc_tool").joinpath("assets/qc-tool.ico"))).resolve()
    return ShortcutSpec(
        path=desktop / SHORTCUT_FILENAME,
        target=target,
        arguments=arguments,
        working_directory=working_directory,
        icon_location=f"{icon},0",
    )


def _metadata_path(data_dir: Path) -> Path:
    return data_dir / SHORTCUT_METADATA_FILENAME


def _read_metadata_spec(data_dir: Path) -> ShortcutSpec | None:
    try:
        payload = json.loads(_metadata_path(data_dir).read_text(encoding="utf-8"))
        fields = (
            "path",
            "target",
            "arguments",
            "working_directory",
            "icon_location",
        )
        if payload.get("schema_version") != SHORTCUT_SCHEMA_VERSION or not all(
            isinstance(payload.get(field), str) for field in fields
        ):
            return None
        return ShortcutSpec(
            path=Path(payload["path"]),
            target=Path(payload["target"]),
            arguments=payload["arguments"],
            working_directory=Path(payload["working_directory"]),
            icon_location=payload["icon_location"],
        )
    except (KeyError, OSError, TypeError, ValueError):
        return None


def _metadata_owns_shortcut(data_dir: Path, observed: ShortcutSpec) -> bool:
    recorded = _read_metadata_spec(data_dir)
    return recorded is not None and (
        recorded.path.resolve() == observed.path.resolve()
        and recorded.target.resolve() == observed.target.resolve()
        and recorded.arguments == observed.arguments
        and recorded.working_directory.resolve()
        == observed.working_directory.resolve()
    )


def _write_metadata(data_dir: Path, spec: ShortcutSpec) -> None:
    private_directory(data_dir)
    destination = _metadata_path(data_dir)
    temporary: Path | None = None
    payload = {
        "schema_version": SHORTCUT_SCHEMA_VERSION,
        **{key: str(value) for key, value in asdict(spec).items()},
    }
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=data_dir,
            prefix=".shortcut.",
            suffix=".tmp",
            delete=False,
        ) as handle:
            json.dump(payload, handle, sort_keys=True, separators=(",", ":"))
            temporary = Path(handle.name)
        private_file(temporary)
        os.replace(temporary, destination)
        temporary = None
        private_file(destination)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def _resolve_backend(
    backend: ShortcutBackend | None,
    platform: str,
) -> ShortcutBackend | None:
    if platform != "win32":
        return None
    return backend or WindowsShellShortcutBackend()


def shortcut_status(
    data_dir: Path,
    *,
    port: int = 8080,
    desktop_dir: Path | None = None,
    interpreter: Path | None = None,
    backend: ShortcutBackend | None = None,
    platform: str = sys.platform,
) -> ShortcutResult:
    selected = _resolve_backend(backend, platform)
    if selected is None:
        return ShortcutResult(ShortcutState.UNSUPPORTED, detail="Windows only")
    spec = shortcut_spec(
        data_dir,
        port=port,
        desktop_dir=desktop_dir,
        interpreter=interpreter,
    )
    observed = selected.read(spec.path)
    if observed is None:
        return ShortcutResult(ShortcutState.MISSING, spec)
    if not _metadata_owns_shortcut(data_dir, observed):
        return ShortcutResult(
            ShortcutState.CONFLICT,
            spec,
            "an unowned QC Tool shortcut already exists on this Desktop",
        )
    if (
        not spec.target.exists()
        or observed.target != spec.target
        or observed.arguments != spec.arguments
        or observed.working_directory != spec.working_directory
    ):
        return ShortcutResult(
            ShortcutState.STALE,
            spec,
            "shortcut target no longer matches this Python environment",
        )
    return ShortcutResult(ShortcutState.INSTALLED, spec)


def install_shortcut(
    data_dir: Path,
    *,
    port: int = 8080,
    desktop_dir: Path | None = None,
    interpreter: Path | None = None,
    backend: ShortcutBackend | None = None,
    platform: str = sys.platform,
) -> ShortcutResult:
    selected = _resolve_backend(backend, platform)
    if selected is None:
        return ShortcutResult(ShortcutState.UNSUPPORTED, detail="Windows only")
    spec = shortcut_spec(
        data_dir,
        port=port,
        desktop_dir=desktop_dir,
        interpreter=interpreter,
    )
    if not spec.target.exists():
        return ShortcutResult(
            ShortcutState.STALE,
            spec,
            "current Python environment has no launcher interpreter",
        )
    current = shortcut_status(
        data_dir,
        port=port,
        desktop_dir=desktop_dir,
        interpreter=interpreter,
        backend=selected,
        platform=platform,
    )
    if current.state is ShortcutState.CONFLICT:
        return current
    if current.state is ShortcutState.INSTALLED:
        return current
    spec.path.parent.mkdir(parents=True, exist_ok=True)
    private_directory(spec.working_directory)
    selected.write(spec)
    _write_metadata(data_dir, spec)
    return shortcut_status(
        data_dir,
        port=port,
        desktop_dir=desktop_dir,
        interpreter=interpreter,
        backend=selected,
        platform=platform,
    )


def remove_shortcut(
    data_dir: Path,
    *,
    port: int = 8080,
    desktop_dir: Path | None = None,
    interpreter: Path | None = None,
    backend: ShortcutBackend | None = None,
    platform: str = sys.platform,
) -> ShortcutResult:
    selected = _resolve_backend(backend, platform)
    if selected is None:
        return ShortcutResult(ShortcutState.UNSUPPORTED, detail="Windows only")
    spec = shortcut_spec(
        data_dir,
        port=port,
        desktop_dir=desktop_dir,
        interpreter=interpreter,
    )
    observed = selected.read(spec.path)
    if observed is not None and not _metadata_owns_shortcut(data_dir, observed):
        return ShortcutResult(
            ShortcutState.CONFLICT,
            spec,
            "refusing to remove an unowned Desktop shortcut",
        )
    selected.remove(spec.path)
    _metadata_path(data_dir).unlink(missing_ok=True)
    return ShortcutResult(ShortcutState.MISSING, spec)
