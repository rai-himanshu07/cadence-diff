"""Windows desktop-shortcut lifecycle contracts with a fake shell backend."""

from __future__ import annotations

from pathlib import Path

from qc_tool.shortcut import (
    ShortcutSpec,
    ShortcutState,
    install_shortcut,
    remove_shortcut,
    shortcut_spec,
    shortcut_status,
)


class _FakeBackend:
    def __init__(self) -> None:
        self.specs: dict[Path, ShortcutSpec] = {}

    def write(self, spec: ShortcutSpec) -> None:
        self.specs[spec.path] = spec
        spec.path.write_bytes(b"fake-lnk")

    def read(self, path: Path) -> ShortcutSpec | None:
        return self.specs.get(path)

    def remove(self, path: Path) -> None:
        self.specs.pop(path, None)
        path.unlink(missing_ok=True)


def test_shortcut_targets_absolute_environment_without_path_lookup(
    tmp_path: Path,
) -> None:
    desktop = tmp_path / "Desktop"
    desktop.mkdir()
    interpreter = tmp_path / "Miniforge Env" / "pythonw.exe"
    interpreter.parent.mkdir()
    interpreter.write_bytes(b"python")
    data_dir = tmp_path / "QC Data"

    spec = shortcut_spec(
        data_dir,
        desktop_dir=desktop,
        interpreter=interpreter,
        port=9123,
    )

    assert spec.path == desktop / "QC Tool.lnk"
    assert spec.target == interpreter.resolve()
    assert spec.target.is_absolute()
    assert "-m qc_tool launch" in spec.arguments
    assert '"' in spec.arguments and str(data_dir) in spec.arguments
    assert spec.working_directory == data_dir.resolve()
    assert "desktop-focus" not in spec.arguments
    assert spec.icon_location.endswith("qc_tool/assets/qc-tool.ico,0")


def test_shortcut_install_status_stale_and_remove_are_idempotent(
    tmp_path: Path,
) -> None:
    desktop = tmp_path / "Desktop"
    desktop.mkdir()
    interpreter = tmp_path / "pythonw.exe"
    interpreter.write_bytes(b"python")
    data_dir = tmp_path / "data"
    backend = _FakeBackend()
    options = {
        "desktop_dir": desktop,
        "interpreter": interpreter,
        "backend": backend,
        "platform": "win32",
    }

    installed = install_shortcut(data_dir, **options)
    assert installed.state is ShortcutState.INSTALLED
    assert shortcut_status(data_dir, **options).state is ShortcutState.INSTALLED

    interpreter.unlink()
    assert shortcut_status(data_dir, **options).state is ShortcutState.STALE

    assert remove_shortcut(data_dir, **options).state is ShortcutState.MISSING
    assert remove_shortcut(data_dir, **options).state is ShortcutState.MISSING


def test_unowned_shortcut_is_never_overwritten_or_removed(tmp_path: Path) -> None:
    desktop = tmp_path / "Desktop"
    desktop.mkdir()
    interpreter = tmp_path / "pythonw.exe"
    interpreter.write_bytes(b"python")
    data_dir = tmp_path / "data"
    backend = _FakeBackend()
    existing = ShortcutSpec(
        path=desktop / "QC Tool.lnk",
        target=tmp_path / "other.exe",
        arguments="--unrelated",
        working_directory=tmp_path,
        icon_location="",
    )
    backend.write(existing)
    options = {
        "desktop_dir": desktop,
        "interpreter": interpreter,
        "backend": backend,
        "platform": "win32",
    }

    assert shortcut_status(data_dir, **options).state is ShortcutState.CONFLICT
    assert install_shortcut(data_dir, **options).state is ShortcutState.CONFLICT
    assert remove_shortcut(data_dir, **options).state is ShortcutState.CONFLICT
    assert backend.read(existing.path) == existing


def test_replaced_shortcut_is_no_longer_treated_as_owned(tmp_path: Path) -> None:
    desktop = tmp_path / "Desktop"
    desktop.mkdir()
    interpreter = tmp_path / "pythonw.exe"
    interpreter.write_bytes(b"python")
    data_dir = tmp_path / "data"
    backend = _FakeBackend()
    options = {
        "desktop_dir": desktop,
        "interpreter": interpreter,
        "backend": backend,
        "platform": "win32",
    }
    installed = install_shortcut(data_dir, **options)
    assert installed.spec is not None
    replacement = ShortcutSpec(
        path=installed.spec.path,
        target=tmp_path / "other.exe",
        arguments="--unrelated",
        working_directory=tmp_path,
        icon_location="",
    )
    backend.write(replacement)

    assert shortcut_status(data_dir, **options).state is ShortcutState.CONFLICT
    assert install_shortcut(data_dir, **options).state is ShortcutState.CONFLICT
    assert remove_shortcut(data_dir, **options).state is ShortcutState.CONFLICT
    assert backend.read(replacement.path) == replacement
