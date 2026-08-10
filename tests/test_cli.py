"""CLI entry-point tests (no server boot)."""

import builtins
import logging
from pathlib import Path

import pytest

import qc_tool.cli as cli
from qc_tool import __version__
from qc_tool.launcher import LaunchOutcome, LaunchResult
from qc_tool.server_config import load_server_config
from qc_tool.shortcut import ShortcutResult, ShortcutState


def test_version_flag(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as excinfo:
        cli.main(["--version"])
    assert excinfo.value.code == 0
    assert __version__ in capsys.readouterr().out


def test_help_mentions_modes(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as excinfo:
        cli.main(["--help"])
    assert excinfo.value.code == 0
    out = capsys.readouterr().out
    assert "preflight" in out and "127.0.0.1" in out


def test_default_data_dir_is_user_scoped() -> None:
    default = cli.default_data_dir()
    assert "qc-tool" in str(default)
    assert not str(default).startswith(str(Path(__file__).parent))


def test_profile_resolution_does_not_import_ui(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    real_import = builtins.__import__

    def guarded_import(name: str, *args, **kwargs):
        if name == "qc_tool.ui.app":
            raise AssertionError("headless profile resolution imported the UI")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", guarded_import)

    assert cli._resolve_profile(None, tmp_path).name == "default"


def test_main_passes_args_to_run_app(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    calls: list[tuple[Path, int, str, str, bool, bool]] = []

    def fake_run_app(
        work_dir: Path,
        *,
        port: int = 8080,
        host: str = "127.0.0.1",
        network_mode=None,
        expires_at=None,
        desktop_focus: bool = False,
        show: bool = True,
    ) -> None:
        assert network_mode is not None
        calls.append((work_dir, port, host, network_mode.value, desktop_focus, show))

    import qc_tool.ui.app as app_module

    monkeypatch.setattr(app_module, "run_app", fake_run_app)
    data_dir = tmp_path / "custom"
    cli.main(["--data-dir", str(data_dir), "--port", "9123"])
    assert calls == [(data_dir, 9123, "127.0.0.1", "local", False, True)]
    assert data_dir.exists()  # created on demand
    calls.clear()
    cli.main(["--data-dir", str(data_dir), "--desktop-focus"])
    assert calls == [(data_dir, 8080, "127.0.0.1", "local", True, True)]
    assert load_server_config(data_dir).desktop_focus is False


def test_launcher_child_suppresses_browser_and_configures_logging(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    observed: dict[str, object] = {}

    def fake_logging(data_dir: Path) -> Path:
        observed["logging"] = data_dir
        return data_dir / "logs" / "server.log"

    def fake_run_app(_work_dir: Path, **kwargs) -> None:
        observed.update(kwargs)

    import qc_tool.launcher as launcher_module
    import qc_tool.ui.app as app_module

    monkeypatch.setattr(launcher_module, "configure_launcher_logging", fake_logging)
    monkeypatch.setattr(app_module, "run_app", fake_run_app)

    assert cli.main(
        [
            "serve",
            "--data-dir",
            str(tmp_path),
            "--no-browser",
            "--launcher-child",
        ]
    ) == 0
    assert observed["logging"] == tmp_path
    assert observed["show"] is False


def test_launcher_child_logs_startup_import_failure_without_paths(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    real_import = builtins.__import__

    def fail_ui_import(name: str, *args, **kwargs):
        if name == "qc_tool.ui.app":
            raise ImportError(f"private path {tmp_path}")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fail_ui_import)

    with pytest.raises(ImportError):
        cli.main(
            [
                "serve",
                "--data-dir",
                str(tmp_path),
                "--no-browser",
                "--launcher-child",
            ]
        )
    root = logging.getLogger()
    log_path = tmp_path / "logs" / "server.log"
    handlers = [
        handler
        for handler in root.handlers
        if getattr(handler, "baseFilename", "") == str(log_path.resolve())
    ]
    for handler in handlers:
        handler.flush()
    content = log_path.read_text(encoding="utf-8")
    assert "server-start-failed ImportError" in content
    assert str(tmp_path) not in content
    for handler in handlers:
        root.removeHandler(handler)
        handler.close()


def test_launcher_child_does_not_log_clean_system_exit(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    def fake_logging(_data_dir: Path) -> Path:
        return tmp_path / "logs" / "server.log"

    def clean_exit(_work_dir: Path, **_kwargs) -> None:
        raise SystemExit(0)

    import qc_tool.launcher as launcher_module
    import qc_tool.ui.app as app_module

    monkeypatch.setattr(launcher_module, "configure_launcher_logging", fake_logging)
    monkeypatch.setattr(app_module, "run_app", clean_exit)

    with (
        caplog.at_level(logging.ERROR, logger="qc_tool.cli"),
        pytest.raises(SystemExit, match="0"),
    ):
        cli.main(
            [
                "serve",
                "--data-dir",
                str(tmp_path),
                "--no-browser",
                "--launcher-child",
            ]
        )

    assert not [
        record for record in caplog.records if "server-start-failed" in record.message
    ]


def test_launch_subcommand_uses_authenticated_launcher(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    calls: list[tuple[Path, int]] = []

    def fake_launch(data_dir: Path, *, port: int) -> LaunchResult:
        calls.append((data_dir, port))
        return LaunchResult(LaunchOutcome.REUSED, f"http://127.0.0.1:{port}/")

    import qc_tool.launcher as launcher_module

    monkeypatch.setattr(launcher_module, "launch_local_app", fake_launch)
    monkeypatch.setattr(
        launcher_module,
        "configure_launcher_logging",
        lambda data_dir: data_dir / "logs" / "server.log",
    )

    assert cli.main(
        ["launch", "--data-dir", str(tmp_path), "--port", "9123"]
    ) == 0
    assert calls == [(tmp_path, 9123)]
    assert "reused" in capsys.readouterr().out


@pytest.mark.parametrize(
    ("action", "state"),
    [
        ("install", ShortcutState.INSTALLED),
        ("status", ShortcutState.INSTALLED),
        ("remove", ShortcutState.MISSING),
    ],
)
def test_shortcut_subcommand_dispatches_shared_backend(
    action: str,
    state: ShortcutState,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
) -> None:
    calls: list[tuple[str, Path, int]] = []

    def operation(name: str):
        def run(data_dir: Path, *, port: int) -> ShortcutResult:
            calls.append((name, data_dir, port))
            return ShortcutResult(state)

        return run

    import qc_tool.shortcut as shortcut_module

    monkeypatch.setattr(shortcut_module, "install_shortcut", operation("install"))
    monkeypatch.setattr(shortcut_module, "shortcut_status", operation("status"))
    monkeypatch.setattr(shortcut_module, "remove_shortcut", operation("remove"))

    assert cli.main(
        ["shortcut", action, "--data-dir", str(tmp_path), "--port", "9123"]
    ) == 0
    assert calls == [(action, tmp_path, 9123)]
    assert state.value in capsys.readouterr().out


def test_serve_uses_temporary_lan_override(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    calls: list[tuple[str, str]] = []

    def fake_run_app(work_dir: Path, **kwargs) -> None:
        calls.append((kwargs["host"], kwargs["network_mode"].value))

    import qc_tool.ui.app as app_module

    monkeypatch.setattr(app_module, "run_app", fake_run_app)
    assert (
        cli.main(
            [
                "--data-dir",
                str(tmp_path),
                "--network",
                "lan",
                "--expose-for",
                "10",
            ]
        )
        == 0
    )
    assert calls == [("0.0.0.0", "lan")]
