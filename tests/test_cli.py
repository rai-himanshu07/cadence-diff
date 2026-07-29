"""CLI entry-point tests (no server boot)."""

from pathlib import Path

import pytest

import qc_tool.cli as cli
from qc_tool import __version__


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


def test_main_passes_args_to_run_app(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    calls: list[tuple[Path, int, str, str]] = []

    def fake_run_app(
        work_dir: Path,
        *,
        port: int = 8080,
        host: str = "127.0.0.1",
        network_mode=None,
        expires_at=None,
    ) -> None:
        assert network_mode is not None
        calls.append((work_dir, port, host, network_mode.value))

    import qc_tool.ui.app as app_module

    monkeypatch.setattr(app_module, "run_app", fake_run_app)
    data_dir = tmp_path / "custom"
    cli.main(["--data-dir", str(data_dir), "--port", "9123"])
    assert calls == [(data_dir, 9123, "127.0.0.1", "local")]
    assert data_dir.exists()  # created on demand


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
