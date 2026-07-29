"""Fail-safe persisted local/LAN server exposure configuration."""

import datetime as dt
import os
from pathlib import Path

import pytest

from qc_tool import cli
from qc_tool.server_config import (
    NetworkMode,
    ServerConfig,
    config_path,
    lan_config_matches,
    load_server_config,
    save_server_config,
    temporary_lan_config,
)


def test_network_defaults_local_and_lan_requires_bounded_expiry(tmp_path: Path) -> None:
    assert load_server_config(tmp_path).network is NetworkMode.LOCAL
    with pytest.raises(ValueError, match="between 1 and 1440"):
        temporary_lan_config(0)

    config = temporary_lan_config(30)
    save_server_config(tmp_path, config)
    loaded = load_server_config(tmp_path)
    assert loaded.network is NetworkMode.LAN
    assert loaded.host == "0.0.0.0"
    assert loaded.expires_at is not None
    assert lan_config_matches(tmp_path, loaded.expires_at)
    if os.name == "posix":
        assert config_path(tmp_path).stat().st_mode & 0o777 == 0o600


def test_expired_lan_config_falls_back_and_persists_local(tmp_path: Path) -> None:
    expired = ServerConfig(
        network=NetworkMode.LAN,
        expires_at=dt.datetime.now(dt.UTC) - dt.timedelta(seconds=1),
    )
    save_server_config(tmp_path, expired)

    loaded = load_server_config(tmp_path)

    assert loaded.network is NetworkMode.LOCAL
    persisted = ServerConfig.model_validate_json(
        config_path(tmp_path).read_text(encoding="utf-8")
    )
    assert persisted.network is NetworkMode.LOCAL
    assert expired.expires_at is not None
    assert not lan_config_matches(tmp_path, expired.expires_at)


def test_network_cli_switches_config(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    assert (
        cli.main(
            ["network", "lan", "--minutes", "15", "--data-dir", str(tmp_path)]
        )
        == 0
    )
    assert "network: lan" in capsys.readouterr().out
    assert load_server_config(tmp_path).network is NetworkMode.LAN

    assert cli.main(["network", "local", "--data-dir", str(tmp_path)]) == 0
    assert load_server_config(tmp_path).network is NetworkMode.LOCAL
    assert not lan_config_matches(tmp_path, dt.datetime.now(dt.UTC))
