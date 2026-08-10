"""Fail-safe persisted server exposure configuration."""

import datetime as dt
import logging
from enum import StrEnum
from pathlib import Path

from pydantic import BaseModel

from qc_tool.security import private_directory, private_file

logger = logging.getLogger(__name__)


class NetworkMode(StrEnum):
    LOCAL = "local"
    LAN = "lan"


class ServerConfig(BaseModel):
    network: NetworkMode = NetworkMode.LOCAL
    expires_at: dt.datetime | None = None
    desktop_focus: bool = False

    @property
    def host(self) -> str:
        return "0.0.0.0" if self.network is NetworkMode.LAN else "127.0.0.1"

    def is_expired(self, *, now: dt.datetime | None = None) -> bool:
        if self.network is not NetworkMode.LAN:
            return False
        current = now or dt.datetime.now(dt.UTC)
        return self.expires_at is None or self.expires_at <= current


def config_path(data_dir: Path) -> Path:
    return data_dir / "server-config.json"


def local_config(*, desktop_focus: bool = False) -> ServerConfig:
    return ServerConfig(
        network=NetworkMode.LOCAL,
        desktop_focus=desktop_focus,
    )


def temporary_lan_config(
    minutes: int,
    *,
    desktop_focus: bool = False,
) -> ServerConfig:
    if minutes < 1 or minutes > 24 * 60:
        raise ValueError("LAN exposure duration must be between 1 and 1440 minutes")
    return ServerConfig(
        network=NetworkMode.LAN,
        expires_at=dt.datetime.now(dt.UTC) + dt.timedelta(minutes=minutes),
        desktop_focus=desktop_focus,
    )


def save_server_config(data_dir: Path, config: ServerConfig) -> Path:
    private_directory(data_dir)
    path = config_path(data_dir)
    path.write_text(config.model_dump_json(indent=2), encoding="utf-8")
    private_file(path)
    return path


def load_server_config(
    data_dir: Path, *, now: dt.datetime | None = None
) -> ServerConfig:
    path = config_path(data_dir)
    if not path.exists():
        return local_config()
    try:
        config = ServerConfig.model_validate_json(path.read_text(encoding="utf-8"))
    except Exception as exc:
        logger.error("invalid server config %s; falling back to local: %s", path, exc)
        config = local_config()
        save_server_config(data_dir, config)
        return config
    if config.is_expired(now=now):
        config = local_config(desktop_focus=config.desktop_focus)
        save_server_config(data_dir, config)
    return config


def lan_config_matches(data_dir: Path, expected_expiry: dt.datetime) -> bool:
    """Whether a running LAN server is still authorized by persisted config."""
    config = load_server_config(data_dir)
    return (
        config.network is NetworkMode.LAN
        and config.expires_at == expected_expiry
        and not config.is_expired()
    )
