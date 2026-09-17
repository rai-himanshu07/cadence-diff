"""Filesystem privacy helper contracts."""

from pathlib import Path

import pytest

import qc_tool.security as security_module


def test_windows_tree_ignores_a_child_that_disappears_during_acl_update(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    transient = tmp_path / "history.sqlite3-shm"
    transient.write_bytes(b"temporary")

    def restrict(path: Path, *, inherit: bool) -> None:
        del inherit
        if path == transient:
            transient.unlink()
            raise FileNotFoundError(2, "file disappeared")

    monkeypatch.setattr(security_module, "_restrict_windows_path", restrict)

    security_module._secure_windows_tree(tmp_path)

    assert not transient.exists()


def test_windows_tree_propagates_non_missing_acl_errors(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    protected = tmp_path / "history.sqlite3"
    protected.write_bytes(b"database")

    def restrict(_path: Path, *, inherit: bool) -> None:
        del inherit
        raise PermissionError(13, "access denied")

    monkeypatch.setattr(security_module, "_restrict_windows_path", restrict)

    with pytest.raises(PermissionError, match="access denied"):
        security_module._secure_windows_tree(tmp_path)
