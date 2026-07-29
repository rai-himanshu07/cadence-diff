"""Local filesystem privacy helpers for QC artifacts."""

import os
from pathlib import Path


def private_directory(path: Path) -> Path:
    """Create a managed directory and restrict it to the current user on POSIX."""
    path.mkdir(parents=True, exist_ok=True, mode=0o700)
    if os.name == "posix":
        path.chmod(0o700)
    return path


def private_file(path: Path) -> Path:
    """Restrict an existing managed file to the current user on POSIX."""
    if os.name == "posix" and path.exists():
        path.chmod(0o600)
    return path


def secure_managed_tree(root: Path) -> Path:
    """Restrict an existing app-managed tree without following symlinks."""
    private_directory(root)
    if os.name != "posix":
        return root
    for path in root.rglob("*"):
        if path.is_symlink():
            continue
        if path.is_dir():
            path.chmod(0o700)
        elif path.is_file():
            path.chmod(0o600)
    return root
