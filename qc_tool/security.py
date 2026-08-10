"""Local filesystem privacy helpers for QC artifacts."""

import os
from pathlib import Path


def _restrict_windows_path(path: Path, *, inherit: bool) -> None:
    import ntsecuritycon  # pyright: ignore[reportMissingModuleSource]
    import win32api  # pyright: ignore[reportMissingModuleSource]
    import win32con  # pyright: ignore[reportMissingModuleSource]
    import win32security  # pyright: ignore[reportMissingModuleSource]

    token = win32security.OpenProcessToken(
        win32api.GetCurrentProcess(),
        win32con.TOKEN_QUERY,
    )
    try:
        user_sid = win32security.GetTokenInformation(
            token,
            win32security.TokenUser,
        )[0]
    finally:
        win32api.CloseHandle(token)
    dacl = win32security.ACL()
    inheritance = (
        win32con.CONTAINER_INHERIT_ACE | win32con.OBJECT_INHERIT_ACE
        if inherit
        else 0
    )
    dacl.AddAccessAllowedAceEx(
        win32security.ACL_REVISION_DS,
        inheritance,
        ntsecuritycon.FILE_ALL_ACCESS,
        user_sid,
    )
    descriptor = win32security.SECURITY_DESCRIPTOR()
    descriptor.SetSecurityDescriptorDacl(True, dacl, False)
    win32security.SetFileSecurity(
        str(path),
        win32security.DACL_SECURITY_INFORMATION,
        descriptor,
    )


def restrict_windows_path_to_current_user(path: Path, *, inherit: bool) -> None:
    """Apply one current-user-only Windows DACL to a file or directory."""
    _restrict_windows_path(path, inherit=inherit)


def private_directory(path: Path) -> Path:
    """Create a managed directory and restrict it to the current user on POSIX."""
    path.mkdir(parents=True, exist_ok=True, mode=0o700)
    if os.name == "posix":
        path.chmod(0o700)
    elif os.name == "nt":
        _restrict_windows_path(path, inherit=True)
    return path


def private_file(path: Path) -> Path:
    """Restrict an existing managed file to the current user on POSIX."""
    if os.name == "posix" and path.exists():
        path.chmod(0o600)
    elif os.name == "nt" and path.exists():
        _restrict_windows_path(path, inherit=False)
    return path


def secure_managed_tree(root: Path) -> Path:
    """Restrict an existing app-managed tree without following symlinks."""
    private_directory(root)
    if os.name == "nt":
        for path in root.rglob("*"):
            if path.is_symlink():
                continue
            _restrict_windows_path(path, inherit=path.is_dir())
        return root
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
