"""Verify wheel and sdist contain only release-required files."""

from __future__ import annotations

import sys
import tarfile
import zipfile
from pathlib import Path, PurePosixPath

_REQUIRED_PACKAGE_FILES = {
    "qc_tool/__init__.py",
    "qc_tool/cli.py",
    "qc_tool/ui/guide.py",
    "qc_tool/report/findings.schema.json",
    "qc_tool/report/templates/report.html.j2",
}
_FORBIDDEN_PARTS = {
    ".github",
    ".nicegui",
    ".vscode",
    "data",
    "docs",
    "scripts",
    "tests",
}


def _wheel_names(path: Path) -> set[str]:
    with zipfile.ZipFile(path) as archive:
        return {
            name.rstrip("/")
            for name in archive.namelist()
            if not name.endswith("/")
        }


def _sdist_names(path: Path) -> set[str]:
    with tarfile.open(path, "r:gz") as archive:
        names = {
            member.name.rstrip("/")
            for member in archive.getmembers()
            if member.isfile()
        }
    roots = {PurePosixPath(name).parts[0] for name in names}
    if len(roots) != 1:
        raise ValueError(f"sdist must have one root directory, found {sorted(roots)}")
    root = next(iter(roots))
    return {
        str(PurePosixPath(name).relative_to(root))
        for name in names
    }


def _audit_names(path: Path, names: set[str], *, wheel: bool) -> list[str]:
    errors: list[str] = []
    for name in sorted(names):
        parts = set(PurePosixPath(name).parts)
        if parts & _FORBIDDEN_PARTS:
            errors.append(f"{path.name} contains forbidden path: {name}")
    package_names = {name for name in names if name.startswith("qc_tool/")}
    missing = _REQUIRED_PACKAGE_FILES - package_names
    if missing:
        errors.append(f"{path.name} misses required package files: {sorted(missing)}")
    if wheel:
        unexpected = {
            name
            for name in names
            if not name.startswith("qc_tool/")
            and ".dist-info/" not in name
        }
        if unexpected:
            errors.append(f"{path.name} has unexpected wheel entries: {sorted(unexpected)}")
    else:
        allowed_roots = {
            ".gitignore",
            "LICENSE",
            "PKG-INFO",
            "README.md",
            "pyproject.toml",
            "qc_tool",
        }
        unexpected = {
            name
            for name in names
            if PurePosixPath(name).parts[0] not in allowed_roots
        }
        if unexpected:
            errors.append(f"{path.name} has unexpected sdist entries: {sorted(unexpected)}")
    return errors


def main(argv: list[str]) -> int:
    directory = Path(argv[1] if len(argv) > 1 else "dist")
    wheels = sorted(directory.glob("*.whl"))
    sdists = sorted(directory.glob("*.tar.gz"))
    if len(wheels) != 1 or len(sdists) != 1:
        print(
            f"Expected one wheel and one sdist in {directory}; "
            f"found {len(wheels)} wheel(s), {len(sdists)} sdist(s)"
        )
        return 1
    errors = [
        *_audit_names(wheels[0], _wheel_names(wheels[0]), wheel=True),
        *_audit_names(sdists[0], _sdist_names(sdists[0]), wheel=False),
    ]
    if errors:
        print("Distribution audit failed:")
        for error in errors:
            print(f"- {error}")
        return 1
    print(f"Distribution audit passed: {wheels[0].name}, {sdists[0].name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
