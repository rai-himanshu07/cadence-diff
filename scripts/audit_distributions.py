"""Verify wheel and sdist contain only release-required files."""

from __future__ import annotations

import email.policy
import sys
import tarfile
import zipfile
from email.parser import BytesParser
from pathlib import Path, PurePosixPath

if __package__:
    from scripts.audit_public_tree import (
        _PRIVATE_SUFFIXES,
        _PROHIBITED_PUBLIC_PATTERNS,
        _SECRET_PATTERNS,
    )
else:
    from audit_public_tree import (
        _PRIVATE_SUFFIXES,
        _PROHIBITED_PUBLIC_PATTERNS,
        _SECRET_PATTERNS,
    )

_REQUIRED_PACKAGE_FILES = {
    "qc_tool/__init__.py",
    "qc_tool/assets/qc-tool.ico",
    "qc_tool/assets/review-queue.png",
    "qc_tool/cli.py",
    "qc_tool/ui/guide.py",
    "qc_tool/fingerprint.schema.json",
    "qc_tool/package-fingerprint.schema.json",
    "qc_tool/package-sanitize.schema.json",
    "qc_tool/report/findings-v2.schema.json",
    "qc_tool/report/findings-v3.schema.json",
    "qc_tool/report/findings.schema.json",
    "qc_tool/report/templates/report.html.j2",
}
_PNG_TEXT_CHUNKS = (b"tEXt", b"zTXt", b"iTXt", b"eXIf")
_FORBIDDEN_PARTS = {
    ".github",
    ".nicegui",
    ".vscode",
    "data",
    "docs",
    "scripts",
    "tests",
}
_EXPECTED_VERSION = "2.0.0"
_EXPECTED_NATIVE_REQUIREMENT = "cadence-diff-native==2.0.0"


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


def _metadata_payload(path: Path, *, wheel: bool) -> bytes:
    if wheel:
        with zipfile.ZipFile(path) as archive:
            names = [name for name in archive.namelist() if name.endswith(".dist-info/METADATA")]
            if len(names) != 1:
                raise ValueError("wheel must contain exactly one METADATA file")
            return archive.read(names[0])
    with tarfile.open(path, "r:gz") as archive:
        members = [
            member
            for member in archive.getmembers()
            if member.isfile() and PurePosixPath(member.name).name == "PKG-INFO"
        ]
        if len(members) != 1:
            raise ValueError("sdist must contain exactly one PKG-INFO file")
        handle = archive.extractfile(members[0])
        if handle is None:
            raise ValueError("sdist PKG-INFO is unreadable")
        return handle.read()


def _audit_project_metadata(path: Path, *, wheel: bool) -> list[str]:
    try:
        message = BytesParser(policy=email.policy.default).parsebytes(
            _metadata_payload(path, wheel=wheel)
        )
    except (OSError, ValueError, tarfile.TarError, zipfile.BadZipFile) as exc:
        return [f"{path.name} metadata cannot be read ({type(exc).__name__})"]
    errors: list[str] = []
    if message["Name"] != "cadence-diff":
        errors.append(f"{path.name} has the wrong project name")
    if message["Version"] != _EXPECTED_VERSION:
        errors.append(f"{path.name} has the wrong release version")
    if message["Requires-Python"] != ">=3.11":
        errors.append(f"{path.name} has the wrong Python requirement")
    requirements = set(message.get_all("Requires-Dist", []))
    if _EXPECTED_NATIVE_REQUIREMENT not in requirements:
        errors.append(f"{path.name} misses the exact native helper requirement")
    if any("xlsbkernel" in requirement.casefold() for requirement in requirements):
        errors.append(f"{path.name} retains the legacy native requirement")
    return errors


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


def _audit_binary_metadata(path: Path, names: set[str], *, wheel: bool) -> list[str]:
    errors: list[str] = []
    if wheel:
        with zipfile.ZipFile(path) as archive:
            payloads = {
                name: archive.read(name)
                for name in names
                if name.lower().endswith(".png")
            }
    else:
        with tarfile.open(path, "r:gz") as archive:
            payloads = {}
            for member in archive.getmembers():
                if not member.isfile() or not member.name.lower().endswith(".png"):
                    continue
                handle = archive.extractfile(member)
                if handle is not None:
                    relative = PurePosixPath(member.name)
                    name = str(PurePosixPath(*relative.parts[1:]))
                    payloads[name] = handle.read()
    for name, payload in payloads.items():
        for chunk in _PNG_TEXT_CHUNKS:
            if chunk in payload:
                errors.append(
                    f"{path.name} PNG contains metadata chunk "
                    f"{chunk.decode()}: {name}"
                )
    return errors


def _archive_payloads(path: Path, *, wheel: bool) -> list[tuple[str, bytes]]:
    if wheel:
        with zipfile.ZipFile(path) as archive:
            return [
                (info.filename.rstrip("/"), archive.read(info))
                for info in archive.infolist()
                if not info.is_dir()
            ]
    with tarfile.open(path, "r:gz") as archive:
        payloads: list[tuple[str, bytes]] = []
        for member in archive.getmembers():
            if not member.isfile():
                continue
            pure = PurePosixPath(member.name)
            relative = str(PurePosixPath(*pure.parts[1:]))
            handle = archive.extractfile(member)
            if handle is not None:
                payloads.append((relative, handle.read()))
        return payloads


def _audit_archive_privacy(path: Path, *, wheel: bool) -> list[str]:
    errors: list[str] = []
    for name, payload in _archive_payloads(path, wheel=wheel):
        if PurePosixPath(name).suffix.casefold() in _PRIVATE_SUFFIXES:
            errors.append(f"{path.name} contains private suffix: {name}")
        for label, pattern in (
            *_SECRET_PATTERNS.items(),
            *_PROHIBITED_PUBLIC_PATTERNS.items(),
        ):
            if pattern.search(name):
                errors.append(f"{path.name} member has possible {label}: {name}")
        text = payload.decode("latin-1")
        for label, pattern in (
            *_SECRET_PATTERNS.items(),
            *_PROHIBITED_PUBLIC_PATTERNS.items(),
        ):
            if pattern.search(text):
                errors.append(f"{path.name} contains possible {label}: {name}")
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
        *_audit_binary_metadata(wheels[0], _wheel_names(wheels[0]), wheel=True),
        *_audit_binary_metadata(sdists[0], _sdist_names(sdists[0]), wheel=False),
        *_audit_archive_privacy(wheels[0], wheel=True),
        *_audit_archive_privacy(sdists[0], wheel=False),
        *_audit_project_metadata(wheels[0], wheel=True),
        *_audit_project_metadata(sdists[0], wheel=False),
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
