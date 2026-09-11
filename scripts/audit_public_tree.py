"""Fail when the tracked Git tree contains non-public or sensitive material."""

from __future__ import annotations

import re
import subprocess
from pathlib import Path

_ALLOWED_FILES = {
    ".gitattributes",
    ".gitignore",
    ".github/dependabot.yml",
    ".github/workflows/ci.yml",
    ".github/workflows/publish.yml",
    ".github/workflows/xlsbkernel-wheels.yml",
    "CHANGELOG.md",
    "LICENSE",
    "README.md",
    "SECURITY.md",
    "architecture.md",
    "environment.yml",
    "main.py",
    "pyproject.toml",
    "scripts/audit_distributions.py",
    "scripts/audit_dependencies.py",
    "scripts/audit_public_tree.py",
    "scripts/formula_phase_diagnostic.py",
    "scripts/windows_excel_acceptance.py",
    "scripts/windows_focus_acceptance.py",
    "scripts/windows_office_discovery.py",
    "scripts/windows_large_workbook_throughput_acceptance.py",
    "scripts/xlsb_load_locality_diagnostic.py",
}
_ALLOWED_PREFIXES = ("qc_tool/", "tests/", "native/xlsbkernel/")
_PRIVATE_SUFFIXES = {
    ".db",
    ".key",
    ".p12",
    ".pem",
    ".pfx",
    ".qca",
    ".sqlite",
    ".sqlite3",
}
_SECRET_PATTERNS = {
    "private key": re.compile(r"-----BEGIN [A-Z ]*PRIVATE KEY-----"),
    "GitHub token": re.compile(r"\b(?:ghp|github_pat)_[A-Za-z0-9_]{20,}\b"),
    "PyPI token": re.compile(r"\bpypi-[A-Za-z0-9_-]{20,}\b"),
    "AWS access key": re.compile(r"\bAKIA[0-9A-Z]{16}\b"),
}
_MAX_TRACKED_BYTES = 5 * 1024 * 1024
_PNG_TEXT_CHUNKS = (b"tEXt", b"zTXt", b"iTXt", b"eXIf")


def _tracked_files() -> list[str]:
    output = subprocess.check_output(
        ["git", "ls-files", "-z"],
        text=False,
    )
    return [item.decode() for item in output.split(b"\0") if item]


def _is_allowed(path: str) -> bool:
    return path in _ALLOWED_FILES or path.startswith(_ALLOWED_PREFIXES)


def main() -> int:
    errors: list[str] = []
    for relative in _tracked_files():
        path = Path(relative)
        if not _is_allowed(relative):
            errors.append(f"tracked path is outside public allowlist: {relative}")
        if path.suffix.lower() in _PRIVATE_SUFFIXES:
            errors.append(f"private artifact suffix is tracked: {relative}")
        if path.exists() and path.stat().st_size > _MAX_TRACKED_BYTES:
            errors.append(
                f"tracked file exceeds {_MAX_TRACKED_BYTES} bytes: {relative}"
            )
        if not path.is_file():
            continue
        if path.suffix.lower() == ".png":
            payload = path.read_bytes()
            for chunk in _PNG_TEXT_CHUNKS:
                if chunk in payload:
                    errors.append(
                        f"PNG contains metadata chunk {chunk.decode()}: {relative}"
                    )
        try:
            text = path.read_text(encoding="utf-8")
        except (UnicodeDecodeError, OSError):
            continue
        for label, pattern in _SECRET_PATTERNS.items():
            if pattern.search(text):
                errors.append(f"possible {label} in tracked file: {relative}")
    if errors:
        print("Public tree audit failed:")
        for error in sorted(set(errors)):
            print(f"- {error}")
        return 1
    tracked = _tracked_files()
    print(f"Public tree audit passed: {len(tracked)} tracked files")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
