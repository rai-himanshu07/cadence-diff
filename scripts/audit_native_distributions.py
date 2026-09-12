"""Audit the three wheel-only cadence-diff-native release artifacts."""

from __future__ import annotations

import email.policy
import re
import sys
import zipfile
from email.parser import BytesParser
from pathlib import Path, PurePosixPath

if __package__:
    from scripts.audit_distributions import _FORBIDDEN_PARTS
    from scripts.audit_public_tree import (
        _PRIVATE_SUFFIXES,
        _PROHIBITED_PUBLIC_PATTERNS,
        _SECRET_PATTERNS,
    )
else:
    from audit_distributions import _FORBIDDEN_PARTS
    from audit_public_tree import (
        _PRIVATE_SUFFIXES,
        _PROHIBITED_PUBLIC_PATTERNS,
        _SECRET_PATTERNS,
    )

EXPECTED_NAME = "cadence-diff-native"
EXPECTED_VERSION = "2.0.0"
EXPECTED_REQUIRES_PYTHON = ">=3.11"
EXPECTED_SUMMARY = "Native Excel analysis accelerator for cadence-diff"
EXPECTED_LICENSE = "Apache-2.0"
EXPECTED_CLASSIFIERS = frozenset(
    {
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Rust",
        "Operating System :: Microsoft :: Windows :: Windows 10",
        "Operating System :: POSIX :: Linux",
    }
)
EXPECTED_PROJECT_URLS = frozenset(
    {
        "Homepage, https://github.com/rai-himanshu07/cadence-diff",
        "Repository, https://github.com/rai-himanshu07/cadence-diff",
    }
)
FALLBACK_TAG = "py3-none-any"
WINDOWS_TAG = "cp311-abi3-win_amd64"
LINUX_TAG_PATTERN = re.compile(
    r"^cp311-abi3-(?:manylinux2014_x86_64|"
    r"manylinux_2_17_x86_64(?:\.manylinux2014_x86_64)?)$"
)
MAX_WHEEL_BYTES = 12 * 1024 * 1024
MAX_MEMBER_BYTES = 10 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 24 * 1024 * 1024


def _wheel_tag_from_filename(path: Path) -> str | None:
    prefix = f"cadence_diff_native-{EXPECTED_VERSION}-"
    if not path.name.startswith(prefix) or not path.name.endswith(".whl"):
        return None
    return path.name[len(prefix) : -len(".whl")]


def _metadata(archive: zipfile.ZipFile, wheel_name: str) -> tuple[object, ...]:
    expected = f"cadence_diff_native-{EXPECTED_VERSION}.dist-info/METADATA"
    if expected not in archive.namelist():
        raise ValueError(f"{wheel_name} misses METADATA")
    message = BytesParser(policy=email.policy.default).parsebytes(archive.read(expected))
    return (
        message["Name"],
        message["Version"],
        message["Requires-Python"],
        message["Summary"],
        message["License-Expression"],
        frozenset(message.get_all("Classifier", [])),
        frozenset(message.get_all("Project-URL", [])),
    )


def _wheel_metadata(
    archive: zipfile.ZipFile,
    wheel_name: str,
) -> tuple[set[str], bool | None]:
    expected = f"cadence_diff_native-{EXPECTED_VERSION}.dist-info/WHEEL"
    if expected not in archive.namelist():
        raise ValueError(f"{wheel_name} misses WHEEL metadata")
    message = BytesParser(policy=email.policy.default).parsebytes(archive.read(expected))
    pure_value = message["Root-Is-Purelib"]
    pure = None if pure_value is None else pure_value.casefold() == "true"
    return set(message.get_all("Tag", [])), pure


def _scan_text(errors: list[str], label: str, payload: bytes, location: str) -> None:
    text = payload.decode("latin-1")
    for pattern_label, pattern in (
        *_SECRET_PATTERNS.items(),
        *_PROHIBITED_PUBLIC_PATTERNS.items(),
    ):
        if pattern.search(text):
            errors.append(f"{label} contains possible {pattern_label}: {location}")


def _audit_wheel(
    path: Path,
) -> tuple[list[str], str | None, tuple[object, ...] | None]:
    errors: list[str] = []
    filename_tag = _wheel_tag_from_filename(path)
    if filename_tag is None:
        return [f"unexpected wheel filename: {path.name}"], None, None
    if path.stat().st_size > MAX_WHEEL_BYTES:
        errors.append(f"wheel exceeds size cap: {path.name}")
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            total_size = sum(info.file_size for info in infos)
            if total_size > MAX_UNCOMPRESSED_BYTES:
                errors.append(f"wheel exceeds uncompressed size cap: {path.name}")
            names = {info.filename for info in infos if not info.is_dir()}
            has_extension = any(
                name.startswith("cadence_diff_native/")
                and name.casefold().endswith((".so", ".pyd"))
                for name in names
            )
            has_init = "cadence_diff_native/__init__.py" in names
            if filename_tag == FALLBACK_TAG:
                if not has_init:
                    errors.append(f"fallback misses package init: {path.name}")
                if has_extension:
                    errors.append(f"fallback contains extension: {path.name}")
                if has_init:
                    init_text = archive.read("cadence_diff_native/__init__.py").decode(
                        "utf-8", errors="replace"
                    )
                    required = (
                        '__version__ = "2.0.0"',
                        "__kernel_api_version__ = 1",
                        "__native_available__ = False",
                    )
                    if any(item not in init_text for item in required):
                        errors.append(f"fallback status contract missing: {path.name}")
            else:
                if not has_extension:
                    errors.append(f"native wheel misses extension: {path.name}")
            for info in infos:
                name = info.filename.rstrip("/")
                if not name:
                    continue
                pure = PurePosixPath(name)
                parts = pure.parts
                if pure.is_absolute() or ".." in parts or "\\" in name:
                    errors.append(f"wheel contains unsafe path: {path.name}:{name}")
                if set(parts) & _FORBIDDEN_PARTS:
                    errors.append(f"wheel contains forbidden path: {path.name}:{name}")
                if pure.suffix.casefold() in _PRIVATE_SUFFIXES:
                    errors.append(f"wheel contains private suffix: {path.name}:{name}")
                if parts and parts[0] != "cadence_diff_native" and not parts[0].startswith(
                    f"cadence_diff_native-{EXPECTED_VERSION}.dist-info"
                ):
                    errors.append(f"wheel contains unexpected root: {path.name}:{name}")
                for pattern_label, pattern in (
                    *_SECRET_PATTERNS.items(),
                    *_PROHIBITED_PUBLIC_PATTERNS.items(),
                ):
                    if pattern.search(name):
                        errors.append(
                            f"wheel member has possible {pattern_label}: {path.name}:{name}"
                        )
                if info.file_size > MAX_MEMBER_BYTES:
                    errors.append(f"wheel member exceeds size cap: {path.name}:{name}")
                    continue
                if not info.is_dir() and info.file_size:
                    _scan_text(errors, path.name, archive.read(info), name)
            internal_tags, root_is_pure = _wheel_metadata(archive, path.name)
            if internal_tags != {filename_tag}:
                errors.append(f"WHEEL tag does not match filename: {path.name}")
            expected_pure = filename_tag == FALLBACK_TAG
            if root_is_pure is not expected_pure:
                errors.append(f"WHEEL purity does not match payload: {path.name}")
            metadata = _metadata(archive, path.name)
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        errors.append(f"cannot audit wheel ({type(exc).__name__}): {path.name}")
        return errors, filename_tag, None
    return errors, filename_tag, metadata


def main(argv: list[str] | None = None) -> int:
    args = sys.argv if argv is None else argv
    if len(args) != 2:
        print("usage: audit_native_distributions.py <artifact-dir>")
        return 2
    directory = Path(args[1])
    if not directory.is_dir():
        print("Native distribution audit failed: artifact directory not found")
        return 2
    wheels = sorted(directory.rglob("*.whl"))
    sdists = sorted(
        path
        for path in directory.rglob("*")
        if path.is_file() and path.name.casefold().endswith((".tar.gz", ".zip"))
    )
    errors = [f"helper sdist is forbidden: {path.name}" for path in sdists]
    if len(wheels) != 3:
        errors.append(f"expected exactly three wheels, found {len(wheels)}")
    tags: set[str] = set()
    metadatas: list[tuple[object, ...]] = []
    for wheel in wheels:
        wheel_errors, tag, metadata = _audit_wheel(wheel)
        errors.extend(wheel_errors)
        if tag is not None:
            tags.add(tag)
        if metadata is not None:
            metadatas.append(metadata)
    if FALLBACK_TAG not in tags:
        errors.append("missing universal fallback wheel")
    if WINDOWS_TAG not in tags:
        errors.append("missing Windows native wheel")
    linux_tags = {tag for tag in tags if LINUX_TAG_PATTERN.fullmatch(tag)}
    if len(linux_tags) != 1:
        errors.append("expected one manylinux2014-compatible native wheel")
    expected_metadata = (
        EXPECTED_NAME,
        EXPECTED_VERSION,
        EXPECTED_REQUIRES_PYTHON,
        EXPECTED_SUMMARY,
        EXPECTED_LICENSE,
        EXPECTED_CLASSIFIERS,
        EXPECTED_PROJECT_URLS,
    )
    if any(metadata != expected_metadata for metadata in metadatas):
        errors.append("helper wheel METADATA does not match release contract")
    if len(set(metadatas)) > 1:
        errors.append("helper wheel METADATA differs between artifacts")
    if errors:
        print("Native distribution audit failed:")
        for error in sorted(set(errors))[:50]:
            print(f"- {error}")
        return 1
    print("Native distribution audit passed: three wheel-only artifacts")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
