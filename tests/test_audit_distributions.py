"""Distribution binary-metadata audit contracts."""

from __future__ import annotations

import io
import subprocess
import sys
import tarfile
import zipfile
from pathlib import Path

import pytest

from scripts.audit_distributions import (
    _audit_archive_privacy,
    _audit_binary_metadata,
    _audit_project_metadata,
)


def _metadata(*, native_requirement: str | None = "cadence-diff-native==2.0.0") -> bytes:
    lines = [
        "Metadata-Version: 2.4",
        "Name: cadence-diff",
        "Version: 2.0.0",
        "Requires-Python: >=3.11",
    ]
    if native_requirement is not None:
        lines.append(f"Requires-Dist: {native_requirement}")
    return ("\n".join(lines) + "\n\n").encode()


def _write_distribution_metadata(
    path: Path,
    *,
    wheel: bool,
    payload: bytes,
) -> None:
    if wheel:
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("cadence_diff-2.0.0.dist-info/METADATA", payload)
        return
    with tarfile.open(path, "w:gz") as archive:
        member = tarfile.TarInfo("cadence_diff-2.0.0/PKG-INFO")
        member.size = len(payload)
        archive.addfile(member, io.BytesIO(payload))


def _write_archive_member(
    path: Path,
    *,
    wheel: bool,
    member_name: str,
    payload: bytes,
) -> None:
    if wheel:
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr(member_name, payload)
        return
    rooted_name = f"cadence_diff-2.0.0/{member_name}"
    with tarfile.open(path, "w:gz") as archive:
        member = tarfile.TarInfo(rooted_name)
        member.size = len(payload)
        archive.addfile(member, io.BytesIO(payload))


def test_main_wheel_and_sdist_metadata_require_the_exact_native_helper(
    tmp_path: Path,
) -> None:
    wheel = tmp_path / "cadence_diff-2.0.0-py3-none-any.whl"
    sdist = tmp_path / "cadence_diff-2.0.0.tar.gz"
    _write_distribution_metadata(wheel, wheel=True, payload=_metadata())
    _write_distribution_metadata(sdist, wheel=False, payload=_metadata())

    assert _audit_project_metadata(wheel, wheel=True) == []
    assert _audit_project_metadata(sdist, wheel=False) == []


def test_main_metadata_rejects_a_missing_or_legacy_native_requirement(
    tmp_path: Path,
) -> None:
    missing = tmp_path / "missing.whl"
    legacy = tmp_path / "legacy.whl"
    _write_distribution_metadata(
        missing,
        wheel=True,
        payload=_metadata(native_requirement=None),
    )
    _write_distribution_metadata(
        legacy,
        wheel=True,
        payload=_metadata(native_requirement="xlsbkernel==0.1.0"),
    )

    assert any(
        "misses the exact native helper" in error
        for error in _audit_project_metadata(missing, wheel=True)
    )
    legacy_errors = _audit_project_metadata(legacy, wheel=True)
    assert any("misses the exact native helper" in error for error in legacy_errors)
    assert any("legacy native requirement" in error for error in legacy_errors)


def test_main_auditor_runs_as_a_direct_script(tmp_path: Path) -> None:
    completed = subprocess.run(
        [sys.executable, "scripts/audit_distributions.py", str(tmp_path)],
        cwd=Path(__file__).parents[1],
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 1
    assert "Expected one wheel and one sdist" in completed.stdout
    assert completed.stderr == ""


@pytest.mark.parametrize("wheel", [True, False])
@pytest.mark.parametrize(
    ("member_name", "payload", "message"),
    [
        ("qc_tool/private.sqlite3", b"content", "private suffix"),
        (
            "qc_tool/safe.txt",
            b"-----BEGIN " + b"PRIVATE KEY-----\nsynthetic",
            "possible private key",
        ),
        (
            f"qc_tool/{bytes.fromhex('54584f').decode('ascii')}.txt",
            b"content",
            "possible private workload codename",
        ),
    ],
)
def test_main_archives_reject_private_names_and_content(
    tmp_path: Path,
    wheel: bool,
    member_name: str,
    payload: bytes,
    message: str,
) -> None:
    path = tmp_path / ("example.whl" if wheel else "example.tar.gz")
    _write_archive_member(
        path,
        wheel=wheel,
        member_name=member_name,
        payload=payload,
    )

    assert any(
        message in error for error in _audit_archive_privacy(path, wheel=wheel)
    )


def test_wheel_png_metadata_is_rejected(tmp_path: Path) -> None:
    wheel = tmp_path / "example.whl"
    name = "qc_tool/assets/review-queue.png"
    with zipfile.ZipFile(wheel, "w") as archive:
        archive.writestr(name, b"\x89PNG\r\n\x1a\ntEXtprivate-host")

    errors = _audit_binary_metadata(wheel, {name}, wheel=True)

    assert errors and "tEXt" in errors[0]


def test_sdist_clean_png_passes_metadata_audit(tmp_path: Path) -> None:
    sdist = tmp_path / "example.tar.gz"
    payload = b"\x89PNG\r\n\x1a\nclean-image-data"
    with tarfile.open(sdist, "w:gz") as archive:
        member = tarfile.TarInfo("example/qc_tool/assets/review-queue.png")
        member.size = len(payload)
        archive.addfile(member, io.BytesIO(payload))

    errors = _audit_binary_metadata(
        sdist,
        {"qc_tool/assets/review-queue.png"},
        wheel=False,
    )

    assert errors == []


def test_sdist_png_metadata_is_rejected(tmp_path: Path) -> None:
    sdist = tmp_path / "example.tar.gz"
    payload = b"\x89PNG\r\n\x1a\niTXtprivate-host"
    with tarfile.open(sdist, "w:gz") as archive:
        member = tarfile.TarInfo("example/qc_tool/assets/review-queue.png")
        member.size = len(payload)
        archive.addfile(member, io.BytesIO(payload))

    errors = _audit_binary_metadata(
        sdist,
        {"qc_tool/assets/review-queue.png"},
        wheel=False,
    )

    assert errors and "iTXt" in errors[0]
