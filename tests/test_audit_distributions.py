"""Distribution binary-metadata audit contracts."""

from __future__ import annotations

import io
import tarfile
import zipfile
from pathlib import Path

from scripts.audit_distributions import _audit_binary_metadata


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
