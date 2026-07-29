"""Package-aware strict sanitization and mapping reconciliation."""

import hashlib
import json
import os
from pathlib import Path

from qc_tool import cli
from qc_tool.package_sanitize import sanitize_package
from qc_tool.privacy import verify_sanitized
from tests.conftest import fixture_profile


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_package_sanitize_preserves_privacy_and_mapping_consistency(
    fixture_dir: Path, tmp_path: Path
) -> None:
    excel = fixture_dir / "current.xlsx"
    ppt = fixture_dir / "current.pptx"
    before = {"excel": _sha256(excel), "ppt": _sha256(ppt)}
    output = tmp_path / "sanitized"

    manifest = sanitize_package(
        excel,
        ppt,
        output,
        fixture_profile(),
        seed=7,
        forbidden_tokens=["Long_Monthly", "Executive Summary"],
    )

    assert manifest.privacy_safe
    assert manifest.mappings_verified == 2
    assert manifest.mappings_unverifiable == 0
    assert all(item.status == "verified" for item in manifest.mapping_results)
    workbook = output / "workbook.sanitized.xlsx"
    deck = output / "deck.sanitized.pptx"
    assert verify_sanitized(workbook, forbidden_tokens=["Long_Monthly"]).safe
    assert verify_sanitized(deck, forbidden_tokens=["Executive Summary"]).safe
    manifest_text = (output / "redaction-manifest.json").read_text(encoding="utf-8")
    assert "Long_Monthly" not in manifest_text
    assert "Executive Summary" not in manifest_text
    assert {"excel": _sha256(excel), "ppt": _sha256(ppt)} == before
    if os.name == "posix":
        assert workbook.stat().st_mode & 0o777 == 0o600
        assert output.stat().st_mode & 0o777 == 0o700


def test_package_sanitize_cli(fixture_dir: Path, tmp_path: Path) -> None:
    profile = tmp_path / "profile.yaml"
    profile.write_text(
        json.dumps(fixture_profile().model_dump(mode="json", by_alias=True)),
        encoding="utf-8",
    )
    output = tmp_path / "cli-package"

    code = cli.main(
        [
            "sanitize-package",
            "--excel",
            str(fixture_dir / "current.xlsx"),
            "--ppt",
            str(fixture_dir / "current.pptx"),
            "--profile",
            str(profile),
            "--output-dir",
            str(output),
            "--forbid",
            "Long_Monthly",
        ]
    )

    assert code == 0
    assert (output / "redaction-manifest.json").exists()
