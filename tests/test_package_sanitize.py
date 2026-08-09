"""Package-aware strict sanitization and mapping reconciliation."""

import hashlib
import json
import os
from pathlib import Path

import pytest

from qc_tool import cli
from qc_tool.package import PackageManifest
from qc_tool.package_sanitize import (
    _replace_deck_figures,
    sanitize_package,
    sanitize_package_files,
)
from qc_tool.ppt.extract import load_deck_snapshot
from qc_tool.privacy import verify_sanitized
from tests.conftest import fixture_profile
from tests.fixtures.ppt_builder import build_grouped_text_deck


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
    assert manifest.schema_version == 1
    assert manifest.package_manifest is None
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


def test_grouped_figure_rewrite_uses_extraction_occurrence_order(
    tmp_path: Path,
) -> None:
    deck_path = build_grouped_text_deck(
        tmp_path / "grouped-rewrite.pptx",
        grouped_lines=("Revenue $100M",),
        nested_lines=("Margin 20%",),
        top_after=("Other $200M",),
    )

    _replace_deck_figures(deck_path, {0: "$999M"})

    assert load_deck_snapshot(deck_path).slides[0].texts == [
        "Grouped KPIs",
        "Revenue $999M",
        "Margin 20%",
        "Other $200M",
    ]


def test_multi_member_sanitize_routes_mapping_and_redacts_member_identity(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    excel = fixture_dir / "current.xlsx"
    ppt = fixture_dir / "current.pptx"
    files = {
        "current_excel:client_core": excel,
        "current_excel:client_ops": excel,
        "current_ppt": ppt,
    }
    package_manifest = PackageManifest.from_role_files(files)
    profile = fixture_profile()
    profile.crosscheck.mappings = [
        profile.crosscheck.mappings[0].model_copy(
            update={"source_member": "client_ops"}
        )
    ]
    before = {role: _sha256(path) for role, path in files.items()}
    output = tmp_path / "multi-sanitized"

    result = sanitize_package_files(
        files,
        package_manifest,
        output,
        profile,
        seed=11,
        forbidden_tokens=["Long_Monthly", "Executive Summary"],
    )

    assert result.schema_version == 2
    assert result.privacy_safe
    assert result.mappings_verified == 1
    assert result.mappings_unverifiable == 0
    assert len(result.mapping_results) == 1
    assert result.mapping_results[0].source_member == "excel_002"
    assert result.mapping_results[0].status == "verified"
    assert result.package_manifest is not None
    assert {member.member_id for member in result.package_manifest.members} == {
        "excel_001",
        "excel_002",
        "primary",
    }
    workbooks = sorted(output.glob("workbook-*.sanitized.xlsx"))
    assert [path.name for path in workbooks] == [
        "workbook-excel_001.sanitized.xlsx",
        "workbook-excel_002.sanitized.xlsx",
    ]
    assert all(verify_sanitized(path).safe for path in workbooks)
    assert verify_sanitized(output / "deck-primary.sanitized.pptx").safe
    manifest_text = (output / "redaction-manifest.json").read_text(
        encoding="utf-8"
    )
    for forbidden in (
        "client_core",
        "client_ops",
        str(excel),
        excel.name,
        "Long_Monthly",
        "Executive Summary",
    ):
        assert forbidden not in manifest_text
    assert {role: _sha256(path) for role, path in files.items()} == before


def test_multi_member_sanitize_refuses_baseline_members(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    files = {
        "baseline_excel:old": fixture_dir / "baseline.xlsx",
        "current_excel:core": fixture_dir / "current.xlsx",
        "current_ppt": fixture_dir / "current.pptx",
    }

    with pytest.raises(ValueError, match="current workbook members only"):
        sanitize_package_files(
            files,
            PackageManifest.from_role_files(files),
            tmp_path / "refused",
            fixture_profile(),
        )


def test_package_sanitize_schema_accepts_v1_and_requires_manifest_for_v2() -> None:
    schema_path = Path(__file__).parents[1] / "qc_tool/package-sanitize.schema.json"
    schema = json.loads(schema_path.read_text(encoding="utf-8"))

    assert schema["properties"]["schema_version"] == {"enum": [1, 2]}
    conditional = schema["allOf"][0]
    assert conditional["if"]["properties"]["schema_version"] == {"const": 2}
    assert conditional["then"]["required"] == ["package_manifest"]
