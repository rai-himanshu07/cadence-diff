"""Structural fingerprint and fail-closed privacy CLI contracts."""

import json
from pathlib import Path

from qc_tool import cli
from qc_tool.fingerprint import deck_fingerprint, fingerprint_file, package_fingerprint
from qc_tool.package import PackageManifest
from tests.fixtures.ppt_builder import build_grouped_text_deck


def test_excel_fingerprint_is_deterministic_and_contains_no_source_content(
    fixture_dir: Path,
) -> None:
    source = fixture_dir / "current.xlsx"
    first = fingerprint_file(source)
    second = fingerprint_file(source)
    serialized = json.dumps(first, sort_keys=True)

    assert first == second
    assert first["artifact"] == "excel"
    assert first["sheet_count"] > 0
    assert first["formula_patterns"]
    for forbidden in (
        "Long_Monthly",
        "Dashboard",
        "RevenuePivot",
        "KPI_Margin",
        "117784",
        "=C7-D7",
    ):
        assert forbidden not in serialized


def test_ppt_fingerprint_contains_no_visible_text(fixture_dir: Path) -> None:
    payload = fingerprint_file(fixture_dir / "current.pptx")
    serialized = json.dumps(payload, sort_keys=True)

    assert payload["artifact"] == "powerpoint"
    assert payload["slide_count"] > 0
    assert payload["capabilities"] == {"charts": True, "notes": True}
    chart = next(
        chart
        for slide in payload["slides"]
        for chart in slide["charts"]
    )
    assert chart["plot_types"]
    assert chart["plot_axis_groups"]
    assert chart["axis_count"] == len(chart["axis_types"])
    assert isinstance(chart["visible_label_count"], int)
    assert all("shape_types" in slide for slide in payload["slides"])
    assert all("has_notes" in slide for slide in payload["slides"])
    assert "Executive Summary" not in serialized
    assert "Revenue by Region" not in serialized
    assert "$3.12M" not in serialized


def test_ppt_fingerprint_counts_grouped_text_without_leaking_it(
    tmp_path: Path,
) -> None:
    path = build_grouped_text_deck(
        tmp_path / "grouped-fingerprint.pptx",
        grouped_lines=("Secret revenue $120M",),
        nested_lines=("Secret margin 41.5%",),
    )

    payload = deck_fingerprint(path)
    serialized = json.dumps(payload, sort_keys=True)
    slide = payload["slides"][0]

    assert slide["text_block_count"] == 3
    assert slide["figure_count"] == 2
    assert "Secret revenue" not in serialized
    assert "Secret margin" not in serialized


def test_fingerprint_and_privacy_cli(
    fixture_dir: Path, tmp_path: Path
) -> None:
    fingerprint = tmp_path / "fingerprint.json"
    assert (
        cli.main(
            [
                "fingerprint",
                str(fixture_dir / "current.xlsx"),
                "--output",
                str(fingerprint),
            ]
        )
        == 0
    )
    assert json.loads(fingerprint.read_text(encoding="utf-8"))["schema_version"] == 1

    strict = tmp_path / "strict.xlsx"
    assert (
        cli.main(
            [
                "sanitize",
                str(fixture_dir / "current.xlsx"),
                "--output",
                str(strict),
                "--redact-text",
            ]
        )
        == 0
    )
    assert cli.main(["verify-sanitized", str(strict)]) == 0
    assert (
        cli.main(
            [
                "verify-sanitized",
                str(fixture_dir / "current.xlsx"),
                "--forbid",
                "Long_Monthly",
            ]
        )
        == 2
    )


def test_package_fingerprint_v2_preserves_pairing_without_identifiers(
    fixture_dir: Path,
) -> None:
    files = {
        "baseline_excel:client_ops": fixture_dir / "baseline.xlsx",
        "current_excel:client_ops": fixture_dir / "current.xlsx",
        "current_ppt": fixture_dir / "current.pptx",
    }
    manifest = PackageManifest.from_role_files(files)

    first = package_fingerprint(files, manifest)
    second = package_fingerprint(files, manifest)
    serialized = json.dumps(first, sort_keys=True)

    assert first == second
    assert first["schema_version"] == 2
    assert first["artifact"] == "package"
    assert first["schema"].endswith("package-fingerprint-v2.json")
    assert [member["role"] for member in first["members"]] == [
        "baseline_excel:excel_001",
        "current_excel:excel_001",
        "current_ppt",
    ]
    redacted_members = first["package_manifest"]["members"]
    assert [member["member_id"] for member in redacted_members] == [
        "excel_001",
        "excel_001",
        "primary",
    ]
    for forbidden in (
        "client_ops",
        "baseline.xlsx",
        "current.xlsx",
        "current.pptx",
        str(fixture_dir),
        "Long_Monthly",
        "Executive Summary",
    ):
        assert forbidden not in serialized


def test_package_fingerprint_cli_writes_v2_payload(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    output = tmp_path / "package-fingerprint.json"

    code = cli.main(
        [
            "fingerprint",
            "--current-workbook",
            f"client_ops={fixture_dir / 'current.xlsx'}",
            "--current-ppt",
            str(fixture_dir / "current.pptx"),
            "--output",
            str(output),
        ]
    )

    payload = json.loads(output.read_text(encoding="utf-8"))
    assert code == 0
    assert payload["schema_version"] == 2
    assert payload["fingerprint_id"]
    assert "client_ops" not in json.dumps(payload)


def test_package_fingerprint_schema_pins_v2_topology() -> None:
    schema_path = Path(__file__).parents[1] / "qc_tool/package-fingerprint.schema.json"
    schema = json.loads(schema_path.read_text(encoding="utf-8"))

    assert schema["properties"]["schema_version"] == {"const": 2}
    assert schema["properties"]["artifact"] == {"const": "package"}
    assert {"package_manifest", "members", "fingerprint_id"} <= set(
        schema["required"]
    )
    assert schema["properties"]["package_manifest"]["properties"]["version"] == {
        "const": 1
    }
