"""Structural fingerprint and fail-closed privacy CLI contracts."""

import json
from pathlib import Path

from qc_tool import cli
from qc_tool.fingerprint import fingerprint_file


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
