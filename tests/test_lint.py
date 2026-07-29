"""Profile lint tests: static checks and checks against real files."""

from pathlib import Path

import pytest

from qc_tool.config.lint import lint_profile
from qc_tool.config.profile import DeliverableProfile
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.ppt.extract import load_deck_snapshot


@pytest.fixture(scope="module")
def workbook(fixture_dir: Path):
    return load_workbook_snapshot(fixture_dir / "current.xlsx")


@pytest.fixture(scope="module")
def deck(fixture_dir: Path):
    return load_deck_snapshot(fixture_dir / "current.pptx")


def test_clean_profile_passes(workbook, deck) -> None:
    profile = DeliverableProfile.model_validate(
        {
            "name": "clean",
            "excel": {"sheets": {"Dashboard": {"refresh_ranges": ["B2:B5"]}}},
            "crosscheck": {
                "mappings": [
                    {
                        "slide": "Executive Summary",
                        "line_skeleton": "Total revenue $#M",
                        "figure_index": 0,
                        "label": "Total revenue",
                        "source_sheet": "Dashboard",
                        "source_cell": "B2",
                    }
                ]
            },
            "ppt": {"required_slides": ["Executive Summary"]},
        }
    )
    issues = lint_profile(profile, workbook=workbook, deck=deck)
    assert [i for i in issues if i.level == "error"] == []


def test_static_errors_detected() -> None:
    profile = DeliverableProfile.model_validate(
        {
            "name": "broken",
            "excel": {
                "sheets": {
                    "Dashboard": {
                        "refresh_ranges": ["NOT A RANGE"],
                        "cadence_bands": [
                            {"range": "ALSO NOT A RANGE", "kind": "month"}
                        ],
                    }
                },
                "controls": {
                    "tie_outs": [
                        {"name": "empty", "target": "no-sheet-ref", "components": []}
                    ]
                },
            },
            "waivers": [
                {
                    "finding_class": "value_changed",
                    "reason": "agreed with client",
                    "expires": "2020-01-01",
                }
            ],
        }
    )
    issues = lint_profile(profile)
    messages = {(i.level, i.message.split(" ")[0]) for i in issues}
    assert any(i.level == "error" and "NOT A RANGE" in i.message for i in issues)
    assert any(
        i.level == "error" and "ALSO NOT A RANGE" in i.message for i in issues
    )
    assert any(i.level == "error" and "no components" in i.message for i in issues)
    assert any(i.level == "error" and "Sheet!Cell" in i.message for i in issues)
    assert any(i.level == "warning" and "expired" in i.message for i in issues)
    assert messages  # sanity


def test_against_workbook_detects_missing_targets(workbook) -> None:
    profile = DeliverableProfile.model_validate(
        {
            "name": "stale",
            "excel": {"sheets": {"Ghost_Sheet": {}}, "ignore_sheets": ["Old_Gone"]},
            "crosscheck": {
                "mappings": [
                    {
                        "slide": "Executive Summary",
                        "line_skeleton": "Total revenue $#M",
                        "source_sheet": "Dashboard",
                        "source_cell": "Z99",
                        "label": "dead cell",
                    }
                ]
            },
        }
    )
    issues = lint_profile(profile, workbook=workbook)
    errors = [i for i in issues if i.level == "error"]
    assert any("Ghost_Sheet" in i.where and "not found" in i.message for i in errors)
    assert any("Old_Gone" in i.message for i in errors)
    assert any("Z99" in i.message and "empty" in i.message for i in errors)


def test_against_deck_detects_unresolvable_anchors(workbook, deck) -> None:
    profile = DeliverableProfile.model_validate(
        {
            "name": "deck-stale",
            "ppt": {"required_slides": ["Slide That Left"]},
            "crosscheck": {
                "mappings": [
                    {
                        "slide": "Executive Summary",
                        "line_skeleton": "A line nobody wrote #",
                        "source_sheet": "Dashboard",
                        "source_cell": "B2",
                        "label": "ghost anchor",
                    }
                ]
            },
        }
    )
    issues = lint_profile(profile, workbook=workbook, deck=deck)
    errors = [i for i in issues if i.level == "error"]
    assert any("Slide That Left" in i.message for i in errors)
    assert any("does not resolve" in i.message for i in errors)


def test_duplicate_anchor_flagged() -> None:
    mapping = {
        "slide": "S",
        "line_skeleton": "x #",
        "source_sheet": "A",
        "source_cell": "B2",
    }
    profile = DeliverableProfile.model_validate(
        {"name": "dup", "crosscheck": {"mappings": [mapping, dict(mapping)]}}
    )
    issues = lint_profile(profile)
    assert any(i.message == "duplicate mapping anchor" for i in issues)


def test_valid_tie_out_range_passes_lint(workbook) -> None:
    profile = DeliverableProfile.model_validate(
        {
            "name": "range-tie-out",
            "excel": {
                "controls": {
                    "tie_outs": [
                        {
                            "name": "Revenue sum",
                            "target": "Dashboard!B2",
                            "components": ["Long_Monthly!C2:C5"],
                        }
                    ]
                }
            },
        }
    )
    errors = [
        issue for issue in lint_profile(profile, workbook=workbook) if issue.level == "error"
    ]
    assert errors == []


def test_inverted_numeric_bounds_are_rejected() -> None:
    profile = DeliverableProfile.model_validate(
        {
            "name": "bad-bounds",
            "excel": {
                "controls": {
                    "numeric_bounds": [
                        {
                            "name": "impossible",
                            "sheet": "Data",
                            "range": "A1",
                            "minimum": 100,
                            "maximum": 10,
                        }
                    ]
                }
            },
        }
    )
    issues = lint_profile(profile)
    assert any(issue.level == "error" and "exceeds maximum" in issue.message for issue in issues)
