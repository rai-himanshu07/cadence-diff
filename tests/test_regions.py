"""Region detection, orientation inference, and profile override tests."""

from pathlib import Path

import pytest

from qc_tool.config.profile import (
    DeliverableProfile,
    load_profile,
    save_profile,
)
from qc_tool.excel.periods import is_period_after, parse_period
from qc_tool.excel.regions import TableRegion, detect_regions
from qc_tool.io.loader import load_workbook_snapshot


@pytest.mark.parametrize(
    ("label", "kind", "sort_key"),
    [
        ("Jan-26", "month", (2026, 1, 0)),
        ("Feb 2026", "month", (2026, 2, 0)),
        ("2026-07", "month", (2026, 7, 0)),
        ("2026-07-26", "date", (2026, 7, 26)),
        ("W05", "week", (0, 0, 5)),
        ("CW30", "week", (0, 0, 30)),
        ("W01-27", "week", (2027, 0, 1)),
        ("2027-W01", "week", (2027, 0, 1)),
        ("Q1 FY26", "quarter", (2026, 1, 0)),
        ("Q3", "quarter", (0, 3, 0)),
    ],
)
def test_parse_period_valid(label: str, kind: str, sort_key: tuple[int, int, int]) -> None:
    period = parse_period(label)
    assert period is not None
    assert period.kind == kind
    assert period.sort_key == sort_key


@pytest.mark.parametrize("label", ["Revenue", "North", "Period", "", None, 42.0, "Week"])
def test_parse_period_rejects_non_periods(label: object) -> None:
    assert parse_period(label) is None


def test_cross_kind_periods_are_not_directly_orderable() -> None:
    month = parse_period("Apr-26")
    quarter = parse_period("Q1 FY26")
    assert month is not None and quarter is not None

    assert not is_period_after(month, quarter)
    assert not is_period_after(quarter, month)


def _regions_by_range(regions: list[TableRegion]) -> dict[str, TableRegion]:
    return {r.cell_range: r for r in regions}


def test_long_monthly_detected_as_long(fixture_dir: Path) -> None:
    snap = load_workbook_snapshot(fixture_dir / "current.xlsx")
    regions = detect_regions(snap.sheet("Long_Monthly"))
    assert len(regions) == 1
    region = regions[0]
    assert region.cell_range == "A1:E25"
    assert region.orientation == "long"
    assert region.period_axis == "rows"
    assert region.header_row == 1
    assert region.key_col == 1  # Period column


def test_wide_weekly_detected_as_wide(fixture_dir: Path) -> None:
    snap = load_workbook_snapshot(fixture_dir / "current.xlsx")
    regions = detect_regions(snap.sheet("Wide_Weekly"))
    assert len(regions) == 1
    region = regions[0]
    assert region.orientation == "wide"
    assert region.period_axis == "columns"
    assert region.min_row == 1 and region.max_row == 4


def test_dashboard_multi_block_detection(fixture_dir: Path) -> None:
    snap = load_workbook_snapshot(fixture_dir / "current.xlsx")
    regions = detect_regions(snap.sheet("Dashboard"))
    assert len(regions) == 2
    by_range = _regions_by_range(regions)

    kpi = by_range["A1:B5"]
    assert kpi.orientation == "block"
    assert kpi.period_axis == "none"

    headcount = by_range["E2:K4"]
    assert headcount.orientation == "wide"
    assert headcount.period_axis == "columns"


def test_summary_is_single_block(fixture_dir: Path) -> None:
    snap = load_workbook_snapshot(fixture_dir / "current.xlsx")
    regions = detect_regions(snap.sheet("Summary"))
    assert len(regions) == 1
    assert regions[0].orientation == "block"


def test_profile_override_wins(fixture_dir: Path) -> None:
    snap = load_workbook_snapshot(fixture_dir / "current.xlsx")
    profile = DeliverableProfile.model_validate(
        {
            "name": "test",
            "excel": {
                "sheets": {
                    "Dashboard": {
                        "regions": [
                            {"range": "A1:K5", "orientation": "block", "key_column": "A"}
                        ]
                    }
                }
            },
        }
    )
    regions = detect_regions(snap.sheet("Dashboard"), profile.sheet_profile("Dashboard"))
    assert len(regions) == 1
    assert regions[0].cell_range == "A1:K5"
    assert regions[0].orientation == "block"
    assert regions[0].key_col == 1


def test_profile_ignore_sheet(fixture_dir: Path) -> None:
    snap = load_workbook_snapshot(fixture_dir / "current.xlsx")
    profile = DeliverableProfile.model_validate(
        {"name": "test", "excel": {"sheets": {"Dashboard": {"ignore": True}}}}
    )
    assert detect_regions(snap.sheet("Dashboard"), profile.sheet_profile("Dashboard")) == []


def test_profile_yaml_roundtrip(tmp_path: Path) -> None:
    profile = DeliverableProfile.model_validate(
        {
            "name": "monthly-pack",
            "tolerance": {"absolute": 0.01, "relative": 0.001},
            "excel": {
                "ignore_sheets": ["Params"],
                "sheets": {
                    "Long_Monthly": {
                        "cadence_bands": [
                            {"range": "A2:A25", "kind": "month"},
                        ],
                        "regions": [
                            {
                                "range": "A1:E25",
                                "orientation": "long",
                                "key_column": "A",
                                "header_row": 1,
                            }
                        ]
                    }
                },
            },
        }
    )
    path = tmp_path / "profile.yaml"
    save_profile(profile, path)
    loaded = load_profile(path)
    assert loaded == profile
    assert loaded.excel.sheets["Long_Monthly"].cadence_bands[0].kind == "month"
