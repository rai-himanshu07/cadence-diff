"""Region detection, orientation inference, and profile override tests."""

import datetime as dt
from pathlib import Path

import pytest

from qc_tool.config.profile import (
    DeliverableProfile,
    load_profile,
    save_profile,
)
from qc_tool.excel.periods import is_period_after, parse_period
from qc_tool.excel.regions import (
    TableRegion,
    detect_regions,
    internal_period_band_suggestions,
)
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.io.model import CellRecord, CellValue, SheetSnapshot


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


# --- Step 3: evidence-scored internal period-band discovery ---------------


def _sheet(name: str, values: dict[tuple[int, int], CellValue]) -> SheetSnapshot:
    cells = {
        coordinate: CellRecord(row=coordinate[0], column=coordinate[1], value=value)
        for coordinate, value in values.items()
        if value is not None
    }
    max_row = max((row for row, _ in cells), default=0)
    max_col = max((col for _, col in cells), default=0)
    return SheetSnapshot(
        name=name, visibility="visible", max_row=max_row, max_column=max_col, cells=cells
    )


def test_dominant_period_band_auto_selects_as_long(fixture_dir: Path) -> None:
    months = [dt.date(2026, m, 1) for m in range(1, 7)]  # 6 monotonic, contiguous
    values: dict[tuple[int, int], CellValue] = {(1, 1): "Label", (1, 2): "Value"}
    for index, month in enumerate(months, start=2):
        values[(index, 1)] = month
        values[(index, 2)] = 100.0 + index
    sheet = _sheet("Dates", values)

    regions = detect_regions(sheet)

    assert len(regions) == 1
    assert regions[0].orientation == "long"
    assert regions[0].period_axis == "rows"
    assert regions[0].key_col == 1


def test_competing_columns_fall_back_to_block_instead_of_guessing(fixture_dir: Path) -> None:
    """Two columns both clear the strict evidence bar with near-identical
    scores (a percent-formatted numeric row's neighbor happening to also look
    period-like): the detector must refuse to pick one, not guess."""
    months = [dt.date(2026, m, 1) for m in range(1, 7)]
    quarters = ["Q1", "Q1", "Q1", "Q2", "Q2", "Q2"]
    values: dict[tuple[int, int], CellValue] = {(1, 1): "A", (1, 2): "B", (1, 3): "C"}
    for index, (month, quarter) in enumerate(zip(months, quarters, strict=True), start=2):
        values[(index, 1)] = month
        values[(index, 2)] = quarter
        values[(index, 3)] = 100.0 + index
    sheet = _sheet("Competing", values)

    regions = detect_regions(sheet)

    assert len(regions) == 1
    assert regions[0].orientation == "block"
    assert regions[0].period_axis == "none"


def test_internal_period_header_produces_profile_suggestion() -> None:
    values: dict[tuple[int, int], CellValue] = {
        (1, 1): "Dashboard",
        (1, 2): "Metric",
        (2, 1): "Current view",
        (3, 1): "Period",
        (4, 1): "Revenue",
    }
    for column, month in enumerate(range(1, 7), start=2):
        values[(3, column)] = dt.date(2026, month, 1)
        values[(4, column)] = float(column)
    sheet = _sheet("Internal", values)

    regions = detect_regions(sheet)
    suggestions = internal_period_band_suggestions(sheet, regions)

    assert len(regions) == 1
    assert regions[0].orientation == "block"
    assert len(suggestions) == 1
    assert suggestions[0].axis == "columns"
    assert suggestions[0].anchor == 3
    assert suggestions[0].period_count == 6


def test_percent_formatted_numbers_never_look_like_a_period_band(fixture_dir: Path) -> None:
    """Plain percent-formatted numbers (not text, not date objects) can never
    satisfy `parse_period`, so they never compete for the axis at all."""
    values: dict[tuple[int, int], CellValue] = {(1, 1): "Label"}
    for index, ratio in enumerate([0.1, 0.2, 0.3, 0.4, 0.5, 0.6], start=2):
        values[(index, 1)] = ratio
    sheet = _sheet("Percent", values)

    regions = detect_regions(sheet)

    assert len(regions) == 1
    assert regions[0].orientation == "block"


def test_short_fixture_below_the_strict_bar_still_uses_lenient_fallback(
    fixture_dir: Path,
) -> None:
    """A 3-period header (below the new 4-period strict gate) must still be
    detected via the preserved lenient fallback, not silently dropped."""
    values: dict[tuple[int, int], CellValue] = {
        (1, 1): "Label",
        (1, 2): "Jan-26",
        (1, 3): "Feb-26",
        (1, 4): "Mar-26",
        (2, 1): "A",
        (2, 2): 1.0,
        (2, 3): 2.0,
        (2, 4): 3.0,
    }
    sheet = _sheet("Short", values)

    regions = detect_regions(sheet)

    assert len(regions) == 1
    assert regions[0].orientation == "wide"
    assert regions[0].period_axis == "columns"


def test_a_backward_step_disqualifies_the_strict_band_but_legacy_fallback_is_unchanged(
    fixture_dir: Path,
) -> None:
    """The strict evidence gate rejects a non-monotonic band outright (it
    never becomes a scored candidate); the preserved legacy fallback never
    checked order either, so a genuinely period-shaped column is still found
    the same way it always was -- this pins that the two paths compose
    correctly rather than silently masking one another."""
    values: dict[tuple[int, int], CellValue] = {(1, 1): "Label"}
    out_of_order = [
        dt.date(2026, 1, 1),
        dt.date(2026, 2, 1),
        dt.date(2026, 1, 15),  # backward step: disqualifies the STRICT gate
        dt.date(2026, 4, 1),
        dt.date(2026, 5, 1),
    ]
    for index, value in enumerate(out_of_order, start=2):
        values[(index, 1)] = value
    sheet = _sheet("OutOfOrder", values)

    regions = detect_regions(sheet)

    assert len(regions) == 1
    # Not "block": below-strict-bar candidates still fall back to the
    # existing lenient count/share check, which is order-agnostic and always
    # has been -- this is the "existing fixtures unchanged" contract.
    assert regions[0].orientation == "long"


def test_monotonic_gate_excludes_a_disordered_candidate_from_scoring(fixture_dir: Path) -> None:
    """A weaker-but-ordered column must win over a longer disordered one,
    proving the disordered column never entered the scored candidate set at
    all (rather than merely losing a close score comparison)."""
    values: dict[tuple[int, int], CellValue] = {(1, 1): "Label", (1, 2): "Other"}
    disordered = [
        dt.date(2026, 1, 1),
        dt.date(2026, 2, 1),
        dt.date(2026, 1, 15),  # backward step disqualifies column 1 entirely
        dt.date(2026, 4, 1),
        dt.date(2026, 5, 1),
    ]
    ordered = [dt.date(2026, m, 1) for m in range(1, 5)]  # exactly 4, valid
    for index, value in enumerate(disordered, start=2):
        values[(index, 1)] = value
    for index, value in enumerate(ordered, start=2):
        values[(index, 2)] = value
    sheet = _sheet("Disordered", values)

    regions = detect_regions(sheet)

    assert len(regions) == 1
    assert regions[0].orientation == "long"
    assert regions[0].key_col == 2  # the ordered column won, not column 1
