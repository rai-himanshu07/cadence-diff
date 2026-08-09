"""Materiality magnitude, temporal context, display evidence, and triage."""

from __future__ import annotations

import datetime as dt
import math

import pytest

from qc_tool.config.profile import (
    DeliverableProfile,
    NumericTolerance,
    RestatementWindows,
    SheetProfile,
)
from qc_tool.excel.align import AxisAlignment, RegionAlignment
from qc_tool.excel.diff_values import _RangeSet, diff_region_values
from qc_tool.excel.materiality import (
    DisplayKind,
    DisplayRule,
    classify_numeric_pair,
    infer_date_cadence,
    is_anomalous_magnitude,
    is_representation_noise,
    parse_number_format,
    recent_positions,
    rendered_number,
    temporal_contexts,
    within_ulps,
)
from qc_tool.excel.periods import Period, parse_period
from qc_tool.excel.regions import TableRegion, period_positions
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingEvidenceTag,
    FindingSubtype,
    FindingTemporalContext,
    Materiality,
    Severity,
)
from qc_tool.io.model import CellRecord, SheetSnapshot
from qc_tool.review_series import anchor_matches_finding, anchor_segment


class TestFormatParser:
    @pytest.mark.parametrize("code", ["General", "general", "GENERAL", " General "])
    def test_general_keyword(self, code: str) -> None:
        assert parse_number_format(code) == DisplayRule(kind=DisplayKind.GENERAL)

    @pytest.mark.parametrize(
        ("code", "decimals"),
        [
            ("0", 0),
            ("#,##0", 0),
            ("0.00", 2),
            ("#,##0.0", 1),
            ("0.0#", 2),  # optional decimals count at full width
            ("#,##0.00;[Red](#,##0.00)", 2),  # only the positive section
            ('"$"#,##0.00', 2),
            ("[$€-407] #,##0.00", 2),
            ("0.00_);(0.00)", 2),
            ("* #,##0.0", 1),
        ],
    )
    def test_decimal_families(self, code: str, decimals: int) -> None:
        rule = parse_number_format(code)
        assert rule is not None
        assert rule.kind is DisplayKind.DECIMAL
        assert rule.decimals == decimals
        assert rule.percents == 0
        assert rule.scaling_commas == 0

    def test_percent_formats(self) -> None:
        assert parse_number_format("0%") == DisplayRule(
            kind=DisplayKind.DECIMAL, decimals=0, percents=1
        )
        assert parse_number_format("0.00%") == DisplayRule(
            kind=DisplayKind.DECIMAL, decimals=2, percents=1
        )

    def test_trailing_comma_scaling(self) -> None:
        assert parse_number_format("#,##0.0,") == DisplayRule(
            kind=DisplayKind.DECIMAL, decimals=1, scaling_commas=1
        )
        assert parse_number_format("#,##0,,") == DisplayRule(
            kind=DisplayKind.DECIMAL, decimals=0, scaling_commas=2
        )

    def test_scientific(self) -> None:
        assert parse_number_format("0.00E+00") == DisplayRule(
            kind=DisplayKind.SCIENTIFIC, decimals=2
        )

    @pytest.mark.parametrize(
        "code",
        [
            "yyyy-mm-dd",
            "[h]:mm:ss",
            "h:mm AM/PM",
            "mmm-yy",
            "@",  # text
            "# ?/?",  # fraction
            "0.0.0",  # double decimal point
            '"unterminated',
            "[unterminated",
            "",  # empty: no placeholders
            '"text only"',
            "0.0,0",  # separator inside decimals
            "0.00E+00E+00",  # double exponent
        ],
    )
    def test_unparseable_formats_fail_closed(self, code: str) -> None:
        assert parse_number_format(code) is None

    def test_quoted_datetime_letters_are_literals(self) -> None:
        rule = parse_number_format('#,##0.0" mths"')
        assert rule == DisplayRule(kind=DisplayKind.DECIMAL, decimals=1)


class TestRendering:
    def test_general_uses_eleven_significant_digits(self) -> None:
        rule = DisplayRule(kind=DisplayKind.GENERAL)
        assert rendered_number(1599.6666000000005, rule) == rendered_number(
            1599.6666, rule
        )
        assert rendered_number(1599.6667, rule) != rendered_number(1599.6666, rule)

    def test_negative_zero_normalizes(self) -> None:
        rule = DisplayRule(kind=DisplayKind.DECIMAL, decimals=2)
        assert rendered_number(-0.0000001, rule) == rendered_number(0.0, rule)

    def test_percent_scaling_changes_rendered_precision(self) -> None:
        rule = DisplayRule(kind=DisplayKind.DECIMAL, decimals=0, percents=1)
        assert rendered_number(0.123, rule) == "12"
        assert rendered_number(0.128, rule) == "13"

    def test_thousands_scaling(self) -> None:
        rule = DisplayRule(kind=DisplayKind.DECIMAL, decimals=1, scaling_commas=2)
        assert rendered_number(12_345_678.0, rule) == "12.3"

    def test_non_finite_returns_none(self) -> None:
        rule = DisplayRule(kind=DisplayKind.DECIMAL, decimals=1)
        assert rendered_number(math.inf, rule) is None


class TestNoiseRule:
    def test_run13_style_ulp_noise_is_noise(self) -> None:
        # Representative of the observed cycle pair: identical at one decimal,
        # delta of a couple ULPs on old-history constants.
        assert is_representation_noise(
            3090.8414999999995, 3090.8414999999986, "#,##0.0"
        )
        assert is_representation_noise(1599.6666000000005, 1599.6666, "#,##0.0")
        assert is_representation_noise(9500.4996, 9500.499599999997, "#,##0.0")

    def test_ulp_noise_holds_even_with_unparseable_format(self) -> None:
        base = 3090.8414999999995
        curr = math.nextafter(base, 0.0)
        assert is_representation_noise(base, curr, "yyyy-mm-dd")
        assert is_representation_noise(base, curr, None)

    def test_beyond_ulps_with_unparseable_format_fails_closed(self) -> None:
        assert not is_representation_noise(3090.8415, 3090.8414, "yyyy-mm-dd")

    def test_visible_change_is_not_noise(self) -> None:
        assert not is_representation_noise(1806.0, 1814.0, "#,##0.0")

    def test_display_identical_but_materially_large_is_not_noise(self) -> None:
        # Millions scaling renders both as 12.3 but the delta is a real 53k.
        assert not is_representation_noise(12_345_678.0, 12_299_000.0, "#,##0.0,,")

    def test_near_zero_residue_is_noise_when_display_identical(self) -> None:
        assert is_representation_noise(2e-17, -1e-17, "0.0")

    def test_near_zero_residue_visible_under_general_is_not_noise(self) -> None:
        assert not is_representation_noise(2e-17, -1e-17, "General")

    def test_display_identical_within_ppm_is_noise(self) -> None:
        assert is_representation_noise(3090.8415, 3090.84150001, "#,##0.0")

    def test_bools_and_text_are_never_noise(self) -> None:
        assert not is_representation_noise(True, False, "0.0")
        assert not is_representation_noise("a", "b", "0.0")
        assert not is_representation_noise(1.0, "1.0", "0.0")

    def test_equal_values_are_not_noise(self) -> None:
        assert not is_representation_noise(5.0, 5.0, "0.0")

    def test_within_ulps_boundary(self) -> None:
        base = 1000.0
        four_up = math.nextafter(
            math.nextafter(
                math.nextafter(math.nextafter(base, math.inf), math.inf), math.inf
            ),
            math.inf,
        )
        assert within_ulps(base, four_up)
        assert not within_ulps(base, four_up + 5 * math.ulp(base))


class TestClassifyPrecedence:
    def test_noise_beats_everything(self) -> None:
        tier = classify_numeric_pair(
            3090.8414999999995,
            3090.8414999999986,
            "#,##0.0",
            within_acceptance=True,
        )
        assert tier is Materiality.NOISE

    def test_acceptance_beats_material(self) -> None:
        tier = classify_numeric_pair(100.0, 105.0, "0.0", within_acceptance=True)
        assert tier is Materiality.WITHIN_TOLERANCE

    def test_default_is_material(self) -> None:
        assert classify_numeric_pair(100.0, 105.0, "0.0") is Materiality.MATERIAL

    def test_non_numeric_pairs_have_no_tier(self) -> None:
        assert classify_numeric_pair("x", "y", "0.0") is None
        assert classify_numeric_pair(None, 5.0, "0.0") is None
        assert classify_numeric_pair(True, 1.0, "0.0") is None

    @pytest.mark.parametrize(
        ("baseline", "current"),
        [
            (1.0, 10.0),
            (10.0, 1.0),
            (1.0, 0.0),
            (0.0, -1.0),
            (-1.0, 1.0),
        ],
    )
    def test_hard_magnitude_boundaries(self, baseline: float, current: float) -> None:
        assert is_anomalous_magnitude(baseline, current)

    @pytest.mark.parametrize(
        ("baseline", "current"),
        [(1.0, 9.99), (-1.0, -9.99), (100.0, 101.0), ("x", 1.0)],
    )
    def test_ordinary_or_non_numeric_pairs_are_not_hard_anomalies(
        self, baseline: object, current: object
    ) -> None:
        assert not is_anomalous_magnitude(baseline, current)


def _region_findings(
    baseline_cells: dict[tuple[int, int], CellRecord],
    current_cells: dict[tuple[int, int], CellRecord],
    rows: int,
) -> list[Finding]:
    baseline = SheetSnapshot("Synthetic", "visible", rows, 1, baseline_cells)
    current = SheetSnapshot("Synthetic", "visible", rows, 1, current_cells)
    region = TableRegion("Synthetic", 1, 1, rows, 1, "block", None, 1, "none")
    alignment = RegionAlignment(
        region,
        region,
        AxisAlignment(pairs=[(row, row) for row in range(1, rows + 1)]),
        AxisAlignment(pairs=[(1, 1)], method="positional"),
    )
    return diff_region_values(
        baseline,
        current,
        alignment,
        NumericTolerance(),
        ignore=_RangeSet([]),
        refresh=_RangeSet([]),
        sheet_profile=None,
    )


class TestDiffValuesAttachment:
    def test_replacement_pairs_carry_tiers_and_severity_is_untouched(self) -> None:
        findings = _region_findings(
            {
                (1, 1): CellRecord(1, 1, 3090.8414999999995, number_format="#,##0.0"),
                (2, 1): CellRecord(2, 1, 1806.0, number_format="#,##0.0"),
                (3, 1): CellRecord(3, 1, "label"),
                (5, 1): CellRecord(5, 1, 7.0),
            },
            {
                (1, 1): CellRecord(1, 1, 3090.8414999999986, number_format="#,##0.0"),
                (2, 1): CellRecord(2, 1, 1814.0, number_format="#,##0.0"),
                (3, 1): CellRecord(3, 1, "renamed"),
                (4, 1): CellRecord(4, 1, 9.0),
            },
            5,
        )
        by_location = {finding.location: finding for finding in findings}
        value_findings = {
            loc: f
            for loc, f in by_location.items()
            if f.finding_class is FindingClass.VALUE_CHANGED
        }

        assert value_findings["A1"].materiality is Materiality.NOISE
        assert value_findings["A2"].materiality is Materiality.MATERIAL
        assert value_findings["A3"].materiality is None  # text replacement
        assert value_findings["A4"].subtype is FindingSubtype.VALUE_ADDED_POPULATION
        assert value_findings["A4"].materiality is None
        assert value_findings["A5"].subtype is FindingSubtype.VALUE_CLEARED_POPULATION
        assert value_findings["A5"].materiality is None
        # Step 1 contract: metadata only, triage untouched.
        assert all(f.severity is None for f in findings)

    def test_format_fallback_to_baseline_side(self) -> None:
        findings = _region_findings(
            {(1, 1): CellRecord(1, 1, 3090.8414999999995, number_format="#,##0.0")},
            {(1, 1): CellRecord(1, 1, 3090.8414999999986)},
            1,
        )
        assert findings[0].materiality is Materiality.NOISE


class TestFindingSerialization:
    def test_round_trip_and_legacy_absence(self) -> None:
        finding = Finding(
            artifact="excel",
            finding_class=FindingClass.VALUE_CHANGED,
            materiality=Materiality.NOISE,
            message="x",
        )
        dumped = finding.model_dump(mode="json")
        assert dumped["materiality"] == "noise"
        assert Finding.model_validate(dumped).materiality is Materiality.NOISE

        legacy = {k: v for k, v in dumped.items() if k != "materiality"}
        assert Finding.model_validate(legacy).materiality is None


def _period(value: object) -> Period:
    period = parse_period(value)
    assert period is not None
    return period


class TestDateCadenceInference:
    def test_weekly_monthly_quarterly_spacing(self) -> None:
        weekly = [
            _period(dt.date(2026, 1, 2) + dt.timedelta(weeks=i)) for i in range(6)
        ]
        monthly = [_period(dt.date(2026, month, 1)) for month in range(1, 7)]
        quarterly = [
            _period(dt.date(2026, month, 1)) for month in (1, 4, 7, 10)
        ]
        yearly = [_period(dt.date(year, 1, 1)) for year in (2024, 2025, 2026)]

        assert infer_date_cadence(weekly) == "week"
        assert infer_date_cadence(monthly) == "month"
        assert infer_date_cadence(quarterly) == "quarter"
        assert infer_date_cadence(yearly) is None

    def test_single_date_has_no_cadence(self) -> None:
        assert infer_date_cadence([_period(dt.date(2026, 1, 2))]) is None


class TestRecentPositions:
    def test_weekly_labels_trailing_window(self) -> None:
        positions = {
            row: _period(f"2026-W{week:02d}") for row, week in enumerate(range(1, 21), 2)
        }
        recent = recent_positions(positions, RestatementWindows())
        assert recent == set(range(14, 22))  # last 8 of rows 2..21

    def test_monthly_window_of_two(self) -> None:
        positions = {
            row: _period(f"2026-{month:02d}") for row, month in enumerate(range(1, 8), 1)
        }
        recent = recent_positions(positions, RestatementWindows())
        assert recent == {6, 7}

    def test_quarterly_window_of_one(self) -> None:
        positions = {
            10: _period("Q1 2026"),
            11: _period("Q2 2026"),
            12: _period("Q3 2026"),
        }
        assert recent_positions(positions, RestatementWindows()) == {12}

    def test_weekly_dates_use_week_window(self) -> None:
        positions = {
            row: _period(dt.date(2026, 1, 2) + dt.timedelta(weeks=row - 2))
            for row in range(2, 22)
        }
        recent = recent_positions(positions, RestatementWindows())
        assert recent == set(range(14, 22))

    def test_axis_shorter_than_window_is_fully_recent(self) -> None:
        positions = {row: _period(f"2026-W{week:02d}") for row, week in [(2, 1), (3, 2)]}
        assert recent_positions(positions, RestatementWindows()) == {2, 3}

    def test_zero_window_disables_recency(self) -> None:
        positions = {row: _period(f"2026-W{week:02d}") for row, week in [(2, 1), (3, 2)]}
        windows = RestatementWindows(week=0)
        assert recent_positions(positions, windows) == set()

    def test_mixed_kinds_use_their_own_edges(self) -> None:
        positions = {
            1: _period("2026-05"),
            2: _period("2026-06"),
            3: _period("2026-07"),
            11: _period("Q1 2026"),
            12: _period("Q2 2026"),
        }
        recent = recent_positions(positions, RestatementWindows())
        assert recent == {2, 3, 12}

    def test_duplicate_edge_periods_are_all_recent(self) -> None:
        positions = {
            1: _period("2026-06"),
            2: _period("2026-07"),
            3: _period("2026-07"),
        }
        windows = RestatementWindows(month=1)
        assert recent_positions(positions, windows) == {2, 3}

    def test_separated_monthly_runs_use_independent_observed_edges(self) -> None:
        positions = {
            **{
                row: _period(f"2025-{row - 1:02d}")
                for row in range(2, 14)
            },
            15: _period("2026-01"),
            16: _period("2026-02"),
            17: _period("2026-03"),
        }

        contexts = temporal_contexts(positions, RestatementWindows())

        assert contexts[11] is FindingTemporalContext.HISTORICAL
        assert contexts[12] is FindingTemporalContext.RECENT_WINDOW
        assert contexts[13] is FindingTemporalContext.CURRENT_PERIOD
        assert contexts[15] is FindingTemporalContext.HISTORICAL
        assert contexts[16] is FindingTemporalContext.RECENT_WINDOW
        assert contexts[17] is FindingTemporalContext.CURRENT_PERIOD

    def test_empty_axis(self) -> None:
        assert recent_positions({}, RestatementWindows()) == set()


class TestPeriodPositions:
    def test_long_region_uses_key_column(self) -> None:
        cells = {
            (1, 1): CellRecord(1, 1, "Month"),
            (1, 2): CellRecord(1, 2, "Value"),
        }
        for row, month in enumerate(range(1, 6), 2):
            cells[(row, 1)] = CellRecord(row, 1, f"2026-{month:02d}")
            cells[(row, 2)] = CellRecord(row, 2, float(month))
        sheet = SheetSnapshot("S", "visible", 6, 2, cells)
        region = TableRegion("S", 1, 1, 6, 2, "long", 1, 1, "rows")

        positions = period_positions(sheet, region)
        assert set(positions) == {2, 3, 4, 5, 6}
        assert positions[6].sort_key == (2026, 5, 0)

    def test_long_region_falls_back_past_non_period_key(self) -> None:
        cells = {(1, 1): CellRecord(1, 1, "Name"), (1, 2): CellRecord(1, 2, "Week")}
        for row, week in enumerate(range(1, 5), 2):
            cells[(row, 1)] = CellRecord(row, 1, f"item-{week}")
            cells[(row, 2)] = CellRecord(row, 2, f"W{week:02d} 2026")
        sheet = SheetSnapshot("S", "visible", 5, 2, cells)
        region = TableRegion("S", 1, 1, 5, 2, "long", 1, 1, "rows")

        positions = period_positions(sheet, region)
        assert set(positions) == {2, 3, 4, 5}

    def test_wide_region_uses_header_row(self) -> None:
        cells = {(1, 1): CellRecord(1, 1, "Metric")}
        for col, month in enumerate(range(1, 5), 2):
            cells[(1, col)] = CellRecord(1, col, f"2026-{month:02d}")
            cells[(2, col)] = CellRecord(2, col, float(month))
        sheet = SheetSnapshot("S", "visible", 2, 5, cells)
        region = TableRegion("S", 1, 1, 2, 5, "wide", 1, 1, "columns")

        positions = period_positions(sheet, region)
        assert set(positions) == {2, 3, 4, 5}

    def test_block_region_has_no_positions(self) -> None:
        sheet = SheetSnapshot("S", "visible", 1, 1, {(1, 1): CellRecord(1, 1, "x")})
        region = TableRegion("S", 1, 1, 1, 1, "block", None, 1, "none")
        assert period_positions(sheet, region) == {}


def _long_period_region(
    rows: int, *, baseline_values: dict[int, float], current_values: dict[int, float]
) -> list[Finding]:
    """Rows 2..rows keyed by monthly labels in column A, constants in column B."""

    def cells(values: dict[int, float]) -> dict[tuple[int, int], CellRecord]:
        result = {
            (1, 1): CellRecord(1, 1, "Month"),
            (1, 2): CellRecord(1, 2, "Actual"),
        }
        for row in range(2, rows + 1):
            result[(row, 1)] = CellRecord(row, 1, f"2026-{row - 1:02d}")
            if row in values:
                result[(row, 2)] = CellRecord(row, 2, values[row], number_format="0.0")
        return result

    baseline = SheetSnapshot("S", "visible", rows, 2, cells(baseline_values))
    current = SheetSnapshot("S", "visible", rows, 2, cells(current_values))
    region = TableRegion("S", 1, 1, rows, 2, "long", 1, 1, "rows")
    alignment = RegionAlignment(
        region,
        region,
        AxisAlignment(pairs=[(row, row) for row in range(1, rows + 1)]),
        AxisAlignment(pairs=[(1, 1), (2, 2)]),
    )
    return diff_region_values(
        baseline,
        current,
        alignment,
        NumericTolerance(),
        ignore=_RangeSet([]),
        refresh=_RangeSet([]),
        sheet_profile=None,
    )


class TestRecencyInDiffValues:
    def test_trailing_change_is_recent_and_old_change_is_material(self) -> None:
        rows = 8  # months Jan..Jul in rows 2..8
        baseline = {row: 100.0 + row for row in range(2, 9)}
        current = dict(baseline)
        current[2] = 205.0  # oldest month restated: material
        current[8] = 300.0  # newest month: within 2-month window
        current[7] = 250.0  # second-newest month: within window

        findings = _long_period_region(
            rows, baseline_values=baseline, current_values=current
        )
        axes = {
            f.location: (f.materiality, f.temporal_context)
            for f in findings
            if f.finding_class is FindingClass.VALUE_CHANGED
        }
        assert axes == {
            "B2": (Materiality.MATERIAL, FindingTemporalContext.HISTORICAL),
            "B7": (Materiality.MATERIAL, FindingTemporalContext.RECENT_WINDOW),
            "B8": (Materiality.MATERIAL, FindingTemporalContext.CURRENT_PERIOD),
        }

    def test_noise_wins_over_recency(self) -> None:
        rows = 4
        baseline = dict.fromkeys(range(2, 5), 100.0)
        current = dict(baseline)
        current[4] = math.nextafter(100.0, math.inf)

        findings = _long_period_region(
            rows, baseline_values=baseline, current_values=current
        )
        value_findings = [
            f for f in findings if f.finding_class is FindingClass.VALUE_CHANGED
        ]
        assert [f.materiality for f in value_findings] == [Materiality.NOISE]
        assert value_findings[0].temporal_context is FindingTemporalContext.CURRENT_PERIOD
        assert FindingEvidenceTag.ULP_SCALE in value_findings[0].evidence_tags

    def test_prefilled_future_labels_do_not_define_the_edge(self) -> None:
        """A calendar pre-filled years ahead must anchor the window at the
        last period WITH data, not the last printed label."""
        rows = 12  # months Jan..Nov in rows 2..12, but data stops at row 8
        baseline = {row: 100.0 + row for row in range(2, 9)}
        current = dict(baseline)
        current[8] = 300.0  # last month with data
        current[7] = 250.0  # second-to-last with data
        current[6] = 200.0  # third back: outside the 2-month window

        findings = _long_period_region(
            rows, baseline_values=baseline, current_values=current
        )
        axes = {
            f.location: (f.materiality, f.temporal_context)
            for f in findings
            if f.finding_class is FindingClass.VALUE_CHANGED
        }
        assert axes == {
            "B6": (Materiality.MATERIAL, FindingTemporalContext.HISTORICAL),
            "B7": (Materiality.MATERIAL, FindingTemporalContext.RECENT_WINDOW),
            "B8": (Materiality.MATERIAL, FindingTemporalContext.CURRENT_PERIOD),
        }

    def test_formula_only_future_calendar_does_not_define_the_edge(self) -> None:
        def cells(values: dict[int, float]) -> dict[tuple[int, int], CellRecord]:
            result = {
                (1, 1): CellRecord(1, 1, "Month"),
                (1, 2): CellRecord(1, 2, "Actual"),
            }
            for row in range(2, 11):
                result[(row, 1)] = CellRecord(
                    row,
                    1,
                    f"2026-{row - 1:02d}",
                    formula=(f"=EDATE(A{row - 1},1)" if row >= 7 else None),
                )
                if row in values:
                    result[(row, 2)] = CellRecord(row, 2, values[row])
            return result

        baseline_values = dict.fromkeys(range(2, 7), 100.0)
        current_values = dict(baseline_values)
        current_values.update({4: 140.0, 5: 150.0, 6: 160.0})
        baseline = SheetSnapshot("S", "visible", 10, 2, cells(baseline_values))
        current = SheetSnapshot("S", "visible", 10, 2, cells(current_values))
        region = TableRegion("S", 1, 1, 10, 2, "long", 1, 1, "rows")
        alignment = RegionAlignment(
            region,
            region,
            AxisAlignment(pairs=[(row, row) for row in range(1, 11)]),
            AxisAlignment(pairs=[(1, 1), (2, 2)]),
        )

        findings = diff_region_values(
            baseline,
            current,
            alignment,
            NumericTolerance(),
            ignore=_RangeSet([]),
            refresh=_RangeSet([]),
            sheet_profile=None,
        )
        contexts = {
            finding.location: finding.temporal_context
            for finding in findings
            if finding.finding_class is FindingClass.VALUE_CHANGED
        }

        assert contexts == {
            "B4": FindingTemporalContext.HISTORICAL,
            "B5": FindingTemporalContext.RECENT_WINDOW,
            "B6": FindingTemporalContext.CURRENT_PERIOD,
        }

    def test_profile_round_trip_parses_windows(self) -> None:
        profile = DeliverableProfile.model_validate(
            {"name": "p", "restatement_windows": {"week": 4, "month": 1}}
        )
        assert profile.restatement_windows.week == 4
        assert profile.restatement_windows.month == 1
        assert profile.restatement_windows.quarter == 1

    def test_negative_window_rejected(self) -> None:
        with pytest.raises(ValueError):
            RestatementWindows(week=-1)


class TestAcceptanceBands:
    def _findings(self, bands: list[dict[str, object]]) -> dict[str, Finding]:
        sheet_profile = SheetProfile.model_validate({"acceptance_bands": bands})
        baseline = SheetSnapshot(
            "S",
            "visible",
            2,
            1,
            {
                (1, 1): CellRecord(1, 1, 100.0, number_format="0.0"),
                (2, 1): CellRecord(2, 1, 100.0, number_format="0.0"),
            },
        )
        current = SheetSnapshot(
            "S",
            "visible",
            2,
            1,
            {
                (1, 1): CellRecord(1, 1, 100.4, number_format="0.0"),
                (2, 1): CellRecord(2, 1, 109.0, number_format="0.0"),
            },
        )
        region = TableRegion("S", 1, 1, 2, 1, "block", None, 1, "none")
        alignment = RegionAlignment(
            region,
            region,
            AxisAlignment(pairs=[(1, 1), (2, 2)]),
            AxisAlignment(pairs=[(1, 1)], method="positional"),
        )
        findings = diff_region_values(
            baseline,
            current,
            alignment,
            NumericTolerance(),
            ignore=_RangeSet([]),
            refresh=_RangeSet([]),
            sheet_profile=sheet_profile,
        )
        return {f.location or "": f for f in findings}

    def test_in_band_is_within_tolerance_and_out_of_band_is_material(self) -> None:
        by_location = self._findings(
            [{"range": "A1:A2", "absolute": 0.5}]
        )
        assert by_location["A1"].materiality is Materiality.WITHIN_TOLERANCE
        assert by_location["A2"].materiality is Materiality.MATERIAL

    def test_relative_band(self) -> None:
        by_location = self._findings([{"range": "A1:A2", "relative": 0.1}])
        assert by_location["A1"].materiality is Materiality.WITHIN_TOLERANCE
        assert by_location["A2"].materiality is Materiality.WITHIN_TOLERANCE

    def test_band_outside_range_does_not_apply(self) -> None:
        by_location = self._findings([{"range": "B1:B9", "absolute": 100.0}])
        assert by_location["A1"].materiality is Materiality.MATERIAL
        assert by_location["A2"].materiality is Materiality.MATERIAL


class TestTriageSeverity:
    def _finding(self, **kwargs: object) -> Finding:
        defaults: dict[str, object] = {
            "artifact": "excel",
            "finding_class": FindingClass.VALUE_CHANGED,
            "message": "x",
        }
        defaults.update(kwargs)
        return Finding.model_validate(defaults)

    def test_tier_defaults(self) -> None:
        from qc_tool.triage.rules import assign_severity

        assert (
            assign_severity(self._finding(materiality=Materiality.NOISE))
            is Severity.INFO
        )
        assert (
            assign_severity(self._finding(materiality=Materiality.WITHIN_TOLERANCE))
            is Severity.INFO
        )
        assert (
            assign_severity(self._finding(materiality=Materiality.RECENT_RESTATEMENT))
            is Severity.WARNING
        )
        assert (
            assign_severity(self._finding(materiality=Materiality.MATERIAL))
            is Severity.CRITICAL
        )
        assert assign_severity(self._finding()) is Severity.CRITICAL

    def test_material_temporal_context_defaults(self) -> None:
        from qc_tool.triage.rules import assign_severity

        assert (
            assign_severity(
                self._finding(
                    materiality=Materiality.MATERIAL,
                    temporal_context=FindingTemporalContext.CURRENT_PERIOD,
                )
            )
            is Severity.WARNING
        )
        assert (
            assign_severity(
                self._finding(
                    materiality=Materiality.MATERIAL,
                    temporal_context=FindingTemporalContext.RECENT_WINDOW,
                )
            )
            is Severity.WARNING
        )
        assert (
            assign_severity(
                self._finding(
                    materiality=Materiality.MATERIAL,
                    temporal_context=FindingTemporalContext.HISTORICAL,
                )
            )
            is Severity.CRITICAL
        )

    def test_profile_tier_override_restores_strictness(self) -> None:
        from qc_tool.triage.rules import assign_severity

        profile = DeliverableProfile.model_validate(
            {"name": "p", "materiality_severity": {"noise": "critical"}}
        )
        assert (
            assign_severity(self._finding(materiality=Materiality.NOISE), profile)
            is Severity.CRITICAL
        )

    def test_profile_can_soften_material(self) -> None:
        from qc_tool.triage.rules import assign_severity

        profile = DeliverableProfile.model_validate(
            {"name": "p", "materiality_severity": {"material": "warning"}}
        )
        assert (
            assign_severity(self._finding(materiality=Materiality.MATERIAL), profile)
            is Severity.WARNING
        )

    def test_expected_growth_still_wins(self) -> None:
        from qc_tool.triage.rules import assign_severity

        finding = self._finding(
            materiality=Materiality.MATERIAL, expected_growth=True
        )
        assert assign_severity(finding) is Severity.EXPECTED

    def test_inherited_error_needs_explicit_na_evidence_to_demote(self) -> None:
        from qc_tool.findings import FindingEvidenceTag, FindingProvenance
        from qc_tool.triage.rules import assign_severity

        inherited_error = self._finding(
            finding_class=FindingClass.FORMULA_ERROR,
            provenance=FindingProvenance.INHERITED,
        )
        inherited_explicit_na = self._finding(
            finding_class=FindingClass.FORMULA_ERROR,
            provenance=FindingProvenance.INHERITED,
            evidence_tags={
                FindingEvidenceTag.EXPLICIT_NA,
                FindingEvidenceTag.FORMULA_TEXT,
                FindingEvidenceTag.FORMULA_PRESENCE,
            },
        )
        inherited_outlier = self._finding(
            finding_class=FindingClass.FORMULA_INCONSISTENT,
            provenance=FindingProvenance.INHERITED,
        )
        new_error = self._finding(
            finding_class=FindingClass.FORMULA_ERROR,
            provenance=FindingProvenance.NEW,
        )
        historical_outlier = self._finding(
            finding_class=FindingClass.FORMULA_INCONSISTENT,
            provenance=FindingProvenance.HISTORICAL_PATTERN,
        )

        assert assign_severity(inherited_error) is Severity.CRITICAL
        assert assign_severity(inherited_explicit_na) is Severity.INFO
        assert assign_severity(inherited_outlier) is Severity.INFO
        assert assign_severity(new_error) is Severity.CRITICAL
        assert assign_severity(historical_outlier) is Severity.WARNING

    def test_inherited_value_change_is_not_demoted(self) -> None:
        from qc_tool.findings import FindingProvenance
        from qc_tool.triage.rules import assign_severity

        finding = self._finding(provenance=FindingProvenance.INHERITED)
        assert assign_severity(finding) is Severity.CRITICAL


class TestLintAcceptanceBands:
    def test_zero_band_and_bad_range_are_errors(self) -> None:
        from qc_tool.config.lint import lint_profile

        profile = DeliverableProfile.model_validate(
            {
                "name": "p",
                "excel": {
                    "sheets": {
                        "S": {
                            "acceptance_bands": [
                                {"range": "A1:A2"},
                                {"range": "not a range", "absolute": 1.0},
                            ]
                        }
                    }
                },
            }
        )
        issues = lint_profile(profile)
        messages = [issue.message for issue in issues]
        assert any("accepts nothing" in message for message in messages)
        assert any("not a valid A1 range" in message for message in messages)


def _wide_period_region(
    cols: int, *, baseline_values: dict[int, float], current_values: dict[int, float]
) -> list[Finding]:
    """Columns 2..cols keyed by monthly labels in row 1, constants in row 2."""

    def cells(values: dict[int, float]) -> dict[tuple[int, int], CellRecord]:
        result = {
            (1, 1): CellRecord(1, 1, "Metric"),
            (2, 1): CellRecord(2, 1, "Actual"),
        }
        for col in range(2, cols + 1):
            result[(1, col)] = CellRecord(1, col, f"2026-{col - 1:02d}")
            if col in values:
                result[(2, col)] = CellRecord(2, col, values[col], number_format="0.0")
        return result

    baseline = SheetSnapshot("S", "visible", 2, cols, cells(baseline_values))
    current = SheetSnapshot("S", "visible", 2, cols, cells(current_values))
    region = TableRegion("S", 1, 1, 2, cols, "wide", 1, 1, "columns")
    alignment = RegionAlignment(
        region,
        region,
        AxisAlignment(pairs=[(1, 1), (2, 2)]),
        AxisAlignment(pairs=[(col, col) for col in range(1, cols + 1)]),
    )
    return diff_region_values(
        baseline,
        current,
        alignment,
        NumericTolerance(),
        ignore=_RangeSet([]),
        refresh=_RangeSet([]),
        sheet_profile=None,
    )


def _two_column_period_region(
    rows: int,
    *,
    baseline_values: dict[tuple[int, int], float],
    current_values: dict[tuple[int, int], float],
) -> list[Finding]:
    """Rows 2..rows keyed by monthly labels in column A, constants in B and C."""

    def cells(
        values: dict[tuple[int, int], float],
    ) -> dict[tuple[int, int], CellRecord]:
        result = {
            (1, 1): CellRecord(1, 1, "Month"),
            (1, 2): CellRecord(1, 2, "Actual"),
            (1, 3): CellRecord(1, 3, "Target"),
        }
        for row in range(2, rows + 1):
            result[(row, 1)] = CellRecord(row, 1, f"2026-{row - 1:02d}")
            for col in (2, 3):
                if (row, col) in values:
                    result[(row, col)] = CellRecord(
                        row, col, values[(row, col)], number_format="0.0"
                    )
        return result

    baseline = SheetSnapshot("S", "visible", rows, 3, cells(baseline_values))
    current = SheetSnapshot("S", "visible", rows, 3, cells(current_values))
    region = TableRegion("S", 1, 1, rows, 3, "long", 1, 1, "rows")
    alignment = RegionAlignment(
        region,
        region,
        AxisAlignment(pairs=[(row, row) for row in range(1, rows + 1)]),
        AxisAlignment(pairs=[(1, 1), (2, 2), (3, 3)]),
    )
    return diff_region_values(
        baseline,
        current,
        alignment,
        NumericTolerance(),
        ignore=_RangeSet([]),
        refresh=_RangeSet([]),
        sheet_profile=None,
    )


class TestSeriesAnchorProducer:
    def test_rows_period_region_anchors_measure_column_and_period_row(self) -> None:
        rows = 8
        baseline = {row: 100.0 + row for row in range(2, 9)}
        current = dict(baseline)
        current[2] = 205.0
        current[8] = 300.0

        findings = _long_period_region(
            rows, baseline_values=baseline, current_values=current
        )
        anchored = {
            finding.location or "": finding.series_anchor
            for finding in findings
            if finding.series_anchor is not None
        }

        assert set(anchored) == {"B2", "B8"}
        for location, anchor in anchored.items():
            assert anchor is not None
            assert anchor.version == 2
            assert anchor_segment(anchor) == "restatement"
            assert anchor.sheet == "S"
            assert anchor.current_region_id == "S!A1:B8"
            assert anchor.period_axis == "rows"
            assert anchor.series_index == 2
            assert anchor.period_index == int(location[1:])

        for finding in findings:
            if finding.series_anchor is None:
                continue
            assert anchor_matches_finding(finding, finding.series_anchor)

    def test_columns_period_region_anchors_measure_row_and_period_column(self) -> None:
        cols = 8
        baseline = {col: 100.0 + col for col in range(2, 9)}
        current = dict(baseline)
        current[2] = 205.0
        current[8] = 300.0

        findings = _wide_period_region(
            cols, baseline_values=baseline, current_values=current
        )
        anchored = {
            finding.location or "": finding.series_anchor
            for finding in findings
            if finding.series_anchor is not None
        }

        assert set(anchored) == {"B2", "H2"}
        for anchor in anchored.values():
            assert anchor is not None
            assert anchor.current_region_id == "S!A1:H2"
            assert anchor.period_axis == "columns"
            assert anchor.series_index == 2
        assert anchored["B2"] is not None and anchored["B2"].period_index == 2
        assert anchored["H2"] is not None and anchored["H2"].period_index == 8

        for finding in findings:
            if finding.series_anchor is None:
                continue
            assert anchor_matches_finding(finding, finding.series_anchor)

    def test_block_region_text_and_population_edits_never_anchor(self) -> None:
        findings = _region_findings(
            {
                (1, 1): CellRecord(1, 1, 1806.0, number_format="#,##0.0"),
                (2, 1): CellRecord(2, 1, "label"),
                (4, 1): CellRecord(4, 1, 7.0),
            },
            {
                (1, 1): CellRecord(1, 1, 1814.0, number_format="#,##0.0"),
                (2, 1): CellRecord(2, 1, "renamed"),
                (3, 1): CellRecord(3, 1, 9.0),
            },
            4,
        )

        assert findings
        assert all(finding.series_anchor is None for finding in findings)

    def test_added_period_anchors_as_a_new_period_and_lone_clear_never_does(self) -> None:
        rows = 8
        baseline = {row: 100.0 + row for row in range(2, 9)}
        baseline.pop(3)  # blank baseline -> added population
        current = dict(baseline)
        current[3] = 42.0
        current.pop(4)  # cleared population; the ONLY measure column, so the
        # period position dies with it and no anchor may exist

        findings = _long_period_region(
            rows, baseline_values=baseline, current_values=current
        )
        by_location = {finding.location: finding for finding in findings}

        added = by_location["B3"]
        assert added.subtype is FindingSubtype.VALUE_ADDED_POPULATION
        assert added.series_anchor is not None
        assert anchor_segment(added.series_anchor) == "new_period"
        assert added.series_anchor.series_index == 2
        assert added.series_anchor.period_index == 3
        assert anchor_matches_finding(added, added.series_anchor)
        # the public finding is untouched, so its canonical decision cannot move
        assert added.materiality is None
        assert added.temporal_context is None

        cleared = by_location["B4"]
        assert cleared.subtype is FindingSubtype.VALUE_CLEARED_POPULATION
        assert cleared.series_anchor is None

    def test_single_cell_clear_anchors_when_a_sibling_keeps_the_period_alive(
        self,
    ) -> None:
        rows = 8
        baseline = {
            (row, col): 100.0 * col + row for row in range(2, 9) for col in (2, 3)
        }
        current = dict(baseline)
        current.pop((4, 2))  # clear Actual for 2026-03; Target C4 keeps row 4 alive

        findings = _two_column_period_region(
            rows, baseline_values=baseline, current_values=current
        )
        cleared = next(f for f in findings if f.location == "B4")

        assert cleared.subtype is FindingSubtype.VALUE_CLEARED_POPULATION
        assert cleared.series_anchor is not None
        assert anchor_segment(cleared.series_anchor) == "cleared_period"
        assert cleared.series_anchor.series_index == 2
        assert cleared.series_anchor.period_index == 4
        assert anchor_matches_finding(cleared, cleared.series_anchor)
        # nothing public moved: no materiality, no temporal context
        assert cleared.materiality is None
        assert cleared.temporal_context is None
        assert "series_anchor" not in cleared.model_dump(mode="json")

    def test_whole_row_wipe_earns_no_anchor_for_any_cleared_cell(self) -> None:
        rows = 8
        baseline = {
            (row, col): 100.0 * col + row for row in range(2, 9) for col in (2, 3)
        }
        current = dict(baseline)
        current.pop((4, 2))
        current.pop((4, 3))  # the whole 2026-03 row is gone: one event

        findings = _two_column_period_region(
            rows, baseline_values=baseline, current_values=current
        )
        wiped = [f for f in findings if f.location in {"B4", "C4"}]

        assert len(wiped) == 2
        assert all(
            f.subtype is FindingSubtype.VALUE_CLEARED_POPULATION for f in wiped
        )
        assert all(f.series_anchor is None for f in wiped)

    def test_cleared_text_baseline_never_anchors(self) -> None:
        rows = 8
        baseline: dict[tuple[int, int], object] = {
            (row, col): 100.0 * col + row for row in range(2, 9) for col in (2, 3)
        }
        baseline[(4, 2)] = "n/a"  # text baseline
        current = dict(baseline)
        current.pop((4, 2))  # cleared, but C4 keeps the period alive

        findings = _two_column_period_region(
            rows,
            baseline_values=baseline,  # type: ignore[arg-type]
            current_values=current,  # type: ignore[arg-type]
        )
        cleared = next(f for f in findings if f.location == "B4")

        assert cleared.subtype is FindingSubtype.VALUE_CLEARED_POPULATION
        assert cleared.series_anchor is None

    def test_added_period_needs_a_dated_position_and_a_numeric_value(self) -> None:
        rows = 8
        baseline: dict[int, object] = {row: 100.0 + row for row in range(2, 9)}
        baseline.pop(3)
        current = dict(baseline)
        current[3] = "n/a"  # text addition

        findings = _long_period_region(
            rows,
            baseline_values=baseline,  # type: ignore[arg-type]
            current_values=current,  # type: ignore[arg-type]
        )
        added = next(f for f in findings if f.location == "B3")

        assert added.subtype is FindingSubtype.VALUE_ADDED_POPULATION
        assert added.series_anchor is None

    def test_added_period_outside_a_period_region_never_anchors(self) -> None:
        findings = _region_findings(
            {(1, 1): CellRecord(1, 1, 1806.0, number_format="#,##0.0")},
            {
                (1, 1): CellRecord(1, 1, 1814.0, number_format="#,##0.0"),
                (3, 1): CellRecord(3, 1, 9.0),
            },
            4,
        )
        added = next(f for f in findings if f.location == "A3")

        assert added.subtype is FindingSubtype.VALUE_ADDED_POPULATION
        assert added.series_anchor is None

    def test_anchors_never_reach_public_serialization(self) -> None:
        rows = 8
        baseline = {row: 100.0 + row for row in range(2, 9)}
        current = dict(baseline)
        current[8] = 300.0

        findings = _long_period_region(
            rows, baseline_values=baseline, current_values=current
        )
        anchored = [f for f in findings if f.series_anchor is not None]

        assert anchored
        for finding in anchored:
            assert "series_anchor" not in finding.model_dump(mode="json")
