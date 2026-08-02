"""Run-14 audit fixes: table growth, block alignment, key volatility, refresh.

Each defect was proven manually against a representative cycle pair: a
sideways table extension blessed as cadence growth, KPI panels key-aligned
by their own values, derived date columns churning composite row identity,
and refresh blocks laundering text/config flips.
"""

from __future__ import annotations

import datetime as dt

import pytest

from qc_tool.config.profile import DeliverableProfile, NumericTolerance
from qc_tool.excel.align import align_regions
from qc_tool.excel.diff_structure import _table_findings
from qc_tool.excel.diff_values import _RangeSet, diff_region_values
from qc_tool.excel.regions import TableRegion
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingExpectedReason,
    FindingTemporalContext,
    Severity,
)
from qc_tool.io.model import (
    CellRecord,
    CellValue,
    SheetSnapshot,
    TableDescriptor,
    WorkbookSnapshot,
)
from qc_tool.triage.rules import assign_severity


def _table_workbook(
    cell_range: str,
    columns: list[str],
    *,
    cells: dict[tuple[int, int], CellRecord] | None = None,
) -> WorkbookSnapshot:
    return WorkbookSnapshot(
        source_name="tables.xlsx",
        file_format="xlsx",
        formulas_available=True,
        styles_available=True,
        tables_available=True,
        sheets=[SheetSnapshot("Data", "visible", 30, 30, cells or {})],
        tables=[
            TableDescriptor(
                sheet="Data",
                name="tbl_Setup",
                display_name="tbl_Setup",
                cell_range=cell_range,
                columns=columns,
                header_row_count=1,
                totals_row_count=0,
                source_id=1,
            )
        ],
    )


class TestTableRangeGrowth:
    def test_sideways_extension_with_new_columns_is_never_expected(self) -> None:
        baseline = _table_workbook("B4:T18", [f"c{i}" for i in range(19)])
        current = _table_workbook(
            "B4:V18", [f"c{i}" for i in range(19)] + ["Enabled A", "Enabled B"]
        )
        findings = _table_findings(baseline, current)
        target = next(
            f for f in findings if f.baseline_value and "B4:T18" in f.baseline_value
        )
        assert not target.expected_growth
        assert "column additions or removals" in target.message
        columns_event = next(
            f for f in findings if "columns added" in f.message
        )
        assert columns_event.event_key == target.event_key

    def test_geometry_only_row_extension_is_warning(self) -> None:
        columns = [f"c{i}" for i in range(5)]
        baseline = _table_workbook("B4:F18", columns)
        current = _table_workbook("B4:F25", columns)
        findings = _table_findings(baseline, current)
        assert len(findings) == 1
        assert findings[0].expected_reason is None
        assert assign_severity(findings[0]) is Severity.WARNING
        assert "without proved cadence evidence" in findings[0].message

    def test_later_period_rows_prove_cadence_extension(self) -> None:
        columns = [f"c{i}" for i in range(5)]
        baseline_cells = {
            (row, 2): CellRecord(row, 2, f"2026-{row - 4:02d}")
            for row in range(5, 9)
        }
        current_cells = {
            **baseline_cells,
            (9, 2): CellRecord(9, 2, "2026-05"),
            (10, 2): CellRecord(10, 2, "2026-06"),
        }
        baseline = _table_workbook("B4:F8", columns, cells=baseline_cells)
        current = _table_workbook("B4:F10", columns, cells=current_cells)

        finding = _table_findings(baseline, current)[0]

        assert finding.expected_reason is FindingExpectedReason.CADENCE_EXTENSION
        assert assign_severity(finding) is Severity.EXPECTED
        assert "proved later-period rows" in finding.message

    def test_profile_refresh_range_explicitly_accepts_row_extension(self) -> None:
        columns = [f"c{i}" for i in range(5)]
        baseline = _table_workbook("B4:F18", columns)
        current = _table_workbook("B4:F25", columns)
        profile = DeliverableProfile.model_validate(
            {
                "name": "table-refresh",
                "excel": {
                    "sheets": {"Data": {"refresh_ranges": ["B19:F25"]}}
                },
            }
        )

        finding = _table_findings(baseline, current, profile)[0]

        assert finding.expected_reason is FindingExpectedReason.PROFILE_REFRESH
        assert assign_severity(finding) is Severity.EXPECTED

    def test_sideways_extension_without_column_metadata_change_not_expected(
        self,
    ) -> None:
        # Same column names but a wider range is still not cadence growth.
        columns = [f"c{i}" for i in range(5)]
        baseline = _table_workbook("B4:F18", columns)
        current = _table_workbook("B4:H18", columns)
        findings = _table_findings(baseline, current)
        assert len(findings) == 1
        assert not findings[0].expected_growth


def _sheet(
    cells: dict[tuple[int, int], tuple[CellValue, str | None]], rows: int, cols: int
) -> SheetSnapshot:
    records = {
        (row, col): CellRecord(
            row, col, value, formula=formula, is_formula=formula is not None
        )
        for (row, col), (value, formula) in cells.items()
    }
    return SheetSnapshot("Panel", "visible", rows, cols, records)


class TestBlockAlignment:
    def test_kpi_panel_with_derived_labels_aligns_positionally(self) -> None:
        """The proven case: label + advancing derived date + one new line must
        never produce delete/growth churn."""
        baseline_sheet = _sheet(
            {
                (5, 7): ("Total Units Through:", '="Total "&INDEX(X,1)'),
                (6, 7): (dt.datetime(2026, 7, 10), "=INDEX(Y,1)"),
            },
            rows=8,
            cols=8,
        )
        current_sheet = _sheet(
            {
                (5, 7): ("Total Units Through:", '="Total "&INDEX(X,1)'),
                (6, 7): (dt.datetime(2026, 7, 17), "=IF(Z,INDEX(Y,1),0)"),
                (7, 7): (None, '=IF(Q,"","note")'),
            },
            rows=8,
            cols=8,
        )
        base_region = TableRegion("Panel", 5, 7, 6, 7, "block", None, 7, "none")
        curr_region = TableRegion("Panel", 5, 7, 7, 7, "block", None, 7, "none")

        alignment = align_regions(
            baseline_sheet, current_sheet, base_region, curr_region
        )

        assert alignment.rows.method == "positional"
        assert alignment.rows.pairs == [(5, 5), (6, 6)]
        assert alignment.rows.deleted == []
        assert alignment.rows.growth == []
        assert alignment.rows.inserted == [7]

    def test_stable_constant_labels_still_key_align(self) -> None:
        baseline_sheet = _sheet(
            {
                (2, 1): ("Alpha", None),
                (3, 1): ("Beta", None),
                (4, 1): ("Gamma", None),
                (2, 2): (1.0, None),
                (3, 2): (2.0, None),
                (4, 2): (3.0, None),
            },
            rows=5,
            cols=2,
        )
        current_sheet = _sheet(
            {
                (2, 1): ("Alpha", None),
                (3, 1): ("Gamma", None),
                (4, 1): ("Beta", None),
                (2, 2): (1.0, None),
                (3, 2): (3.0, None),
                (4, 2): (2.0, None),
            },
            rows=5,
            cols=2,
        )
        region_b = TableRegion("Panel", 2, 1, 4, 2, "block", None, 1, "none")
        region_c = TableRegion("Panel", 2, 1, 4, 2, "block", None, 1, "none")

        alignment = align_regions(baseline_sheet, current_sheet, region_b, region_c)

        assert alignment.rows.method == "keys"
        assert (3, 4) in alignment.rows.pairs  # Beta followed its label
        assert (4, 3) in alignment.rows.pairs  # Gamma followed its label

    @staticmethod
    def _label_sheet(labels: list[str]) -> SheetSnapshot:
        cells: dict[tuple[int, int], tuple[CellValue, str | None]] = {
            (row, 1): (label, None)
            for row, label in enumerate(labels, start=2)
        }
        cells.update(
            {
                (row, 2): (float(row), None)
                for row in range(2, len(labels) + 2)
            }
        )
        return _sheet(cells, rows=len(labels) + 2, cols=2)

    def test_stable_but_disjoint_labels_align_positionally(self) -> None:
        baseline = self._label_sheet(["Alpha", "Beta", "Gamma", "Delta"])
        current = self._label_sheet(["East", "West", "North", "South"])
        region = TableRegion("Panel", 2, 1, 5, 2, "block", None, 1, "none")

        alignment = align_regions(baseline, current, region, region)

        assert alignment.rows.method == "positional"
        assert not alignment.rows.low_confidence_fallback
        assert alignment.rows.pairs == [(2, 2), (3, 3), (4, 4), (5, 5)]

    def test_eighty_percent_cross_side_overlap_keeps_key_alignment(self) -> None:
        baseline = self._label_sheet(["A", "B", "C", "D", "E"])
        current = self._label_sheet(["A", "B", "C", "D", "F"])
        region = TableRegion("Panel", 2, 1, 6, 2, "block", None, 1, "none")

        alignment = align_regions(baseline, current, region, region)

        assert alignment.rows.method == "keys"
        assert len(alignment.rows.pairs) == 4
        assert alignment.rows.deleted == [6]
        assert alignment.rows.inserted == [6]

    def test_duplicate_block_labels_force_positional_alignment(self) -> None:
        baseline = self._label_sheet(["A", "A", "B", "C"])
        current = self._label_sheet(["A", "A", "B", "C"])
        region = TableRegion("Panel", 2, 1, 5, 2, "block", None, 1, "none")

        alignment = align_regions(baseline, current, region, region)

        assert alignment.rows.method == "positional"
        assert alignment.rows.pairs == [(2, 2), (3, 3), (4, 4), (5, 5)]


class TestKeyColumnVolatility:
    def _long_sheet(self, *, date_offset_days: int) -> SheetSnapshot:
        cells: dict[tuple[int, int], tuple[CellValue, str | None]] = {
            (1, 1): ("Brand", None),
            (1, 2): ("Data Through", None),
            (1, 3): ("Value", None),
        }
        for offset, brand in enumerate(("Alpha", "Beta", "Gamma", "Delta"), start=2):
            cells[(offset, 1)] = (brand, "=INDEX(tbl_S[Brand],1)")
            cells[(offset, 2)] = (
                dt.datetime(2026, 7, 10) + dt.timedelta(days=date_offset_days),
                "=INDEX(tbl_S[Through],1)",
            )
            cells[(offset, 3)] = (float(offset), None)
        return _sheet(cells, rows=6, cols=3)

    def test_derived_date_column_never_joins_the_composite_key(self) -> None:
        baseline_sheet = self._long_sheet(date_offset_days=0)
        current_sheet = self._long_sheet(date_offset_days=7)  # every date advanced
        region = TableRegion("Panel", 1, 1, 5, 3, "long", 1, 1, "rows")

        alignment = align_regions(baseline_sheet, current_sheet, region, region)

        assert alignment.rows.pairs == [(2, 2), (3, 3), (4, 4), (5, 5)]
        assert alignment.rows.deleted == []
        assert alignment.rows.inserted == []
        assert alignment.rows.growth == []

    def test_constant_period_axis_remains_the_key(self) -> None:
        def sheet(weeks: int) -> SheetSnapshot:
            cells: dict[tuple[int, int], tuple[CellValue, str | None]] = {
                (1, 1): ("Week", None),
                (1, 2): ("Value", None),
            }
            for offset in range(weeks):
                cells[(2 + offset, 1)] = (
                    dt.datetime(2026, 1, 2) + dt.timedelta(weeks=offset),
                    None,  # constant axis dates
                )
                cells[(2 + offset, 2)] = (float(offset), None)
            return _sheet(cells, rows=2 + weeks, cols=2)

        baseline_sheet = sheet(4)
        current_sheet = sheet(5)
        base_region = TableRegion("Panel", 1, 1, 5, 2, "long", 1, 1, "rows")
        curr_region = TableRegion("Panel", 1, 1, 6, 2, "long", 1, 1, "rows")

        alignment = align_regions(
            baseline_sheet, current_sheet, base_region, curr_region
        )

        assert len(alignment.rows.pairs) == 4
        assert alignment.rows.growth == [6]  # the new trailing week
        assert alignment.rows.deleted == []


class TestRefreshBlockTextGuard:
    def _findings(
        self,
        *,
        baseline_value: float = 100.0,
        current_value: float = 140.0,
        refresh_ranges: list[str] | None = None,
    ) -> dict[str, Finding]:
        baseline_sheet = _sheet(
            {
                (1, 1): ("Cycle", None),
                (1, 2): ("Jun-26", None),
                (2, 1): ("Revenue", None),
                (2, 2): (baseline_value, None),
                (3, 1): ("Mode", None),
                (3, 2): ("Yes", None),
            },
            rows=3,
            cols=2,
        )
        current_sheet = _sheet(
            {
                (1, 1): ("Cycle", None),
                (1, 2): ("Jul-26", None),
                (2, 1): ("Revenue", None),
                (2, 2): (current_value, None),
                (3, 1): ("Mode", None),
                (3, 2): ("Units Only", None),
            },
            rows=3,
            cols=2,
        )
        region = TableRegion("Panel", 1, 1, 3, 2, "block", None, 1, "none")
        alignment = align_regions(baseline_sheet, current_sheet, region, region)
        findings = diff_region_values(
            baseline_sheet,
            current_sheet,
            alignment,
            NumericTolerance(),
            ignore=_RangeSet([]),
            refresh=_RangeSet(refresh_ranges or []),
            sheet_profile=None,
        )
        return {
            f.location or "": f
            for f in findings
            if f.finding_class is FindingClass.VALUE_CHANGED
        }

    def test_implicit_numeric_refresh_is_warning_but_text_flip_is_critical(
        self,
    ) -> None:
        by_location = self._findings()

        assert by_location["B1"].expected_reason is (
            FindingExpectedReason.PERIOD_PROGRESSION
        )
        assert assign_severity(by_location["B1"]) is Severity.EXPECTED
        assert not by_location["B2"].expected_growth
        assert by_location["B2"].temporal_context is (
            FindingTemporalContext.CURRENT_PERIOD
        )
        assert assign_severity(by_location["B2"]) is Severity.WARNING
        assert not by_location["B3"].expected_growth
        assert assign_severity(by_location["B3"]) is Severity.CRITICAL
        assert "block refresh" not in by_location["B3"].message

    @pytest.mark.parametrize(
        ("baseline_value", "current_value"),
        [
            (100.0, 1_000_000_000.0),
            (100.0, -100.0),
            (0.0, 1.0),
            (1.0, 10.0),
            (100.0, 10.0),
        ],
    )
    def test_implicit_refresh_hard_anomalies_remain_critical(
        self, baseline_value: float, current_value: float
    ) -> None:
        finding = self._findings(
            baseline_value=baseline_value,
            current_value=current_value,
        )["B2"]

        assert finding.expected_reason is None
        assert finding.temporal_context is None
        assert assign_severity(finding) is Severity.CRITICAL
        assert "anomalous" in finding.message

    def test_profile_refresh_remains_expected_even_for_a_hard_anomaly(self) -> None:
        finding = self._findings(
            baseline_value=100.0,
            current_value=1_000_000_000.0,
            refresh_ranges=["B2"],
        )["B2"]

        assert finding.expected_reason is FindingExpectedReason.PROFILE_REFRESH
        assert assign_severity(finding) is Severity.EXPECTED
