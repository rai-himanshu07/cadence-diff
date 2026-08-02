"""Builds the baseline/current Excel fixture pair with seeded defects.

Layout summary (both cycles):

- ``Long_Monthly`` — long-format table, one row per (month, region);
  Margin column E carries formulas ``=Cn-Dn``. Growth = new rows.
- ``Wide_Weekly`` — wide-format table, one column per week; Margin row 4
  carries formulas. Growth = new column; current also *deletes* W03.
- ``Dashboard`` — multi-block sheet: KPI constants block (cross-check
  anchors), headcount block, and a line chart over Long_Monthly.
- ``Summary`` — cross-sheet formulas (dependency-graph target); current
  gains seeded error values.
- ``Old_Sheet`` (baseline only), ``New_Analysis`` (current only),
  ``Params`` (visible in baseline, hidden in current).

Pivot parts are injected as raw XML because openpyxl cannot author pivots;
the loader contract is to scan ``xl/pivotTables/*.xml`` and
``xl/pivotCache/pivotCacheDefinition*.xml`` inside the package.
"""

import datetime as dt
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.worksheet import Worksheet

from tests.fixtures import domain
from tests.fixtures.manifest_schema import (
    Artifact,
    DefectClass,
    ExpectedChange,
    SeededDefect,
)

FIXED_DOC_TIME = dt.datetime(2026, 7, 1, tzinfo=dt.UTC)

BASELINE_FILL = PatternFill(fill_type="solid", fgColor="FFFFFF00")
DEFECT_FILL = PatternFill(fill_type="solid", fgColor="FFFF0000")

PIVOT_CACHE_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<cacheSource type="worksheet"><worksheetSource ref="{ref}" sheet="Long_Monthly"/>'
    '</cacheSource><cacheFields count="0"/></pivotCacheDefinition>'
)

PIVOT_TABLE_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    ' name="RevenuePivot" cacheId="1"><location ref="A3:C10" firstHeaderRow="1"'
    ' firstDataRow="2" firstDataCol="1"/></pivotTableDefinition>'
)


def weekly_labels(*, current: bool) -> list[str]:
    if current:
        return [
            w for w in domain.WEEK_LABELS[: domain.CURRENT_WEEKS] if w != domain.DELETED_WEEK
        ]
    return domain.WEEK_LABELS[: domain.BASELINE_WEEKS]


def week_cell(label: str, row: int, *, current: bool) -> str:
    """A1 address of a week's cell in Wide_Weekly for the given cycle."""
    col = 2 + weekly_labels(current=current).index(label)
    return f"{get_column_letter(col)}{row}"


def _build_long_monthly(ws: Worksheet, *, current: bool) -> None:
    for col, header in enumerate(["Period", "Region", "Revenue", "Cost", "Margin"], start=1):
        ws.cell(row=1, column=col, value=header)
    months = domain.CURRENT_MONTHS if current else domain.BASELINE_MONTHS
    revenue_fn = domain.current_monthly_revenue if current else domain.monthly_revenue
    cost_fn = domain.current_monthly_cost if current else domain.monthly_cost
    row = 2
    for month in range(months):
        for region in range(len(domain.REGION_LABELS)):
            ws.cell(row=row, column=1, value=domain.MONTH_LABELS[month])
            ws.cell(row=row, column=2, value=domain.REGION_LABELS[region])
            ws.cell(row=row, column=3, value=revenue_fn(month, region))
            ws.cell(row=row, column=4, value=cost_fn(month, region))
            ws.cell(row=row, column=5, value=f"=C{row}-D{row}")
            row += 1
    if current:
        # E02: formula overwritten with its computed constant (month 2, North).
        ws["E10"] = domain.monthly_revenue(2, 0) - domain.monthly_cost(2, 0)
        # E03: formula logic changed.
        ws["E14"] = "=C14-D14*1.1"
        # E04: formula not extended into the last new row.
        ws["E25"] = None


def _build_wide_weekly(ws: Worksheet, *, current: bool) -> None:
    labels = weekly_labels(current=current)
    ws.cell(row=1, column=1, value="Metric")
    ws.cell(row=2, column=1, value="Revenue")
    ws.cell(row=3, column=1, value="Cost")
    ws.cell(row=4, column=1, value="Margin")
    for offset, label in enumerate(labels):
        col = 2 + offset
        week_idx = int(label[1:]) - 1
        letter = get_column_letter(col)
        ws.cell(row=1, column=col, value=label)
        revenue = domain.weekly_revenue(week_idx)
        if current and label == "W05":
            revenue += domain.E06_DELTA  # E06: historical value edited
        ws.cell(row=2, column=col, value=revenue)
        ws.cell(row=3, column=col, value=domain.weekly_cost(week_idx))
        formula = f"={letter}2-{letter}3"
        if current and label == "W15":
            formula = f"={letter}2-{letter}3*2"  # E08: inconsistent within range
        ws.cell(row=4, column=col, value=formula)


def _build_dashboard(ws: Worksheet, ws_long: Worksheet, *, current: bool) -> None:
    months = domain.CURRENT_MONTHS if current else domain.BASELINE_MONTHS
    last_data_row = 1 + months * len(domain.REGION_LABELS)

    ws["A1"] = "KPI Summary"
    ws["A2"] = "Total Revenue"
    ws["A2"].fill = DEFECT_FILL if current else BASELINE_FILL  # E10: style changed
    ws["B2"] = domain.total_revenue(current=current)
    ws["A3"] = "Total Cost"
    ws["B3"] = domain.total_cost(current=current)
    ws["A4"] = "Margin %"
    ws["B4"] = domain.margin_ratio(current=current)
    ws["B4"].number_format = "0.00" if current else "0.0%"  # E09: format changed
    ws["A5"] = "Cycle"
    ws["B5"] = domain.MONTH_LABELS[months - 1]

    ws["E2"] = "Headcount"
    for month in range(months):
        ws.cell(row=2, column=6 + month, value=domain.MONTH_LABELS[month])
    for region_idx, region in enumerate(domain.HEADCOUNT_REGIONS):
        ws.cell(row=3 + region_idx, column=5, value=region)
        for month in range(months):
            ws.cell(row=3 + region_idx, column=6 + month, value=domain.headcount(month, region_idx))

    chart = LineChart()
    chart.title = "Revenue trend (workbook)"
    if current:
        # E16: series repointed from Revenue (C) to Cost (D).
        data = Reference(ws_long, min_col=4, min_row=2, max_row=last_data_row)
        cats = Reference(ws_long, min_col=1, min_row=2, max_row=last_data_row)
    else:
        data = Reference(ws_long, min_col=3, min_row=2, max_row=last_data_row)
        cats = Reference(ws_long, min_col=1, min_row=2, max_row=last_data_row)
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)
    ws.add_chart(chart, "E8")


def _build_summary(ws: Worksheet, *, current: bool) -> None:
    last = 1 + (domain.CURRENT_MONTHS if current else domain.BASELINE_MONTHS) * len(
        domain.REGION_LABELS
    )
    ws["A2"] = "Revenue total"
    ws["B2"] = f"=SUM(Long_Monthly!C2:C{last})"
    ws["A3"] = "Cost total"
    ws["B3"] = f"=SUM(Long_Monthly!D2:D{last})"
    ws["A4"] = "Margin"
    ws["B4"] = "=B2-B3"
    ws["A5"] = "Margin (E col)"
    ws["B5"] = f"=SUM(Long_Monthly!E2:E{last})"
    ws["A6"] = "Ratio"
    ws["B6"] = "=B4/B2"
    if current:
        ws["C5"] = "#REF!"  # E05: error constant
        ws["D5"] = "=SUM(#REF!)"  # E13: error inside formula text


def _build_workbook(*, current: bool) -> Workbook:
    wb = Workbook()
    ws_long = wb.active
    if ws_long is None:  # pragma: no cover - openpyxl always provides one
        raise RuntimeError("openpyxl workbook has no active sheet")
    ws_long.title = "Long_Monthly"
    _build_long_monthly(ws_long, current=current)

    _build_wide_weekly(wb.create_sheet("Wide_Weekly"), current=current)
    _build_dashboard(wb.create_sheet("Dashboard"), ws_long, current=current)
    _build_summary(wb.create_sheet("Summary"), current=current)

    if current:
        new_ws = wb.create_sheet("New_Analysis")  # E15
        new_ws["A1"] = "New analysis"
        new_ws["B1"] = 456.0
    else:
        old_ws = wb.create_sheet("Old_Sheet")  # E14 (removed in current)
        old_ws["A1"] = "Archived metric"
        old_ws["B1"] = 123.0

    params = wb.create_sheet("Params")
    params["A1"] = "Threshold"
    params["B1"] = 0.05
    params.sheet_state = "hidden" if current else "visible"  # E11

    last = 1 + (domain.CURRENT_MONTHS if current else domain.BASELINE_MONTHS) * len(
        domain.REGION_LABELS
    )
    wb.defined_names["RevenueData"] = DefinedName(
        "RevenueData", attr_text=f"Long_Monthly!$A$1:$E${last}"
    )
    kpi_target = "Dashboard!$B$5" if current else "Dashboard!$B$4"  # E12
    wb.defined_names["KPI_Margin"] = DefinedName("KPI_Margin", attr_text=kpi_target)

    wb.properties.created = FIXED_DOC_TIME
    wb.properties.modified = FIXED_DOC_TIME
    return wb


def pivot_parts(*, current: bool) -> dict[str, bytes]:
    """Raw pivot parts to inject into the package (E17 seeds a shrunk source)."""
    ref = "A1:D21" if current else "A1:E21"
    return {
        "xl/pivotCache/pivotCacheDefinition1.xml": PIVOT_CACHE_TEMPLATE.format(ref=ref).encode(
            "utf-8"
        ),
        "xl/pivotTables/pivotTable1.xml": PIVOT_TABLE_XML.encode("utf-8"),
    }


PIVOT_CONTENT_TYPES = {
    "/xl/pivotCache/pivotCacheDefinition1.xml": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
    ),
    "/xl/pivotTables/pivotTable1.xml": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
    ),
}


def build_workbooks(dest: Path) -> tuple[list[SeededDefect], list[ExpectedChange]]:
    """Write baseline.xlsx and current.xlsx; return seeded ground truth."""
    _build_workbook(current=False).save(dest / "baseline.xlsx")
    _build_workbook(current=True).save(dest / "current.xlsx")

    def defect(defect_id: str, classes: list[DefectClass], **kwargs: Any) -> SeededDefect:
        return SeededDefect(
            defect_id=defect_id, artifact=Artifact.EXCEL, classes=classes, **kwargs
        )

    defects = [
        defect(
            "E01",
            [DefectClass.VALUE_CHANGED],
            sheet="Long_Monthly",
            cell="C7",
            baseline=str(domain.monthly_revenue(1, 1)),
            current=str(domain.monthly_revenue(1, 1) + domain.E01_DELTA),
            impacts=["Summary!B2", "Summary!B4", "Summary!B6"],
            note="historical revenue edited (Feb-26 / South)",
        ),
        defect(
            "E19",
            [DefectClass.VALUE_CHANGED],
            sheet="Long_Monthly",
            cell="D4",
            baseline=str(domain.monthly_cost(0, 2)),
            current=str(domain.current_monthly_cost(0, 2)),
            impacts=["Summary!B3", "Summary!B4", "Summary!B6"],
            note="old-history constant re-exported one ULP off: representation noise",
        ),
        defect(
            "E20",
            [DefectClass.VALUE_CHANGED],
            sheet="Long_Monthly",
            cell="C21",
            baseline=str(domain.monthly_revenue(4, 3)),
            current=str(domain.monthly_revenue(4, 3) + domain.E20_DELTA),
            impacts=["Summary!B2", "Summary!B4", "Summary!B6"],
            note="May-26 restated inside the trailing 2-month window",
        ),
        defect(
            "E02",
            [DefectClass.FORMULA_HARDCODED],
            sheet="Long_Monthly",
            cell="E10",
            baseline="=C10-D10",
            current=str(domain.monthly_revenue(2, 0) - domain.monthly_cost(2, 0)),
            impacts=["Summary!B5"],
            note="margin formula replaced by constant",
        ),
        defect(
            "E03",
            [DefectClass.FORMULA_LOGIC_CHANGED],
            sheet="Long_Monthly",
            cell="E14",
            baseline="=C14-D14",
            current="=C14-D14*1.1",
            impacts=["Summary!B5"],
            note="margin formula logic altered",
        ),
        defect(
            "E04",
            [DefectClass.FORMULA_NOT_EXTENDED],
            sheet="Long_Monthly",
            cell="E25",
            baseline="=C25-D25 expected by column pattern",
            current="",
            impacts=["Summary!B5"],
            note="margin formula missing in last new row",
        ),
        defect(
            "E05",
            [DefectClass.FORMULA_ERROR],
            sheet="Summary",
            cell="C5",
            current="#REF!",
            note="error constant",
        ),
        defect(
            "E13",
            [DefectClass.FORMULA_ERROR],
            sheet="Summary",
            cell="D5",
            current="=SUM(#REF!)",
            note="error reference inside formula text",
        ),
        defect(
            "E06",
            [DefectClass.VALUE_CHANGED],
            sheet="Wide_Weekly",
            cell=week_cell("W05", 2, current=True),
            baseline_cell=week_cell("W05", 2, current=False),
            baseline=str(domain.weekly_revenue(4)),
            current=str(domain.weekly_revenue(4) + domain.E06_DELTA),
            note="historical weekly revenue edited",
        ),
        defect(
            "E07",
            [DefectClass.COLUMN_DELETED],
            sheet="Wide_Weekly",
            baseline_cell=week_cell(domain.DELETED_WEEK, 1, current=False),
            baseline=domain.DELETED_WEEK,
            note="historical week column deleted in current",
        ),
        defect(
            "E08",
            [DefectClass.FORMULA_INCONSISTENT],
            sheet="Wide_Weekly",
            cell=week_cell("W15", 4, current=True),
            baseline=f"={week_cell('W15', 2, current=True)[:-1]}2"
            f"-{week_cell('W15', 2, current=True)[:-1]}3",
            current=f"={week_cell('W15', 2, current=True)[:-1]}2"
            f"-{week_cell('W15', 2, current=True)[:-1]}3*2",
            note="margin formula deviates from row pattern",
        ),
        defect(
            "E09",
            [DefectClass.NUMBER_FORMAT_CHANGED],
            sheet="Dashboard",
            cell="B4",
            baseline="0.0%",
            current="0.00",
            note="margin KPI display format broken",
        ),
        defect(
            "E18",
            [DefectClass.STYLE_CHANGED],
            sheet="Dashboard",
            cell="A2",
            baseline="FFFFFF00",
            current="FFFF0000",
            note="KPI label fill changed",
        ),
        defect(
            "E11",
            [DefectClass.HIDDEN_CHANGED],
            sheet="Params",
            baseline="visible",
            current="hidden",
            note="Params sheet hidden in current",
        ),
        defect(
            "E12",
            [DefectClass.NAMED_RANGE_CHANGED],
            element="KPI_Margin",
            baseline="Dashboard!$B$4",
            current="Dashboard!$B$5",
            note="named range repointed",
        ),
        defect(
            "E14",
            [DefectClass.SHEET_REMOVED],
            sheet="Old_Sheet",
            note="baseline-only sheet",
        ),
        defect(
            "E15",
            [DefectClass.SHEET_ADDED],
            sheet="New_Analysis",
            note="current-only sheet",
        ),
        defect(
            "E16",
            [DefectClass.CHART_SERIES_CHANGED],
            sheet="Dashboard",
            element="Revenue trend (workbook)",
            baseline="'Long_Monthly'!$C$2:$C$21",
            current="'Long_Monthly'!$D$2:$D$25",
            note="chart series repointed from Revenue to Cost",
        ),
        defect(
            "E17",
            [DefectClass.PIVOT_SOURCE_CHANGED],
            element="RevenuePivot",
            baseline="A1:E21",
            current="A1:D21",
            note="pivot source lost Margin column and was not extended",
        ),
    ]

    expected = [
        ExpectedChange(
            change_id="EX01",
            artifact=Artifact.EXCEL,
            kind="new_rows",
            sheet="Long_Monthly",
            detail="Jun-26 rows 22:25 appended",
        ),
        ExpectedChange(
            change_id="EX02",
            artifact=Artifact.EXCEL,
            kind="new_column",
            sheet="Wide_Weekly",
            detail="W21 column appended",
        ),
        ExpectedChange(
            change_id="EX03",
            artifact=Artifact.EXCEL,
            kind="aggregate_refresh",
            sheet="Dashboard",
            detail="KPI block B2:B5 reflects the new cycle totals",
        ),
        ExpectedChange(
            change_id="EX04",
            artifact=Artifact.EXCEL,
            kind="named_range_extended",
            detail="RevenueData extended $E$21 -> $E$25",
        ),
        ExpectedChange(
            change_id="EX05",
            artifact=Artifact.EXCEL,
            kind="new_column",
            sheet="Dashboard",
            detail="headcount Jun-26 column appended",
        ),
        ExpectedChange(
            change_id="EX06",
            artifact=Artifact.EXCEL,
            kind="formula_range_extension",
            sheet="Summary",
            detail="B2/B3/B5 SUM ranges extended row 21 -> 25; not a logic change",
        ),
    ]
    return defects, expected
