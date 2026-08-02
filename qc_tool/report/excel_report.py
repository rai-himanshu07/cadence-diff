"""Annotated Excel diff report: summary sheet + filterable findings sheet.

Consumes the triaged `QCRunResult` only — no diff logic here. The output
is a brand-new workbook; sources are never touched.
"""

import datetime as dt
import logging
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell.cell import Cell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.worksheet import Worksheet

from qc_tool.coverage import capability_limited
from qc_tool.engine import QCRunResult
from qc_tool.findings import Severity
from qc_tool.review import build_pattern_groups, count_pattern_groups, format_group_ranges
from qc_tool.security import private_directory, private_file
from qc_tool.story import build_stories

logger = logging.getLogger(__name__)

_SEVERITY_FILLS = {
    Severity.CRITICAL: PatternFill(fill_type="solid", fgColor="FFC00000"),
    Severity.WARNING: PatternFill(fill_type="solid", fgColor="FFED7D31"),
    Severity.INFO: PatternFill(fill_type="solid", fgColor="FF2E75B6"),
    Severity.EXPECTED: PatternFill(fill_type="solid", fgColor="FF548235"),
}
_WHITE_BOLD = Font(bold=True, color="FFFFFFFF")
_HEADER_FILL = PatternFill(fill_type="solid", fgColor="FF262626")
_ReportCellValue = str | int | float | bool | dt.date | dt.datetime | None

_COLUMNS = [
    ("ID", 8),
    ("Severity", 12),
    ("Artifact", 10),
    ("Class", 24),
    ("Sheet / Slide", 20),
    ("Location", 14),
    ("Element", 22),
    ("Provenance", 18),
    ("Subtype", 20),
    ("Materiality", 18),
    ("Temporal context", 20),
    ("Expected reason", 22),
    ("Evidence tags", 50),
    ("Baseline", 32),
    ("Current", 32),
    ("Message", 60),
    ("Downstream impact", 40),
    ("Analyst comment", 36),
    ("Root cause", 36),
    ("Waiver", 50),
]


def _dynamic_cell(
    sheet: Worksheet, *, row: int, column: int, value: _ReportCellValue
) -> Cell:
    """Write dynamic content as data, never as an executable spreadsheet formula."""
    cell = sheet.cell(row=row, column=column, value=value)
    if isinstance(value, str):
        cell.data_type = "s"
    return cell


def _internal_link(
    cell: Cell,
    *,
    location: str,
    label: str,
    tooltip: str | None = None,
) -> None:
    """Create workbook-local navigation without an executable formula."""
    cell.value = label
    cell.data_type = "s"
    cell.hyperlink = Hyperlink(
        ref=cell.coordinate,
        location=location,
        display=label,
        tooltip=tooltip,
    )
    cell.style = "Hyperlink"


def write_excel_report(result: QCRunResult, path: Path) -> None:
    review_groups = build_pattern_groups(result.findings)
    counts = count_pattern_groups(review_groups)
    stories = build_stories(result.findings)
    workbook = Workbook()
    summary = workbook.active
    if summary is None:  # pragma: no cover - openpyxl always provides one
        raise RuntimeError("openpyxl workbook has no active sheet")
    summary.title = "Summary"

    summary["A1"] = "QC Report"
    summary["A1"].font = Font(bold=True, size=16)
    rows: list[tuple[str, str]] = [
        ("Generated (UTC)", dt.datetime.now(dt.UTC).isoformat(timespec="seconds")),
        ("Profile", result.profile_name),
        ("Mode", result.mode.value),
        *((f"File: {role}", name) for role, name in result.files.items()),
        (
            "Result status",
            "capability-limited"
            if capability_limited(result.coverage)
            else "all required checks ran",
        ),
        ("Verified cross-checks", str(result.verified_crosschecks)),
        ("Change stories", str(len(stories))),
        *(
            (f"{severity.value.title()} pattern review items", str(count))
            for severity, count in counts.review_items.items()
        ),
        *(
            (f"{severity.value.title()} atomic findings", str(count))
            for severity, count in counts.atomic_findings.items()
        ),
    ]
    for offset, (label, value) in enumerate(rows, start=3):
        _dynamic_cell(summary, row=offset, column=1, value=label).font = Font(bold=True)
        _dynamic_cell(summary, row=offset, column=2, value=value)
    if result.mapping_coverage is not None:
        mapping = result.mapping_coverage
        mapping_rows = (
            ("Eligible PPT figures", mapping.eligible),
            ("Mapped figures", mapping.mapped),
            ("Verified figures", mapping.verified),
            ("Mismatched figures", mapping.mismatched),
            ("Unresolved mappings", mapping.unresolved),
            ("Unmapped figures", mapping.unmapped),
        )
        start = 3 + len(rows)
        for offset, (label, value) in enumerate(mapping_rows, start=start):
            _dynamic_cell(summary, row=offset, column=1, value=label).font = Font(bold=True)
            _dynamic_cell(summary, row=offset, column=2, value=value)
    disclosure_row = 3 + len(rows) + 1
    for offset, disclosure in enumerate(result.disclosures):
        cell = _dynamic_cell(
            summary,
            row=disclosure_row + offset,
            column=1,
            value=f"NOTE: {disclosure}",
        )
        cell.font = Font(italic=True, color="FFC00000")
    summary.column_dimensions["A"].width = 28
    summary.column_dimensions["B"].width = 60
    _internal_link(summary["D3"], location="'Stories'!A1", label="Change stories")
    _internal_link(summary["D4"], location="'Review Groups'!A1", label="Review groups")
    _internal_link(summary["D5"], location="'Findings'!A1", label="Atomic findings")
    _internal_link(summary["D6"], location="'Coverage'!A1", label="Coverage")
    summary.column_dimensions["D"].width = 20

    stories_sheet = workbook.create_sheet("Stories")
    story_headers = (
        "Story",
        "Kind",
        "Title",
        "Narrative and evidence",
        "Findings",
        "Severity mix",
    )
    for col, header in enumerate(story_headers, start=1):
        cell = stories_sheet.cell(row=1, column=col, value=header)
        cell.font = _WHITE_BOLD
        cell.fill = _HEADER_FILL
    for row, story in enumerate(stories, start=2):
        narrative = story.description
        if story.evidence:
            narrative += "\n" + "\n".join(story.evidence)
        mix = " | ".join(
            f"{name}: {count}"
            for name, count in story.severity_counts.items()
            if count
        )
        values = (
            story.story_id,
            story.kind.value,
            story.title,
            narrative,
            story.member_count,
            mix,
        )
        for col, value in enumerate(values, start=1):
            _dynamic_cell(stories_sheet, row=row, column=col, value=value)
    stories_sheet.freeze_panes = "A2"
    stories_sheet.column_dimensions["C"].width = 44
    stories_sheet.column_dimensions["D"].width = 80

    coverage = workbook.create_sheet("Coverage")
    coverage_headers = ("Artifact", "Check", "Status", "Findings", "Detail")
    for col, header in enumerate(coverage_headers, start=1):
        cell = coverage.cell(row=1, column=col, value=header)
        cell.font = _WHITE_BOLD
        cell.fill = _HEADER_FILL
    for row, item in enumerate(result.coverage, start=2):
        values = (item.artifact, item.label, item.state.value, item.findings, item.detail)
        for col, value in enumerate(values, start=1):
            _dynamic_cell(coverage, row=row, column=col, value=value)
    coverage.freeze_panes = "A2"
    for col, width in enumerate((14, 34, 14, 12, 60), start=1):
        coverage.column_dimensions[get_column_letter(col)].width = width

    if result.mapping_suggestions:
        suggestions = workbook.create_sheet("Mapping Suggestions")
        headers = ("Slide", "Figure", "Context", "Candidate", "Value", "Display match")
        for col, header in enumerate(headers, start=1):
            cell = suggestions.cell(row=1, column=col, value=header)
            cell.font = _WHITE_BOLD
            cell.fill = _HEADER_FILL
        row = 2
        for suggestion in result.mapping_suggestions:
            if not suggestion.candidates:
                values = (
                    suggestion.slide,
                    suggestion.figure_raw,
                    suggestion.line,
                    "",
                    "",
                    "",
                )
                for column, value in enumerate(values, start=1):
                    _dynamic_cell(suggestions, row=row, column=column, value=value)
                row += 1
                continue
            for candidate in suggestion.candidates:
                values = (
                    suggestion.slide,
                    suggestion.figure_raw,
                    suggestion.line,
                    f"{candidate.sheet}!{candidate.cell}",
                    candidate.value,
                    candidate.display_match,
                )
                for column, value in enumerate(values, start=1):
                    _dynamic_cell(suggestions, row=row, column=column, value=value)
                row += 1
        for col, width in enumerate((24, 14, 50, 24, 18, 14), start=1):
            suggestions.column_dimensions[get_column_letter(col)].width = width

    review_sheet = workbook.create_sheet("Review Groups")
    review_headers = (
        "Group",
        "Severity",
        "Class",
        "Sheet / Slide",
        "Affected range",
        "Findings",
        "Baseline range",
        "Element",
        "Review item",
        "Detail",
    )
    review_widths = (18, 12, 25, 20, 28, 12, 28, 24, 58, 20)
    for column, (header, width) in enumerate(
        zip(review_headers, review_widths, strict=True), start=1
    ):
        cell = review_sheet.cell(row=1, column=column, value=header)
        cell.font = _WHITE_BOLD
        cell.fill = _HEADER_FILL
        review_sheet.column_dimensions[get_column_letter(column)].width = width

    sheet = workbook.create_sheet("Findings")
    for col, (header, width) in enumerate(_COLUMNS, start=1):
        cell = sheet.cell(row=1, column=col, value=header)
        cell.font = _WHITE_BOLD
        cell.fill = _HEADER_FILL
        sheet.column_dimensions[get_column_letter(col)].width = width
    finding_rows: dict[str, int] = {}
    for row, finding in enumerate(result.findings, start=2):
        finding_rows[finding.finding_id] = row
        severity = finding.severity or Severity.WARNING
        severity_label = severity.value + (" *" if finding.severity_overridden else "")
        values = [
            finding.finding_id,
            severity_label,
            finding.artifact,
            finding.finding_class.value,
            finding.sheet or finding.slide or "",
            finding.location or finding.baseline_location or "",
            finding.element or "",
            finding.provenance.value if finding.provenance is not None else "",
            finding.subtype.value if finding.subtype is not None else "",
            finding.materiality.value if finding.materiality is not None else "",
            (
                finding.temporal_context.value
                if finding.temporal_context is not None
                else ""
            ),
            (
                finding.expected_reason.value
                if finding.expected_reason is not None
                else ""
            ),
            "; ".join(sorted(tag.value for tag in finding.evidence_tags)),
            finding.baseline_value or "",
            finding.current_value or "",
            finding.message,
            "; ".join(finding.impacts),
            finding.analyst_comment,
            finding.root_cause_key,
            (
                f"{finding.waiver_reason} (expires {finding.waiver_expires})"
                if finding.waiver_reason
                else ""
            ),
        ]
        for col, value in enumerate(values, start=1):
            cell = _dynamic_cell(sheet, row=row, column=col, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=col >= 8)
        severity_cell = sheet.cell(row=row, column=2)
        severity_cell.fill = _SEVERITY_FILLS[severity]
        severity_cell.font = _WHITE_BOLD
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(_COLUMNS))}{len(result.findings) + 1}"
    if any(f.severity_overridden for f in result.findings):
        note_row = len(result.findings) + 3
        note = sheet.cell(
            row=note_row, column=1, value="* severity manually set by the analyst"
        )
        note.font = Font(italic=True)

    for row, group in enumerate(review_groups, start=2):
        values = (
            group.group_id,
            group.severity.value,
            group.finding_class.value,
            group.sheet or group.slide or "",
            format_group_ranges(group),
            group.member_count,
            "; ".join(group.baseline_ranges),
            group.element,
            (
                group.members[0].message
                if group.member_count == 1
                else (
                    f"{group.member_count:,} contiguous "
                    f"{group.finding_class.value.replace('_', ' ')} findings"
                )
            ),
            "",
        )
        for column, value in enumerate(values, start=1):
            cell = _dynamic_cell(review_sheet, row=row, column=column, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=column in {5, 7, 9})
        severity_cell = review_sheet.cell(row=row, column=2)
        severity_cell.fill = _SEVERITY_FILLS[group.severity]
        severity_cell.font = _WHITE_BOLD
        first_row = finding_rows[group.members[0].finding_id]
        _internal_link(
            review_sheet.cell(row=row, column=10),
            location=f"'Findings'!A{first_row}",
            label="View atomic detail",
            tooltip="Clear any Findings sheet filter if the target row is hidden.",
        )
    review_sheet.freeze_panes = "A2"
    review_sheet.auto_filter.ref = (
        f"A1:{get_column_letter(len(review_headers))}{len(review_groups) + 1}"
    )

    private_directory(path.parent)
    workbook.save(path)
    private_file(path)
    logger.info(
        "Excel report written to %s (%d review items, %d findings)",
        path,
        len(review_groups),
        len(result.findings),
    )
