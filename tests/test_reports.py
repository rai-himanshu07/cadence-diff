"""Report generation tests (criterion 11)."""

from pathlib import Path

from openpyxl import load_workbook

from qc_tool.engine import QCRunResult, run_qc
from qc_tool.findings import Finding, FindingClass, Severity
from qc_tool.report.excel_report import write_excel_report
from qc_tool.report.html_report import render_html_report, write_html_report
from qc_tool.report.json_report import result_payload


def test_excel_report_structure(qc_result: QCRunResult, tmp_path: Path) -> None:
    path = tmp_path / "report.xlsx"
    write_excel_report(qc_result, path)

    workbook = load_workbook(path)
    assert workbook.sheetnames == [
        "Summary",
        "Stories",
        "Coverage",
        "Review Groups",
        "Findings",
    ]

    findings_sheet = workbook["Findings"]
    assert findings_sheet.max_row == len(qc_result.findings) + 1
    assert findings_sheet["A1"].value == "ID"
    assert findings_sheet["A2"].value == "F0001"
    assert findings_sheet["B2"].value == "critical"
    assert findings_sheet.auto_filter.ref is not None

    summary = workbook["Summary"]
    summary_text = " ".join(
        str(cell.value) for row in summary.iter_rows() for cell in row if cell.value
    )
    assert "fixture" in summary_text  # profile name
    assert "cycle_comparison" in summary_text
    assert "current.xlsx" in summary_text
    assert "Critical pattern review items" in summary_text
    assert "Critical atomic findings" in summary_text

    coverage = workbook["Coverage"]
    assert coverage["A1"].value == "Artifact"
    assert coverage.max_row == len(qc_result.coverage) + 1

    review_groups = workbook["Review Groups"]
    assert review_groups["A1"].value == "Group"
    assert review_groups.max_row <= len(qc_result.findings) + 1
    assert review_groups["J2"].hyperlink is not None
    assert review_groups["J2"].hyperlink.location.startswith("'Findings'!A")
    assert "Clear any Findings sheet filter" in review_groups["J2"].hyperlink.tooltip
    assert summary["D3"].hyperlink is not None
    assert summary["D3"].hyperlink.location == "'Stories'!A1"
    assert summary["D4"].hyperlink is not None
    assert summary["D4"].hyperlink.location == "'Review Groups'!A1"

    stories_sheet = workbook["Stories"]
    assert stories_sheet["A1"].value == "Story"
    assert stories_sheet.max_row >= 2  # at least one story for the fixture pair


def test_excel_report_discloses_xlsb_degradation(fixture_dir: Path, tmp_path: Path) -> None:
    result = run_qc(
        baseline_excel=fixture_dir / "baseline.xlsb",
        current_excel=fixture_dir / "current.xlsb",
    )
    path = tmp_path / "xlsb_report.xlsx"
    write_excel_report(result, path)
    summary = load_workbook(path)["Summary"]
    summary_text = " ".join(
        str(cell.value) for row in summary.iter_rows() for cell in row if cell.value
    )
    assert "NOTE:" in summary_text and "degraded" in summary_text


def test_excel_report_keeps_formula_text_inert(tmp_path: Path) -> None:
    result = QCRunResult(profile_name="test")
    result.findings = [
        Finding(
            artifact="excel",
            finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
            baseline_value='=HYPERLINK("https://example.invalid","open")',
            current_value="=1+1",
            message="formula changed",
            severity=Severity.WARNING,
        )
    ]
    path = tmp_path / "formula-text.xlsx"

    write_excel_report(result, path)

    workbook = load_workbook(path, data_only=False)
    findings = workbook["Findings"]
    assert findings["N2"].value.startswith("=HYPERLINK")
    assert findings["N2"].data_type == "s"
    assert findings["O2"].value == "=1+1"
    assert findings["O2"].data_type == "s"


def test_html_report_contents(qc_result: QCRunResult, tmp_path: Path) -> None:
    path = tmp_path / "report.html"
    write_html_report(qc_result, path)
    html = path.read_text(encoding="utf-8")

    assert html.startswith("<!DOCTYPE html>")
    assert "F0001" in html
    assert "cycle_comparison" in html
    assert "Check coverage" in html
    for severity in Severity:
        assert f'class="card {severity.value}"' in html
    # Self-contained: no external asset references.
    assert "http://" not in html and "https://" not in html
    assert "pattern review items" in html
    assert "atomic findings" in html
    # Every atomic finding remains available in safely escaped inline JSON.
    assert all(finding.finding_id in html for finding in qc_result.findings)
    assert "data-member-body" in html
    assert "const pageSize = 50" in html


def test_html_report_escapes_client_content() -> None:
    hostile = QCRunResult(profile_name="test")
    finding = Finding(
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        severity=Severity.CRITICAL,
        finding_id="F0001",
        sheet="Sheet1",
        location="A1",
        baseline_value="<script>alert('x')</script>",
        current_value="<img src=x onerror=alert(1)>",
        message="cell containing <b>markup</b> changed",
    )
    hostile.findings = [finding]
    html = render_html_report(hostile)
    assert "<script>alert" not in html
    assert "<img src=x" not in html
    assert r"\u003cscript\u003e" in html
    assert r"\u003cimg src=x onerror=alert(1)\u003e" in html
    assert "&lt;b&gt;markup&lt;/b&gt;" in html


def test_json_context_is_private_and_opt_in(qc_result: QCRunResult) -> None:
    private_payload = result_payload(qc_result)
    assert private_payload["schema_version"] == 1
    assert private_payload["context_included"] is False
    assert private_payload["mapping_suggestions"] == []
    assert all(
        "baseline_excerpt" not in finding and "current_excerpt" not in finding
        for finding in private_payload["findings"]
    )
    assert all(
        {"expected_reason", "temporal_context", "evidence_tags"} <= finding.keys()
        for finding in private_payload["findings"]
    )

    diagnostic_payload = result_payload(qc_result, include_context=True)
    assert diagnostic_payload["context_included"] is True
    assert any(
        finding.get("baseline_excerpt") or finding.get("current_excerpt")
        for finding in diagnostic_payload["findings"]
    )
