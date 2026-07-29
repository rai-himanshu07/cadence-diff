"""Declarative workbook controls, root-cause grouping, and expiring waivers."""

from qc_tool.config.profile import DeliverableProfile
from qc_tool.excel.preflight import preflight_workbook
from qc_tool.findings import Finding, FindingClass, Severity
from qc_tool.io.model import CellRecord, SheetSnapshot, WorkbookSnapshot
from qc_tool.triage.rules import triage


def test_profile_controls_cover_required_unique_bounds_and_tie_out() -> None:
    sheet = SheetSnapshot(
        name="Data",
        visibility="visible",
        max_row=4,
        max_column=3,
        cells={
            (1, 1): CellRecord(1, 1, "ID"),
            (1, 2): CellRecord(1, 2, "Value"),
            (1, 3): CellRecord(1, 3, "Total"),
            (2, 1): CellRecord(2, 1, "A"),
            (2, 2): CellRecord(2, 2, 60),
            (2, 3): CellRecord(2, 3, 200),
            (3, 1): CellRecord(3, 1, "A"),
            (3, 2): CellRecord(3, 2, 150),
            (4, 1): CellRecord(4, 1, "B"),
        },
    )
    workbook = WorkbookSnapshot(
        source_name="current.xlsx",
        file_format="xlsx",
        formulas_available=True,
        styles_available=True,
        sheets=[sheet],
    )
    profile = DeliverableProfile.model_validate(
        {
            "name": "controlled",
            "excel": {
                "controls": {
                    "required_ranges": [
                        {"name": "Values required", "sheet": "Data", "range": "B2:B4"}
                    ],
                    "unique_ranges": [
                        {"name": "Unique IDs", "sheet": "Data", "range": "A1:A4"}
                    ],
                    "numeric_bounds": [
                        {
                            "name": "Plausible values",
                            "sheet": "Data",
                            "range": "B2:B3",
                            "minimum": 0,
                            "maximum": 100,
                        }
                    ],
                    "tie_outs": [
                        {
                            "name": "Reported total",
                            "target": "Data!C2",
                            "components": ["Data!B2:B3"],
                        }
                    ],
                }
            },
        }
    )

    result = preflight_workbook(workbook, profile)
    classes = {finding.finding_class for finding in result.findings}

    assert {
        FindingClass.REQUIRED_VALUE_MISSING,
        FindingClass.DUPLICATE_KEY,
        FindingClass.NUMERIC_BOUND_VIOLATION,
        FindingClass.TIE_OUT_MISMATCH,
    } <= classes
    coverage = next(
        item for item in result.coverage if item.check_id == "excel-profile-controls"
    )
    assert coverage.detail == "4 controls configured"
    assert coverage.findings == 4


def test_active_and_expired_waivers_preserve_evidence_and_group_roots() -> None:
    findings = [
        Finding(
            artifact="excel",
            finding_class=FindingClass.FORMULA_INCONSISTENT,
            sheet="Data",
            location="B2",
            message="inconsistent",
        ),
        Finding(
            artifact="excel",
            finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
            sheet="Data",
            location="B2",
            message="logic changed",
        ),
    ]
    profile = DeliverableProfile.model_validate(
        {
            "name": "waived",
            "waivers": [
                {
                    "finding_class": "formula_inconsistent",
                    "sheet": "Data",
                    "location": "B2",
                    "reason": "approved model exception",
                    "expires": "2999-12-31",
                },
                {
                    "finding_class": "style_changed",
                    "reason": "old branding exception",
                    "expires": "2000-01-01",
                },
            ],
        }
    )

    results = triage(findings, profile)
    waived = next(
        finding
        for finding in results
        if finding.finding_class is FindingClass.FORMULA_INCONSISTENT
    )
    logic = next(
        finding
        for finding in results
        if finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
    )
    expired = next(
        finding
        for finding in results
        if finding.finding_class is FindingClass.WAIVER_EXPIRED
    )

    assert waived.severity is Severity.EXPECTED
    assert waived.waiver_reason == "approved model exception"
    assert waived.waiver_expires == "2999-12-31"
    assert logic.severity is Severity.WARNING
    assert waived.root_cause_key == logic.root_cause_key == "excel:Data:B2"
    assert expired.severity is Severity.WARNING
