"""Privacy-safe findings for intrinsic workbook package risks."""

from qc_tool.findings import Finding, FindingClass, FindingProvenance
from qc_tool.io.model import WorkbookRiskKind, WorkbookSnapshot

_EXTERNAL_KINDS = frozenset(
    {
        WorkbookRiskKind.EXTERNAL_WORKBOOK_LINK,
        WorkbookRiskKind.EXTERNAL_RELATIONSHIP,
        WorkbookRiskKind.EXTERNAL_DATA_CONNECTION,
        WorkbookRiskKind.QUERY_TABLE,
    }
)

_RISK_LABELS = {
    WorkbookRiskKind.EXTERNAL_WORKBOOK_LINK: "external workbook link",
    WorkbookRiskKind.EXTERNAL_RELATIONSHIP: "external package relationship",
    WorkbookRiskKind.EXTERNAL_DATA_CONNECTION: "external data connection",
    WorkbookRiskKind.QUERY_TABLE: "external query table",
    WorkbookRiskKind.VBA_PROJECT: "VBA project",
    WorkbookRiskKind.EXCEL4_MACRO_SHEET: "Excel 4.0 macro sheet",
    WorkbookRiskKind.ACTIVEX_CONTROL: "ActiveX control",
    WorkbookRiskKind.EMBEDDED_OLE: "embedded OLE content",
    WorkbookRiskKind.CONTROL_CONTENT: "control content",
    WorkbookRiskKind.DIALOG_SHEET: "dialog sheet",
    WorkbookRiskKind.CUSTOM_OFFICE_UI: "custom Office UI",
    WorkbookRiskKind.UNREADABLE_RELATIONSHIP_METADATA: (
        "unreadable relationship metadata"
    ),
}


def workbook_risk_findings(
    current: WorkbookSnapshot,
    baseline: WorkbookSnapshot | None = None,
) -> list[Finding]:
    """Emit one bounded finding per risk kind; package targets never enter it."""
    baseline_kinds = (
        {risk.kind for risk in baseline.intrinsic_risks}
        if baseline is not None
        else set()
    )
    findings: list[Finding] = []
    for risk in sorted(current.intrinsic_risks, key=lambda item: item.kind.value):
        provenance = None
        if baseline is not None:
            provenance = (
                FindingProvenance.INHERITED
                if risk.kind in baseline_kinds
                else FindingProvenance.NEW
            )
        label = _RISK_LABELS[risk.kind]
        findings.append(
            Finding(
                artifact="excel",
                finding_class=(
                    FindingClass.EXTERNAL_LINK
                    if risk.kind in _EXTERNAL_KINDS
                    else FindingClass.ACTIVE_CONTENT
                ),
                provenance=provenance,
                event_key=f"workbook-risk:{risk.kind.value}",
                element=risk.kind.value,
                current_value=str(risk.count),
                message=(
                    f"structural package scan detected {risk.count} {label} "
                    f"instance{'s' if risk.count != 1 else ''}"
                ),
            )
        )
    return findings
