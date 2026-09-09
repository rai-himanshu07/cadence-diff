"""Privacy-safe findings for intrinsic workbook package risks."""

from qc_tool.coverage import CoverageItem, CoverageState
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

#: Risk kinds sourced only from passive external-workbook-link metadata
#: (Step 4a classification). Eligible for suppression once Step 4b's
#: reachability pass proves, with complete evidence, that no current formula
#: depends on the link -- never for data connections or query tables, which
#: stay hard blockers regardless of reachability.
_PASSIVE_LINK_KINDS = frozenset(
    {WorkbookRiskKind.EXTERNAL_WORKBOOK_LINK, WorkbookRiskKind.EXTERNAL_RELATIONSHIP}
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
    WorkbookRiskKind.UNREADABLE_RELATIONSHIP_METADATA: ("unreadable relationship metadata"),
}


def workbook_risk_findings(
    current: WorkbookSnapshot,
    baseline: WorkbookSnapshot | None = None,
) -> list[Finding]:
    """Emit one bounded finding per risk kind; package targets never enter it.

    A passive-link risk kind (`_PASSIVE_LINK_KINDS`) is suppressed only when
    `current.external_link_reachability` is both proven and not live --
    complete trusted formula text and defined-name reachability confirmed no
    current formula depends on it. Incomplete proof, or a live dependency,
    always keeps the finding; `external_link_reachability_coverage` discloses
    the verdict either way.
    """
    baseline_kinds = (
        {risk.kind for risk in baseline.intrinsic_risks} if baseline is not None else set()
    )
    reachability = current.external_link_reachability
    suppress_passive = bool(
        reachability is not None and reachability.proven and not reachability.live
    )
    findings: list[Finding] = []
    for risk in sorted(current.intrinsic_risks, key=lambda item: item.kind.value):
        if suppress_passive and risk.kind in _PASSIVE_LINK_KINDS:
            continue
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


def external_link_reachability_coverage(*workbooks: WorkbookSnapshot) -> CoverageItem:
    """Disclose passive-link reachability across every supplied workbook.

    Returns a `CoverageState.CHECKED` item (empty detail) when no supplied
    workbook carries a passive-link risk kind. When at least one does, state
    is `CHECKED` only if every such workbook's reachability is proven
    (whether or not it turned out live); any unproven workbook keeps the
    item `DEGRADED`, since usage cannot be confirmed either way.
    """
    relevant = [
        book
        for book in workbooks
        if any(risk.kind in _PASSIVE_LINK_KINDS for risk in book.intrinsic_risks)
    ]
    if not relevant:
        return CoverageItem(
            check_id="xlsb-external-links",
            label="External workbook links",
            artifact="excel",
            state=CoverageState.CHECKED,
            detail="",
        )
    details: list[str] = []
    degraded = False
    for book in relevant:
        reachability = book.external_link_reachability
        if reachability is not None and reachability.proven:
            verdict = "in use" if reachability.live else "not used"
            details.append(
                f"{book.source_name}: external link metadata present; formula and "
                f"defined-name reachability confirms it is {verdict} by a current formula"
            )
        else:
            degraded = True
            details.append(
                f"{book.source_name}: external link metadata present; formula and "
                "defined-name reachability is incomplete, so usage cannot be proven "
                "either way"
            )
    return CoverageItem(
        check_id="xlsb-external-links",
        label="External workbook links",
        artifact="excel",
        state=CoverageState.DEGRADED if degraded else CoverageState.CHECKED,
        detail="; ".join(details),
    )
