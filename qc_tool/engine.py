"""QC run orchestrator: load, align, diff, cross-check, triage.

One `run_qc` call is one QC run — a baseline/current workbook pair and/or
a baseline/current deck pair, an optional profile, and per-file passwords.
The result carries triaged findings plus capability disclosures, including
whether XLSB formula text was enriched or only formula presence was checked.
"""

import logging
from dataclasses import dataclass, field
from pathlib import Path

from qc_tool.availability import (
    availability_coverage,
    excel_availability_issues,
    ppt_availability_issues,
)
from qc_tool.config.profile import DeliverableProfile, default_profile
from qc_tool.coverage import CoverageItem, CoverageState, MappingCoverage, QCRunMode
from qc_tool.crosscheck.package import reconcile_package
from qc_tool.crosscheck.trace import (
    MappingSuggestion,
    annotate_ppt_chart_impacts,
    verify_mappings,
)
from qc_tool.excel.align import align_workbooks
from qc_tool.excel.charts import annotate_chart_impacts, chart_reference_coverage
from qc_tool.excel.context import attach_current_excerpts, attach_excerpts
from qc_tool.excel.controls import evaluate_controls
from qc_tool.excel.dependency import (
    DependencyGraph,
    annotate_impacts,
    build_dependency_graph,
    limit_impacts,
)
from qc_tool.excel.diff_structure import diff_workbook_structure
from qc_tool.excel.diff_values import diff_workbook_values
from qc_tool.excel.formulas import diff_workbook_formulas, formula_text_compatible
from qc_tool.excel.interaction import (
    conditional_style_coverage,
    interaction_rule_coverage,
)
from qc_tool.excel.preflight import preflight_workbook
from qc_tool.findings import Finding, FindingClass, Severity
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.io.model import WorkbookSnapshot
from qc_tool.ppt.diff import diff_decks
from qc_tool.ppt.extract import load_deck_snapshot
from qc_tool.ppt.match import match_slides
from qc_tool.ppt.preflight import preflight_deck
from qc_tool.triage.rules import triage

logger = logging.getLogger(__name__)


def _pair_coverage_state(
    baseline: CoverageState,
    current: CoverageState,
) -> CoverageState:
    if CoverageState.UNAVAILABLE in {baseline, current}:
        return CoverageState.UNAVAILABLE
    if CoverageState.DEGRADED in {baseline, current}:
        return CoverageState.DEGRADED
    return CoverageState.CHECKED


def _single_formula_disclosure(workbook: WorkbookSnapshot) -> str | None:
    if getattr(workbook, "formulas_available", False):
        return None
    if getattr(workbook, "formula_presence_available", False):
        return (
            "Formula QC is degraded: formula-record presence was checked, so missing "
            "formulas and saved error values were evaluated. Formula text was unavailable, "
            "so logic, consistency, range-extension semantics, and error references inside "
            "formula text could not be evaluated."
        )
    return (
        "Formula QC is degraded: neither reliable formula text nor formula-record "
        "presence was available; only saved error values were evaluated."
    )


def _pair_formula_disclosure(
    base_wb: WorkbookSnapshot, curr_wb: WorkbookSnapshot
) -> str | None:
    if formula_text_compatible(base_wb, curr_wb):
        return None
    if getattr(base_wb, "formula_presence_available", False) and getattr(
        curr_wb, "formula_presence_available", False
    ):
        return (
            "Formula QC is degraded for this pair: formula-record presence was checked, "
            "so hardcoding, removals, and missing fill formulas were evaluated. Compatible "
            "formula text was unavailable, so logic changes, consistency, range-extension "
            "semantics, and error references inside formula text could not be evaluated."
        )
    return (
        "Formula QC is degraded for this pair: reliable formula-record presence was "
        "unavailable for one or both workbooks; only saved error values were evaluated."
    )


@dataclass(slots=True)
class QCRunResult:
    profile_name: str
    mode: QCRunMode = QCRunMode.CYCLE_COMPARISON
    files: dict[str, str] = field(default_factory=dict)  # role -> file name
    findings: list[Finding] = field(default_factory=list)
    disclosures: list[str] = field(default_factory=list)
    coverage: list[CoverageItem] = field(default_factory=list)
    mapping_coverage: MappingCoverage | None = None
    mapping_suggestions: list[MappingSuggestion] = field(default_factory=list)
    verified_crosschecks: int = 0

    @property
    def counts(self) -> dict[Severity, int]:
        result = dict.fromkeys(Severity, 0)
        for finding in self.findings:
            if finding.severity is not None:
                result[finding.severity] += 1
        return result


def run_qc(
    *,
    baseline_excel: Path | None = None,
    current_excel: Path | None = None,
    baseline_ppt: Path | None = None,
    current_ppt: Path | None = None,
    profile: DeliverableProfile | None = None,
    passwords: dict[str, str] | None = None,
    mode: QCRunMode = QCRunMode.CYCLE_COMPARISON,
) -> QCRunResult:
    """Run a full QC comparison. ``passwords`` is keyed by file name."""
    mode = QCRunMode(mode)
    if mode is QCRunMode.CURRENT_FILE_PREFLIGHT:
        if baseline_excel is not None or baseline_ppt is not None:
            raise ValueError("current-file preflight does not accept baseline files")
        if current_excel is None and current_ppt is None:
            raise ValueError("current-file preflight needs a current Excel or PowerPoint file")
        profile = profile or default_profile()
        passwords = passwords or {}
        result = QCRunResult(profile_name=profile.name, mode=mode)
        findings: list[Finding] = []
        if current_excel is not None:
            workbook = load_workbook_snapshot(
                current_excel, password=passwords.get(current_excel.name)
            )
            result.files["current_excel"] = current_excel.name
            excel_preflight = preflight_workbook(workbook, profile)
            attach_current_excerpts(excel_preflight.findings, workbook)
            findings.extend(excel_preflight.findings)
            result.coverage.extend(excel_preflight.coverage)
            if disclosure := _single_formula_disclosure(workbook):
                result.disclosures.append(disclosure)
        else:
            result.coverage.append(
                CoverageItem(
                    check_id="excel-intrinsic",
                    label="Current Excel intrinsic checks",
                    artifact="excel",
                    state=CoverageState.UNAVAILABLE,
                    detail="No current workbook supplied",
                )
            )
        if current_ppt is not None:
            deck = load_deck_snapshot(current_ppt, password=passwords.get(current_ppt.name))
            result.files["current_ppt"] = current_ppt.name
            ppt_preflight = preflight_deck(deck, profile.ppt)
            findings.extend(ppt_preflight.findings)
            result.coverage.extend(ppt_preflight.coverage)
        else:
            result.coverage.append(
                CoverageItem(
                    check_id="ppt-intrinsic",
                    label="Current PowerPoint intrinsic checks",
                    artifact="ppt",
                    state=CoverageState.UNAVAILABLE,
                    detail="No current deck supplied",
                )
            )
        result.findings = triage(findings, profile)
        result.coverage.extend(
            [
            CoverageItem(
                check_id="excel-cycle-comparison",
                label="Historical Excel changes",
                artifact="excel",
                state=CoverageState.UNAVAILABLE,
                detail="No baseline workbook supplied",
            ),
            CoverageItem(
                check_id="ppt-cycle-comparison",
                label="Historical PowerPoint changes",
                artifact="ppt",
                state=CoverageState.UNAVAILABLE,
                detail="No baseline deck supplied",
            ),
            CoverageItem(
                check_id="excel-ppt-crosscheck",
                label="Excel to PowerPoint mappings",
                artifact="package",
                state=CoverageState.UNAVAILABLE,
                detail="Use final-package mode for current Excel-to-PowerPoint QC",
            ),
            ]
        )
        return result
    if mode is QCRunMode.FINAL_PACKAGE:
        if baseline_excel is not None or baseline_ppt is not None:
            raise ValueError("final-package QC does not accept baseline files")
        if current_excel is None or current_ppt is None:
            raise ValueError("final-package QC needs current Excel and PowerPoint files")
        profile = profile or default_profile()
        passwords = passwords or {}
        workbook = load_workbook_snapshot(
            current_excel, password=passwords.get(current_excel.name)
        )
        deck = load_deck_snapshot(current_ppt, password=passwords.get(current_ppt.name))
        result = QCRunResult(profile_name=profile.name, mode=mode)
        result.files = {
            "current_excel": current_excel.name,
            "current_ppt": current_ppt.name,
        }
        findings: list[Finding] = []
        excel_preflight = preflight_workbook(workbook, profile)
        ppt_preflight = preflight_deck(deck, profile.ppt)
        package = reconcile_package(workbook, deck, profile.crosscheck)
        attach_current_excerpts(excel_preflight.findings, workbook)
        findings.extend(excel_preflight.findings)
        findings.extend(ppt_preflight.findings)
        findings.extend(package.findings)
        if excel_preflight.dependency_graph is not None:
            annotate_ppt_chart_impacts(
                findings,
                deck,
                profile.crosscheck,
                excel_preflight.dependency_graph,
            )
            limit_impacts(findings)
        result.findings = triage(findings, profile)
        result.coverage = [
            *excel_preflight.coverage,
            *ppt_preflight.coverage,
            *package.coverage,
            CoverageItem(
                check_id="cycle-comparison",
                label="Historical cycle comparison",
                artifact="package",
                state=CoverageState.UNAVAILABLE,
                detail="No baseline files supplied",
            ),
        ]
        result.mapping_coverage = package.mapping_coverage
        result.mapping_suggestions = package.suggestions
        result.verified_crosschecks = package.mapping_coverage.verified
        if disclosure := _single_formula_disclosure(workbook):
            result.disclosures.append(disclosure)
        return result
    if mode is not QCRunMode.CYCLE_COMPARISON:
        raise NotImplementedError(f"{mode.value} is not enabled yet")
    if (baseline_excel is None) != (current_excel is None):
        raise ValueError("Excel comparison needs both a baseline and a current workbook")
    if (baseline_ppt is None) != (current_ppt is None):
        raise ValueError("PPT comparison needs both a baseline and a current deck")
    if baseline_excel is None and baseline_ppt is None:
        raise ValueError("nothing to compare: supply an Excel pair and/or a PPT pair")

    profile = profile or default_profile()
    passwords = passwords or {}
    result = QCRunResult(profile_name=profile.name, mode=mode)
    findings: list[Finding] = []
    current_workbook = None
    dependency_graph: DependencyGraph | None = None

    if baseline_excel is not None and current_excel is not None:
        base_wb = load_workbook_snapshot(
            baseline_excel, password=passwords.get(baseline_excel.name)
        )
        curr_wb = load_workbook_snapshot(
            current_excel, password=passwords.get(current_excel.name)
        )
        current_workbook = curr_wb
        result.files["baseline_excel"] = baseline_excel.name
        result.files["current_excel"] = current_excel.name
        if disclosure := _pair_formula_disclosure(base_wb, curr_wb):
            result.disclosures.append(disclosure)
        alignment = align_workbooks(base_wb, curr_wb, profile)
        value_start = len(findings)
        findings += diff_workbook_values(base_wb, curr_wb, alignment, profile)
        result.coverage.append(
            CoverageItem(
                check_id="excel-values",
                label="Excel values and presentation",
                artifact="excel",
                state=CoverageState.CHECKED,
                findings=len(findings) - value_start,
            )
        )
        structure_start = len(findings)
        structure_findings = diff_workbook_structure(
            base_wb,
            curr_wb,
            alignment,
            profile,
        )
        findings += structure_findings
        base_chart_state, base_chart_detail = chart_reference_coverage(base_wb)
        curr_chart_state, curr_chart_detail = chart_reference_coverage(curr_wb)
        structure_complete = (
            base_wb.tables_available
            and curr_wb.tables_available
            and base_chart_state is CoverageState.CHECKED
            and curr_chart_state is CoverageState.CHECKED
        )
        result.coverage.append(
            CoverageItem(
                check_id="excel-structure",
                label="Excel workbook structure",
                artifact="excel",
                state=(
                    CoverageState.CHECKED
                    if structure_complete
                    else CoverageState.DEGRADED
                ),
                findings=len(findings) - structure_start,
                detail=(
                    ""
                    if structure_complete
                    else "; ".join(
                        detail
                        for detail in (
                            (
                                "Excel table metadata is unavailable for one or both files"
                                if not base_wb.tables_available
                                or not curr_wb.tables_available
                                else ""
                            ),
                            (
                                f"baseline charts: {base_chart_detail}"
                                if base_chart_state is not CoverageState.CHECKED
                                else ""
                            ),
                            (
                                f"current charts: {curr_chart_detail}"
                                if curr_chart_state is not CoverageState.CHECKED
                                else ""
                            ),
                        )
                        if detail
                    )
                ),
            )
        )
        interaction_classes = {
            FindingClass.DATA_VALIDATION_CHANGED,
            FindingClass.CONDITIONAL_FORMAT_CHANGED,
        }
        interaction_findings = sum(
            finding.finding_class in interaction_classes
            for finding in structure_findings
        )
        baseline_rule_state, baseline_rule_detail = interaction_rule_coverage(base_wb)
        current_rule_state, current_rule_detail = interaction_rule_coverage(curr_wb)
        rule_state = _pair_coverage_state(
            baseline_rule_state,
            current_rule_state,
        )
        result.coverage.append(
            CoverageItem(
                check_id="excel-interaction-rules",
                label="Data validation and conditional-format rules",
                artifact="excel",
                state=rule_state,
                findings=interaction_findings,
                detail=(
                    "Complete interaction-rule comparison"
                    if rule_state is CoverageState.CHECKED
                    else (
                        f"baseline: {baseline_rule_detail}; "
                        f"current: {current_rule_detail}"
                    )
                ),
            )
        )
        baseline_style_state, baseline_style_detail = conditional_style_coverage(
            base_wb
        )
        current_style_state, current_style_detail = conditional_style_coverage(
            curr_wb
        )
        style_state = _pair_coverage_state(
            baseline_style_state,
            current_style_state,
        )
        result.coverage.append(
            CoverageItem(
                check_id="excel-conditional-format-styles",
                label="Conditional-format differential styles",
                artifact="excel",
                state=style_state,
                detail=(
                    "Stable differential styles compared"
                    if style_state is CoverageState.CHECKED
                    else (
                        f"baseline: {baseline_style_detail}; "
                        f"current: {current_style_detail}"
                    )
                ),
            )
        )
        result.coverage.append(
            availability_coverage(
                artifact="excel",
                rule_count=sum(
                    len(sheet.availability_rules)
                    for sheet in profile.excel.sheets.values()
                ),
                issues=excel_availability_issues(curr_wb, profile),
            )
        )
        formula_start = len(findings)
        findings += diff_workbook_formulas(
            base_wb,
            curr_wb,
            alignment,
            profile,
        )
        formula_text_checked = formula_text_compatible(base_wb, curr_wb)
        result.coverage.append(
            CoverageItem(
                check_id="excel-formulas",
                label="Excel formulas",
                artifact="excel",
                state=(
                    CoverageState.CHECKED
                    if formula_text_checked
                    else CoverageState.DEGRADED
                ),
                findings=len(findings) - formula_start,
                detail=(
                    f"Compatible formula text source: {base_wb.formula_source}"
                    if formula_text_checked
                    else (
                        "Formula presence checks completed; semantic formula text "
                        "checks unavailable"
                    )
                ),
            )
        )
        if curr_wb.formulas_available:
            dependency_graph = build_dependency_graph(curr_wb)
            annotate_impacts(findings, dependency_graph)
            dependency_state = dependency_graph.coverage_state
            dependency_detail = dependency_graph.coverage_detail
        else:
            dependency_state = CoverageState.UNAVAILABLE
            dependency_detail = "Formula text is unavailable for dependency extraction"
        result.coverage.append(
            CoverageItem(
                check_id="excel-dependencies",
                label="Formula dependency impact tracing",
                artifact="excel",
                state=dependency_state,
                detail=dependency_detail,
            )
        )
        controls = evaluate_controls(curr_wb, profile.excel.controls)
        findings += controls.findings
        if controls.coverage is not None:
            result.coverage.append(controls.coverage)
        annotate_chart_impacts(findings, curr_wb, dependency_graph)
        limit_impacts(findings)
        attach_excerpts(findings, base_wb, curr_wb)

    current_deck = None
    if baseline_ppt is not None and current_ppt is not None:
        base_deck = load_deck_snapshot(
            baseline_ppt, password=passwords.get(baseline_ppt.name)
        )
        current_deck = load_deck_snapshot(
            current_ppt, password=passwords.get(current_ppt.name)
        )
        result.files["baseline_ppt"] = baseline_ppt.name
        result.files["current_ppt"] = current_ppt.name
        matching = match_slides(base_deck, current_deck, profile.ppt)
        ppt_start = len(findings)
        findings += diff_decks(matching, profile.ppt)
        result.coverage.append(
            CoverageItem(
                check_id="ppt-comparison",
                label="PowerPoint content and charts",
                artifact="ppt",
                state=(
                    CoverageState.CHECKED
                    if base_deck.charts_available and current_deck.charts_available
                    else CoverageState.DEGRADED
                ),
                findings=len(findings) - ppt_start,
                detail=(
                    "Complete native chart extraction unavailable for one or both decks"
                    if not base_deck.charts_available
                    or not current_deck.charts_available
                    else "Complete semantic slide-element comparison"
                ),
            )
        )
        result.coverage.append(
            availability_coverage(
                artifact="ppt",
                rule_count=len(profile.ppt.availability_rules),
                issues=ppt_availability_issues(current_deck, profile.ppt),
            )
        )

    if (
        current_deck is not None
        and dependency_graph is not None
        and profile.crosscheck.mappings
    ):
        annotate_ppt_chart_impacts(
            findings,
            current_deck,
            profile.crosscheck,
            dependency_graph,
        )
        limit_impacts(findings)

    if baseline_excel is None:
        result.coverage.extend(
            [
                CoverageItem(
                    check_id="excel-values",
                    label="Excel values and presentation",
                    artifact="excel",
                    state=CoverageState.UNAVAILABLE,
                    detail="Excel pair not supplied",
                ),
                CoverageItem(
                    check_id="excel-structure",
                    label="Excel workbook structure",
                    artifact="excel",
                    state=CoverageState.UNAVAILABLE,
                    detail="Excel pair not supplied",
                ),
                CoverageItem(
                    check_id="excel-formulas",
                    label="Excel formulas",
                    artifact="excel",
                    state=CoverageState.UNAVAILABLE,
                    detail="Excel pair not supplied",
                ),
                CoverageItem(
                    check_id="excel-dependencies",
                    label="Formula dependency impact tracing",
                    artifact="excel",
                    state=CoverageState.UNAVAILABLE,
                    detail="Excel pair not supplied",
                ),
                CoverageItem(
                    check_id="excel-interaction-rules",
                    label="Data validation and conditional-format rules",
                    artifact="excel",
                    state=CoverageState.UNAVAILABLE,
                    detail="Excel pair not supplied",
                ),
                CoverageItem(
                    check_id="excel-conditional-format-styles",
                    label="Conditional-format differential styles",
                    artifact="excel",
                    state=CoverageState.UNAVAILABLE,
                    detail="Excel pair not supplied",
                ),
                CoverageItem(
                    check_id="excel-availability",
                    label="Availability boundaries",
                    artifact="excel",
                    state=CoverageState.UNAVAILABLE,
                    detail="Excel pair not supplied",
                ),
            ]
        )
    if baseline_ppt is None:
        result.coverage.extend(
            [
                CoverageItem(
                    check_id="ppt-comparison",
                    label="PowerPoint content and charts",
                    artifact="ppt",
                    state=CoverageState.UNAVAILABLE,
                    detail="PowerPoint pair not supplied",
                ),
                CoverageItem(
                    check_id="ppt-availability",
                    label="Availability boundaries",
                    artifact="ppt",
                    state=CoverageState.UNAVAILABLE,
                    detail="PowerPoint pair not supplied",
                ),
            ]
        )

    if current_deck is not None and current_workbook is not None and profile.crosscheck.mappings:
        crosscheck_start = len(findings)
        crosscheck = verify_mappings(current_deck, current_workbook, profile.crosscheck)
        findings += crosscheck.findings
        result.verified_crosschecks = len(crosscheck.verified)
        result.coverage.append(
            CoverageItem(
                check_id="excel-ppt-crosscheck",
                label="Excel to PowerPoint mappings",
                artifact="package",
                state=(
                    CoverageState.CHECKED
                    if current_deck.charts_available
                    else CoverageState.DEGRADED
                ),
                findings=len(findings) - crosscheck_start,
                detail=(
                    f"{len(crosscheck.verified)} mappings verified"
                    + (
                        ""
                        if current_deck.charts_available
                        else "; visible native chart labels unavailable"
                    )
                ),
            )
        )
    else:
        reason = (
            "Both current Excel and PowerPoint are required"
            if current_deck is None or current_workbook is None
            else "No confirmed mappings in the selected profile"
        )
        result.coverage.append(
            CoverageItem(
                check_id="excel-ppt-crosscheck",
                label="Excel to PowerPoint mappings",
                artifact="package",
                state=CoverageState.UNAVAILABLE,
                detail=reason,
            )
        )

    result.findings = triage(findings, profile)
    logger.info(
        "QC run complete: %s findings (%s)",
        len(result.findings),
        ", ".join(f"{sev.value}={count}" for sev, count in result.counts.items()),
    )
    return result


@dataclass(slots=True)
class FindingsDelta:
    """How a re-QC run compares to its predecessor (non-expected findings)."""

    resolved: int
    new: int
    persisting: int


def _identity_key(finding: Finding) -> tuple[str, ...]:
    return (
        finding.artifact,
        finding.finding_class.value,
        finding.sheet or "",
        finding.slide or "",
        finding.location or finding.baseline_location or "",
        finding.element or "",
    )


def compare_findings(previous: list[Finding], current: list[Finding]) -> FindingsDelta:
    """Match findings by identity (not by id) to compute a fix-progress delta."""
    previous_keys = {
        _identity_key(f) for f in previous if f.severity is not Severity.EXPECTED
    }
    current_keys = {
        _identity_key(f) for f in current if f.severity is not Severity.EXPECTED
    }
    return FindingsDelta(
        resolved=len(previous_keys - current_keys),
        new=len(current_keys - previous_keys),
        persisting=len(previous_keys & current_keys),
    )
