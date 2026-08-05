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
from qc_tool.config.profile import (
    CrosscheckProfile,
    DeliverableProfile,
    NumericTolerance,
    default_profile,
)
from qc_tool.coverage import CoverageItem, CoverageState, MappingCoverage, QCRunMode
from qc_tool.crosscheck.package import reconcile_package
from qc_tool.crosscheck.trace import (
    MappingSuggestion,
    annotate_ppt_chart_impacts,
    verify_mappings,
)
from qc_tool.excel.align import align_workbooks
from qc_tool.excel.charts import annotate_chart_impacts, chart_reference_coverage
from qc_tool.excel.complexity import WorkbookComplexity, assess_workbook_complexity
from qc_tool.excel.context import attach_current_excerpts, attach_excerpts
from qc_tool.excel.controls import evaluate_controls
from qc_tool.excel.dependency import (
    DependencyGraph,
    annotate_impacts,
    build_dependency_graph,
    limit_impacts,
)
from qc_tool.excel.diff_metadata import (
    comment_coverage,
    connection_coverage,
    diff_workbook_metadata,
    external_connection_findings,
    power_query_coverage,
)
from qc_tool.excel.diff_structure import diff_workbook_structure
from qc_tool.excel.diff_values import diff_workbook_values
from qc_tool.excel.diff_vba import diff_workbook_vba, vba_coverage
from qc_tool.excel.formulas import diff_workbook_formulas, formula_text_compatible
from qc_tool.excel.interaction import (
    conditional_style_coverage,
    interaction_rule_coverage,
)
from qc_tool.excel.preflight import defined_name_scope_coverage, preflight_workbook
from qc_tool.excel.workbook_risks import workbook_risk_findings
from qc_tool.findings import Finding, FindingClass, Severity, limit_findings
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.io.model import WorkbookSnapshot
from qc_tool.ppt.diff import diff_decks
from qc_tool.ppt.extract import load_deck_snapshot
from qc_tool.ppt.match import match_slides
from qc_tool.ppt.model import DeckSnapshot
from qc_tool.ppt.preflight import preflight_deck
from qc_tool.progress import (
    CancellationToken,
    ProgressCallback,
    RunPhase,
    check_cancelled,
    report_progress,
)
from qc_tool.scope import ComparisonScope
from qc_tool.story import annotate_story_evidence
from qc_tool.triage.rules import triage

logger = logging.getLogger(__name__)


def _load_excel_file(
    path: Path,
    *,
    password: str | None,
    phase: RunPhase,
    allow_large_workbooks: bool,
    cancellation_token: CancellationToken | None,
    on_progress: ProgressCallback | None,
) -> WorkbookSnapshot:
    check_cancelled(cancellation_token)
    report_progress(on_progress, phase, total=1, detail=path.name)
    workbook = load_workbook_snapshot(
        path,
        password=password,
        allow_large_workbook=allow_large_workbooks,
        cancellation_token=cancellation_token,
    )
    check_cancelled(cancellation_token)
    report_progress(on_progress, phase, processed=1, total=1, detail=path.name)
    return workbook


def _load_powerpoint_file(
    path: Path,
    *,
    password: str | None,
    phase: RunPhase,
    cancellation_token: CancellationToken | None,
    on_progress: ProgressCallback | None,
):
    check_cancelled(cancellation_token)
    report_progress(on_progress, phase, total=1, detail=path.name)
    deck = load_deck_snapshot(
        path,
        password=password,
        cancellation_token=cancellation_token,
    )
    check_cancelled(cancellation_token)
    report_progress(on_progress, phase, processed=1, total=1, detail=path.name)
    return deck


def _workload_coverage(*workbooks: WorkbookSnapshot) -> CoverageItem:
    ooxml = [
        workbook
        for workbook in workbooks
        if workbook.file_format in {"xlsx", "xlsm"}
    ]
    if not ooxml:
        return CoverageItem(
            check_id="excel-workload",
            label="Excel workload safeguards",
            artifact="excel",
            state=CoverageState.UNAVAILABLE,
            detail="OOXML workload metrics are unavailable for XLSB",
        )
    incomplete = len(ooxml) != len(workbooks)
    degraded = incomplete or any(workbook.workload.degraded for workbook in ooxml)
    details = [
        f"{workbook.source_name}: {workbook.workload.detail}"
        for workbook in ooxml
    ]
    if incomplete:
        details.append("XLSB workload metrics unavailable for one workbook")
    return CoverageItem(
        check_id="excel-workload",
        label="Excel workload safeguards",
        artifact="excel",
        state=CoverageState.DEGRADED if degraded else CoverageState.CHECKED,
        detail="; ".join(details),
    )


_CAPABILITY_ONLY_COVERAGE = frozenset({"excel-dependencies"})


def _complexity_coverage(complexity: WorkbookComplexity) -> CoverageItem:
    return CoverageItem(
        check_id="excel-dependency-workload",
        label="Excel dependency workload safeguards",
        artifact="excel",
        state=(
            CoverageState.DEGRADED if complexity.degraded else CoverageState.CHECKED
        ),
        detail="; ".join(
            complexity.warning_reasons
            or (
                f"{complexity.formula_count:,} formulas, "
                f"{complexity.reference_operands:,} reference operands, "
                f"{complexity.projected_concrete_edges:,} projected cell "
                f"dependencies, {complexity.interaction_rule_count:,} interaction "
                "rules within local limits",
            )
        ),
    )


def _apply_findings_budget(
    findings: list[Finding],
    coverage: list[CoverageItem],
) -> list[Finding]:
    budget = limit_findings(findings)
    if not budget.omitted_by_artifact:
        return budget.findings
    affected_artifacts = set(budget.omitted_by_artifact)
    if "crosscheck" in affected_artifacts:
        affected_artifacts.add("package")
    if budget.global_omitted:
        affected_artifacts.update(item.artifact for item in coverage)
    omitted_total = sum(budget.omitted_by_artifact.values())
    detail = (
        f"Output budget omitted {omitted_total} findings; summary finding(s) "
        "identify affected classes and scopes"
    )
    capability_detail = (
        "Dependency extraction and indexing were complete; downstream impacts and "
        f"cell context were not computed for {omitted_total} omitted findings"
    )
    for item in coverage:
        if item.artifact not in affected_artifacts:
            continue
        if item.check_id in _CAPABILITY_ONLY_COVERAGE:
            item.detail = (
                f"{item.detail}; {capability_detail}"
                if item.detail
                else capability_detail
            )
            continue
        if item.state is CoverageState.CHECKED:
            item.state = CoverageState.DEGRADED
        item.detail = f"{item.detail}; {detail}" if item.detail else detail
    return budget.findings


def _enrich_retained_findings(
    findings: list[Finding],
    *,
    baseline_workbook: WorkbookSnapshot | None = None,
    current_workbook: WorkbookSnapshot | None = None,
    current_deck: DeckSnapshot | None = None,
    dependency_graph: DependencyGraph | None = None,
    crosscheck: CrosscheckProfile | None = None,
) -> None:
    """Compute impacts and cell context for retained atomics only.

    Runs after the findings budget so omitted atomics never trigger a transitive
    closure query or a grid excerpt.
    """
    if dependency_graph is not None:
        annotate_impacts(findings, dependency_graph)
    if current_workbook is not None:
        annotate_chart_impacts(findings, current_workbook, dependency_graph)
    if (
        current_deck is not None
        and dependency_graph is not None
        and crosscheck is not None
        and crosscheck.mappings
    ):
        annotate_ppt_chart_impacts(findings, current_deck, crosscheck, dependency_graph)
    limit_impacts(findings)
    if baseline_workbook is not None and current_workbook is not None:
        attach_excerpts(findings, baseline_workbook, current_workbook)
    elif current_workbook is not None:
        attach_current_excerpts(findings, current_workbook)


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
    comparison_scope: ComparisonScope = field(default_factory=ComparisonScope)

    @property
    def counts(self) -> dict[Severity, int]:
        result = dict.fromkeys(Severity, 0)
        for finding in self.findings:
            if finding.severity is not None:
                result[finding.severity] += 1
        return result


def _apply_comparison_scope(
    result: QCRunResult,
    findings: list[Finding],
    requested: ComparisonScope,
    *,
    workbooks: tuple[WorkbookSnapshot, ...] = (),
    decks: tuple[DeckSnapshot, ...] = (),
) -> list[Finding]:
    validated = requested.validate_loaded(workbooks=workbooks, decks=decks)
    result.comparison_scope = validated
    result.coverage.append(validated.coverage_item(workbooks=workbooks, decks=decks))
    if disclosure := validated.disclosure():
        result.disclosures.append(disclosure)
    return validated.filter_findings(findings)


def run_qc(
    *,
    baseline_excel: Path | None = None,
    current_excel: Path | None = None,
    baseline_ppt: Path | None = None,
    current_ppt: Path | None = None,
    profile: DeliverableProfile | None = None,
    passwords: dict[str, str] | None = None,
    mode: QCRunMode = QCRunMode.CYCLE_COMPARISON,
    allow_large_workbooks: bool = False,
    run_acceptance: NumericTolerance | None = None,
    compare_sheets: list[str] | None = None,
    compare_slides: list[int] | None = None,
    cancellation_token: CancellationToken | None = None,
    on_progress: ProgressCallback | None = None,
) -> QCRunResult:
    """Run a full QC comparison. ``passwords`` is keyed by file name.

    ``run_acceptance`` is the analyst's run-level threshold: numeric value
    differences within either bound stay visible as within-tolerance Info
    findings (never suppressed). ``compare_sheets`` / ``compare_slides``
    narrow which sheets (by name) and slides (1-based index) may produce
    findings; files still load fully so cross-references keep resolving.
    """
    check_cancelled(cancellation_token)
    mode = QCRunMode(mode)
    requested_scope = ComparisonScope(
        excel_sheets=(tuple(compare_sheets) if compare_sheets is not None else None),
        ppt_slide_indices=(
            tuple(compare_slides) if compare_slides is not None else None
        ),
    )
    acceptance_requested = bool(
        run_acceptance is not None
        and (run_acceptance.absolute > 0 or run_acceptance.relative > 0)
    )
    acceptance_not_applied = (
        "Analyst acceptance threshold was not applied because this mode has no "
        "baseline/current numeric value comparison"
    )
    report_progress(
        on_progress,
        RunPhase.PREPARING,
        processed=1,
        total=1,
        detail=mode.value,
    )
    if mode is QCRunMode.CURRENT_FILE_PREFLIGHT:
        if baseline_excel is not None or baseline_ppt is not None:
            raise ValueError("current-file preflight does not accept baseline files")
        if current_excel is None and current_ppt is None:
            raise ValueError("current-file preflight needs a current Excel or PowerPoint file")
        profile = profile or default_profile()
        passwords = passwords or {}
        result = QCRunResult(profile_name=profile.name, mode=mode)
        if acceptance_requested:
            logger.warning("acceptance threshold ignored in %s", mode.value)
            result.disclosures.append(acceptance_not_applied)
        findings: list[Finding] = []
        preflight_workbook_snapshot: WorkbookSnapshot | None = None
        preflight_dependency_graph: DependencyGraph | None = None
        preflight_deck_snapshot: DeckSnapshot | None = None
        if current_excel is not None:
            workbook = _load_excel_file(
                current_excel,
                password=passwords.get(current_excel.name),
                phase=RunPhase.LOADING_CURRENT_EXCEL,
                allow_large_workbooks=allow_large_workbooks,
                cancellation_token=cancellation_token,
                on_progress=on_progress,
            )
            result.files["current_excel"] = current_excel.name
            report_progress(on_progress, RunPhase.ANALYZING_EXCEL, total=1)
            excel_preflight = preflight_workbook(
                workbook,
                profile,
                defer_impacts=True,
                cancellation_token=cancellation_token,
            )
            check_cancelled(cancellation_token)
            report_progress(
                on_progress,
                RunPhase.ANALYZING_EXCEL,
                processed=1,
                total=1,
            )
            findings.extend(excel_preflight.findings)
            preflight_workbook_snapshot = workbook
            preflight_dependency_graph = excel_preflight.dependency_graph
            result.coverage.extend(excel_preflight.coverage)
            result.coverage.append(_workload_coverage(workbook))
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
            deck = _load_powerpoint_file(
                current_ppt,
                password=passwords.get(current_ppt.name),
                phase=RunPhase.LOADING_CURRENT_POWERPOINT,
                cancellation_token=cancellation_token,
                on_progress=on_progress,
            )
            result.files["current_ppt"] = current_ppt.name
            report_progress(on_progress, RunPhase.ANALYZING_POWERPOINT, total=1)
            ppt_preflight = preflight_deck(deck, profile.ppt)
            preflight_deck_snapshot = deck
            check_cancelled(cancellation_token)
            report_progress(
                on_progress,
                RunPhase.ANALYZING_POWERPOINT,
                processed=1,
                total=1,
            )
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
        findings = _apply_comparison_scope(
            result,
            findings,
            requested_scope,
            workbooks=(
                (preflight_workbook_snapshot,)
                if preflight_workbook_snapshot is not None
                else ()
            ),
            decks=(
                (preflight_deck_snapshot,)
                if preflight_deck_snapshot is not None
                else ()
            ),
        )
        findings = _apply_findings_budget(findings, result.coverage)
        _enrich_retained_findings(
            findings,
            current_workbook=preflight_workbook_snapshot,
            dependency_graph=preflight_dependency_graph,
        )
        result.findings = triage(findings, profile)
        annotate_story_evidence(result.findings)
        return result
    if mode is QCRunMode.FINAL_PACKAGE:
        if baseline_excel is not None or baseline_ppt is not None:
            raise ValueError("final-package QC does not accept baseline files")
        if current_excel is None or current_ppt is None:
            raise ValueError("final-package QC needs current Excel and PowerPoint files")
        profile = profile or default_profile()
        passwords = passwords or {}
        workbook = _load_excel_file(
            current_excel,
            password=passwords.get(current_excel.name),
            phase=RunPhase.LOADING_CURRENT_EXCEL,
            allow_large_workbooks=allow_large_workbooks,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        deck = _load_powerpoint_file(
            current_ppt,
            password=passwords.get(current_ppt.name),
            phase=RunPhase.LOADING_CURRENT_POWERPOINT,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        result = QCRunResult(profile_name=profile.name, mode=mode)
        if acceptance_requested:
            logger.warning("acceptance threshold ignored in %s", mode.value)
            result.disclosures.append(acceptance_not_applied)
        result.files = {
            "current_excel": current_excel.name,
            "current_ppt": current_ppt.name,
        }
        findings: list[Finding] = []
        report_progress(on_progress, RunPhase.ANALYZING_EXCEL, total=1)
        excel_preflight = preflight_workbook(
            workbook,
            profile,
            defer_impacts=True,
            cancellation_token=cancellation_token,
        )
        check_cancelled(cancellation_token)
        report_progress(
            on_progress,
            RunPhase.ANALYZING_EXCEL,
            processed=1,
            total=1,
        )
        report_progress(on_progress, RunPhase.ANALYZING_POWERPOINT, total=1)
        ppt_preflight = preflight_deck(deck, profile.ppt)
        check_cancelled(cancellation_token)
        report_progress(
            on_progress,
            RunPhase.ANALYZING_POWERPOINT,
            processed=1,
            total=1,
        )
        report_progress(on_progress, RunPhase.CROSSCHECKING, total=1)
        package = reconcile_package(workbook, deck, profile.crosscheck)
        check_cancelled(cancellation_token)
        report_progress(
            on_progress,
            RunPhase.CROSSCHECKING,
            processed=1,
            total=1,
        )
        findings.extend(excel_preflight.findings)
        findings.extend(ppt_preflight.findings)
        findings.extend(package.findings)
        result.coverage = [
            *excel_preflight.coverage,
            _workload_coverage(workbook),
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
        findings = _apply_comparison_scope(
            result,
            findings,
            requested_scope,
            workbooks=(workbook,),
            decks=(deck,),
        )
        findings = _apply_findings_budget(findings, result.coverage)
        _enrich_retained_findings(
            findings,
            current_workbook=workbook,
            current_deck=deck,
            dependency_graph=excel_preflight.dependency_graph,
            crosscheck=profile.crosscheck,
        )
        result.findings = triage(findings, profile)
        annotate_story_evidence(result.findings)
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
    if run_acceptance is not None and (
        run_acceptance.absolute > 0 or run_acceptance.relative > 0
    ):
        bounds = []
        if run_acceptance.absolute > 0:
            bounds.append(f"±{run_acceptance.absolute:g} absolute")
        if run_acceptance.relative > 0:
            bounds.append(f"±{run_acceptance.relative * 100:g}% relative")
        result.disclosures.append(
            "Analyst acceptance threshold active: numeric differences within "
            + " or ".join(bounds)
            + " are reported as within-tolerance (Info), not suppressed"
        )
    else:
        run_acceptance = None
    findings: list[Finding] = []
    current_workbook = None
    baseline_workbook: WorkbookSnapshot | None = None
    dependency_graph: DependencyGraph | None = None
    loaded_workbooks: list[WorkbookSnapshot] = []
    loaded_decks: list[DeckSnapshot] = []

    if baseline_excel is not None and current_excel is not None:
        base_wb = _load_excel_file(
            baseline_excel,
            password=passwords.get(baseline_excel.name),
            phase=RunPhase.LOADING_BASELINE_EXCEL,
            allow_large_workbooks=allow_large_workbooks,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        curr_wb = _load_excel_file(
            current_excel,
            password=passwords.get(current_excel.name),
            phase=RunPhase.LOADING_CURRENT_EXCEL,
            allow_large_workbooks=allow_large_workbooks,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        current_workbook = curr_wb
        loaded_workbooks.extend((base_wb, curr_wb))
        result.files["baseline_excel"] = baseline_excel.name
        result.files["current_excel"] = current_excel.name
        result.coverage.append(_workload_coverage(base_wb, curr_wb))
        risk_findings = workbook_risk_findings(curr_wb, base_wb)
        findings.extend(risk_findings)
        result.coverage.append(
            CoverageItem(
                check_id="excel-workbook-risks",
                label="Workbook package risks",
                artifact="excel",
                state=CoverageState.CHECKED,
                findings=len(risk_findings),
                detail="Structural package risk inventory compared across the pair",
            )
        )
        result.coverage.append(defined_name_scope_coverage(base_wb, curr_wb))
        vba_findings = diff_workbook_vba(base_wb, curr_wb)
        findings.extend(vba_findings)
        result.coverage.append(vba_coverage(base_wb, curr_wb))
        findings.extend(diff_workbook_metadata(base_wb, curr_wb))
        findings.extend(external_connection_findings(curr_wb))
        result.coverage.append(comment_coverage(base_wb, curr_wb))
        result.coverage.append(power_query_coverage(base_wb, curr_wb))
        result.coverage.append(connection_coverage(base_wb, curr_wb))
        if disclosure := _pair_formula_disclosure(base_wb, curr_wb):
            result.disclosures.append(disclosure)
        report_progress(on_progress, RunPhase.ANALYZING_EXCEL, total=1)
        report_progress(on_progress, RunPhase.DIFFING_EXCEL, total=1)
        alignment = align_workbooks(
            base_wb,
            curr_wb,
            profile,
            cancellation_token=cancellation_token,
        )
        check_cancelled(cancellation_token)
        alignment_detail = (
            "Low-confidence key alignment skipped cell-level comparison for: "
            + ", ".join(alignment.low_confidence_regions)
            if alignment.low_confidence_regions
            else ""
        )
        value_start = len(findings)
        findings += diff_workbook_values(
            base_wb,
            curr_wb,
            alignment,
            profile,
            run_acceptance=run_acceptance,
            cancellation_token=cancellation_token,
        )
        check_cancelled(cancellation_token)
        result.coverage.append(
            CoverageItem(
                check_id="excel-values",
                label="Excel values and presentation",
                artifact="excel",
                state=(
                    CoverageState.DEGRADED
                    if alignment.low_confidence_regions
                    else CoverageState.CHECKED
                ),
                findings=len(findings) - value_start,
                detail=alignment_detail,
            )
        )
        structure_start = len(findings)
        structure_findings = diff_workbook_structure(
            base_wb,
            curr_wb,
            alignment,
            profile,
        )
        check_cancelled(cancellation_token)
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
                    for sheet_name, sheet in profile.excel.sheets.items()
                    if sheet_name not in profile.excel.ignore_sheets
                    and not sheet.ignore
                ),
                issues=excel_availability_issues(curr_wb, profile),
            )
        )
        formula_start = len(findings)
        report_progress(
            on_progress, RunPhase.DIFFING_EXCEL, processed=1, total=1
        )
        report_progress(on_progress, RunPhase.COMPARING_FORMULAS, total=1)
        findings += diff_workbook_formulas(
            base_wb,
            curr_wb,
            alignment,
            profile,
            cancellation_token=cancellation_token,
        )
        check_cancelled(cancellation_token)
        report_progress(
            on_progress, RunPhase.COMPARING_FORMULAS, processed=1, total=1
        )
        formula_text_checked = formula_text_compatible(base_wb, curr_wb)
        formula_complete = (
            formula_text_checked and not alignment.low_confidence_regions
        )
        result.coverage.append(
            CoverageItem(
                check_id="excel-formulas",
                label="Excel formulas",
                artifact="excel",
                state=(
                    CoverageState.CHECKED
                    if formula_complete
                    else CoverageState.DEGRADED
                ),
                findings=len(findings) - formula_start,
                detail=(
                    f"Compatible formula text source: {base_wb.formula_source}"
                    if formula_complete
                    else (
                        alignment_detail
                        if alignment.low_confidence_regions
                        else (
                            "Formula presence checks completed; semantic formula text "
                            "checks unavailable"
                        )
                    )
                ),
            )
        )
        if curr_wb.formulas_available:
            report_progress(on_progress, RunPhase.INDEXING_DEPENDENCIES, total=1)
            complexity = assess_workbook_complexity(
                curr_wb,
                allow_complex_workbook=allow_large_workbooks,
                cancellation_token=cancellation_token,
            )
            result.coverage.append(_complexity_coverage(complexity))
            dependency_graph = build_dependency_graph(
                curr_wb,
                cancellation_token=cancellation_token,
            )
            report_progress(
                on_progress, RunPhase.INDEXING_DEPENDENCIES, processed=1, total=1
            )
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
        ignored_sheets = set(profile.excel.ignore_sheets)
        ignored_sheets.update(
            sheet_name
            for sheet_name, sheet_profile in profile.excel.sheets.items()
            if sheet_profile.ignore
        )
        controls = evaluate_controls(
            curr_wb,
            profile.excel.controls,
            ignored_sheets=ignored_sheets,
        )
        findings += controls.findings
        if controls.coverage is not None:
            result.coverage.append(controls.coverage)
        baseline_workbook = base_wb
        check_cancelled(cancellation_token)
        report_progress(
            on_progress,
            RunPhase.ANALYZING_EXCEL,
            processed=1,
            total=1,
        )

    current_deck = None
    if baseline_ppt is not None and current_ppt is not None:
        base_deck = _load_powerpoint_file(
            baseline_ppt,
            password=passwords.get(baseline_ppt.name),
            phase=RunPhase.LOADING_BASELINE_POWERPOINT,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        current_deck = _load_powerpoint_file(
            current_ppt,
            password=passwords.get(current_ppt.name),
            phase=RunPhase.LOADING_CURRENT_POWERPOINT,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        loaded_decks.extend((base_deck, current_deck))
        result.files["baseline_ppt"] = baseline_ppt.name
        result.files["current_ppt"] = current_ppt.name
        report_progress(on_progress, RunPhase.ANALYZING_POWERPOINT, total=1)
        matching = match_slides(base_deck, current_deck, profile.ppt)
        check_cancelled(cancellation_token)
        ppt_start = len(findings)
        findings += diff_decks(matching, profile.ppt)
        check_cancelled(cancellation_token)
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
        report_progress(
            on_progress,
            RunPhase.ANALYZING_POWERPOINT,
            processed=1,
            total=1,
        )
        result.coverage.append(
            availability_coverage(
                artifact="ppt",
                rule_count=len(profile.ppt.availability_rules),
                issues=ppt_availability_issues(current_deck, profile.ppt),
            )
        )

    if baseline_excel is None:
        result.coverage.extend(
            [
                CoverageItem(
                    check_id="excel-workload",
                    label="Excel workload safeguards",
                    artifact="excel",
                    state=CoverageState.UNAVAILABLE,
                    detail="Excel pair not supplied",
                ),
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
        report_progress(on_progress, RunPhase.CROSSCHECKING, total=1)
        crosscheck_start = len(findings)
        crosscheck = verify_mappings(current_deck, current_workbook, profile.crosscheck)
        check_cancelled(cancellation_token)
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
        report_progress(
            on_progress,
            RunPhase.CROSSCHECKING,
            processed=1,
            total=1,
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

    findings = _apply_comparison_scope(
        result,
        findings,
        requested_scope,
        workbooks=tuple(loaded_workbooks),
        decks=tuple(loaded_decks),
    )
    findings = _apply_findings_budget(findings, result.coverage)
    report_progress(on_progress, RunPhase.QUERYING_IMPACTS, total=1)
    _enrich_retained_findings(
        findings,
        baseline_workbook=baseline_workbook,
        current_workbook=current_workbook,
        current_deck=current_deck,
        dependency_graph=dependency_graph,
        crosscheck=profile.crosscheck,
    )
    report_progress(on_progress, RunPhase.QUERYING_IMPACTS, processed=1, total=1)
    report_progress(on_progress, RunPhase.BUILDING_REVIEW, total=1)
    result.findings = triage(findings, profile)
    annotate_story_evidence(result.findings)
    report_progress(on_progress, RunPhase.BUILDING_REVIEW, processed=1, total=1)
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
