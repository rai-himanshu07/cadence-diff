"""QC run orchestrator: load, align, diff, cross-check, triage.

One `run_qc` call is one QC run — a baseline/current workbook pair and/or
a baseline/current deck pair, an optional profile, and per-file passwords.
The result carries triaged findings plus capability disclosures, including
whether XLSB formula text was enriched or only formula presence was checked.
"""

import contextlib
import datetime as dt
import logging
import shutil
import tempfile
import weakref
from collections import Counter
from collections.abc import Iterable, Iterator, Sequence
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
    profile_for_excel_member,
)
from qc_tool.coverage import CoverageItem, CoverageState, MappingCoverage, QCRunMode
from qc_tool.crosscheck.package import MultiPackageReconciler, reconcile_package
from qc_tool.crosscheck.trace import (
    MappingSuggestion,
    annotate_ppt_chart_impacts,
    verify_mappings,
)
from qc_tool.excel.align import (
    AlignmentTrustManifest,
    WorkbookAlignment,
    align_workbooks,
    build_alignment_trust_manifest,
)
from qc_tool.excel.charts import annotate_chart_impacts, chart_reference_coverage
from qc_tool.excel.complexity import WorkbookComplexity, assess_workbook_complexity
from qc_tool.excel.context import attach_current_excerpts, attach_excerpts
from qc_tool.excel.controls import evaluate_controls
from qc_tool.excel.dependency import (
    DependencyGraph,
    ImpactAccumulator,
    build_dependency_graph,
    detect_circular_references,
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
from qc_tool.excel.diff_values import iter_region_findings, region_range_sets
from qc_tool.excel.diff_vba import diff_workbook_vba, vba_coverage
from qc_tool.excel.formulas import diff_workbook_formulas, formula_text_compatible
from qc_tool.excel.interaction import (
    conditional_style_coverage,
    interaction_rule_coverage,
)
from qc_tool.excel.preflight import defined_name_scope_coverage, preflight_workbook
from qc_tool.excel.workbook_risks import workbook_risk_findings
from qc_tool.findings import Finding, FindingClass, Severity
from qc_tool.findings_store import (
    BLOCK_FINDINGS,
    FindingSequence,
    SpillWriter,
    finding_payload,
    merge_spill,
    write_finding_blocks,
)
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.io.model import WorkbookSnapshot
from qc_tool.package import PackageArtifact, PackageManifest, PackageSide, paths_by_member
from qc_tool.ppt.diff import diff_decks
from qc_tool.ppt.element_match import match_slide_elements
from qc_tool.ppt.extract import load_deck_snapshot
from qc_tool.ppt.match import match_slides
from qc_tool.ppt.model import DeckSnapshot
from qc_tool.ppt.preflight import media_structural_coverage, preflight_deck
from qc_tool.progress import (
    CancellationToken,
    ProgressCallback,
    RunPhase,
    check_cancelled,
    report_progress,
)
from qc_tool.scope import ComparisonScope
from qc_tool.story import StoryEvidenceCollector, annotate_story_evidence
from qc_tool.triage.rules import (
    assign_severities,
    expired_waiver_findings,
    root_cause_candidate_key,
    triage,
    triage_sort_key,
)

logger = logging.getLogger(__name__)


def _password_for(
    passwords: dict[str, str],
    role: str,
    path: Path,
) -> str | None:
    return passwords.get(role, passwords.get(path.name))


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
    measured = [workbook for workbook in workbooks if workbook.workload.metrics_available]
    if not measured:
        return CoverageItem(
            check_id="excel-workload",
            label="Excel workload safeguards",
            artifact="excel",
            state=CoverageState.UNAVAILABLE,
            detail="Workbook workload metrics are unavailable",
        )
    incomplete = len(measured) != len(workbooks)
    degraded = incomplete or any(workbook.workload.degraded for workbook in measured)
    details = [
        f"{workbook.source_name}: {workbook.workload.detail}"
        for workbook in measured
    ]
    if incomplete:
        unavailable = len(workbooks) - len(measured)
        details.append(f"Workload metrics unavailable for {unavailable} workbook(s)")
    return CoverageItem(
        check_id="excel-workload",
        label="Excel workload safeguards",
        artifact="excel",
        state=CoverageState.DEGRADED if degraded else CoverageState.CHECKED,
        detail="; ".join(details),
    )


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
        accumulator = ImpactAccumulator(dependency_graph)
        accumulator.annotate(findings)
    else:
        accumulator = None
    if current_workbook is not None:
        annotate_chart_impacts(findings, current_workbook, dependency_graph)
    if (
        current_deck is not None
        and dependency_graph is not None
        and crosscheck is not None
        and crosscheck.mappings
    ):
        annotate_ppt_chart_impacts(findings, current_deck, crosscheck, dependency_graph)
    if accumulator is not None:
        accumulator.finalize(findings)
    else:
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


class _FindingStream:
    """Spill-backed accumulator producing the triaged findings sequence.

    Findings are spilled in production order under their global triage sort
    key; ``finalize`` merges them in key order, assigns ``F%04d`` ids, applies
    the root-cause counter and story evidence tags, and returns a lazy
    ``FindingSequence`` byte-equal to what ``triage`` plus
    ``annotate_story_evidence`` produce on an in-memory list.
    """

    def __init__(self) -> None:
        self._dir = Path(tempfile.mkdtemp(prefix="qc-findings-"))
        self._closed = False
        self._spill = SpillWriter(self._dir / "spill.qcfb")
        self._spill.__enter__()
        self._ordinal = 0
        self._collector = StoryEvidenceCollector()
        self._root_counts: Counter[str] = Counter()

    def add(self, findings: Iterable[Finding]) -> None:
        """Spill severity-assigned, enriched findings in production order."""
        for finding in findings:
            self._collector.observe(finding)
            key = root_cause_candidate_key(finding)
            if key is not None:
                self._root_counts[key] += 1
            self._spill.append(
                (*triage_sort_key(finding), self._ordinal),
                finding_payload(finding),
            )
            self._ordinal += 1

    def finalize(self) -> tuple[FindingSequence, dict[Severity, int]]:
        if self._closed:
            raise RuntimeError("finding stream is already closed")
        self._spill.__exit__(None, None, None)
        context = self._collector.context()
        root_counts = self._root_counts
        counts = dict.fromkeys(Severity, 0)

        def _finalized() -> Iterator[object]:
            for index, payload in enumerate(merge_spill(self._spill.path), start=1):
                finding = Finding.model_validate(payload)
                finding.finding_id = f"F{index:04d}"
                key = root_cause_candidate_key(finding)
                if key is not None and root_counts[key] > 1:
                    finding.root_cause_key = key
                context.apply(finding)
                if finding.severity is not None:
                    counts[finding.severity] += 1
                yield finding_payload(finding)

        container = write_finding_blocks(self._dir / "findings.qcfb", _finalized())
        self._spill.path.unlink(missing_ok=True)
        sequence = FindingSequence(container)
        # The sequence reads lazily, so the directory lives exactly as long
        # as the result that exposes it (atexit covers process shutdown).
        weakref.finalize(sequence, shutil.rmtree, str(self._dir), True)
        self._closed = True
        return sequence, counts

    def abort(self) -> None:
        if self._closed:
            return
        with contextlib.suppress(Exception):
            self._spill.__exit__(None, None, None)
        shutil.rmtree(self._dir, ignore_errors=True)
        self._closed = True

    def __del__(self) -> None:
        with contextlib.suppress(Exception):
            self.abort()


def _process_cycle_chunk(
    chunk: list[Finding],
    *,
    stream: _FindingStream,
    scope: ComparisonScope,
    profile: DeliverableProfile,
    today: dt.date,
    baseline_workbook: WorkbookSnapshot,
    current_workbook: WorkbookSnapshot,
    current_deck: DeckSnapshot | None,
    dependency_graph: DependencyGraph | None,
) -> None:
    """Scope-filter, triage-assign, enrich, and spill one part chunk."""
    retained = scope.filter_findings(chunk)
    if not retained:
        return
    assign_severities(retained, profile, today=today)
    _enrich_retained_findings(
        retained,
        baseline_workbook=baseline_workbook,
        current_workbook=current_workbook,
        current_deck=current_deck,
        dependency_graph=dependency_graph,
        crosscheck=profile.crosscheck,
    )
    stream.add(retained)


def _run_value_parts(
    *,
    stream: _FindingStream,
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    alignment: WorkbookAlignment,
    profile: DeliverableProfile,
    scope: ComparisonScope,
    run_acceptance: NumericTolerance | None,
    today: dt.date,
    current_deck: DeckSnapshot | None,
    dependency_graph: DependencyGraph | None,
    cancellation_token: CancellationToken | None,
    on_progress: ProgressCallback | None,
) -> int:
    """Region-batched value diff with per-sheet memory release.

    Feeds each aligned region's findings in bounded chunks through
    ``_process_cycle_chunk`` and deletes a sheet's ``CellRecord`` dicts from
    both snapshots once its last region is done, so peak memory follows the
    largest chunk instead of the whole findings population. Returns the
    pre-scope-filter production count for the values coverage row.
    """
    tolerance = profile.tolerance
    windows = profile.restatement_windows
    part_total = sum(len(regions) for regions in alignment.regions.values())
    report_progress(
        on_progress, RunPhase.DIFFING_EXCEL, total=max(part_total, 1)
    )
    if part_total == 0:
        report_progress(on_progress, RunPhase.DIFFING_EXCEL, processed=1, total=1)
        return 0
    produced = 0
    part_index = 0
    for sheet_name, regions in alignment.regions.items():
        check_cancelled(cancellation_token)
        base_sheet = baseline.sheet(sheet_name)
        curr_sheet = current.sheet(sheet_name)
        sheet_profile = profile.sheet_profile(sheet_name)
        ignore, refresh = region_range_sets(sheet_profile)
        for region in regions:
            check_cancelled(cancellation_token)
            chunk: list[Finding] = []
            for finding in iter_region_findings(
                base_sheet,
                curr_sheet,
                region,
                tolerance,
                ignore=ignore,
                refresh=refresh,
                sheet_profile=sheet_profile,
                windows=windows,
                run_acceptance=run_acceptance,
            ):
                produced += 1
                chunk.append(finding)
                if len(chunk) >= BLOCK_FINDINGS:
                    _process_cycle_chunk(
                        chunk,
                        stream=stream,
                        scope=scope,
                        profile=profile,
                        today=today,
                        baseline_workbook=baseline,
                        current_workbook=current,
                        current_deck=current_deck,
                        dependency_graph=dependency_graph,
                    )
                    chunk = []
                    check_cancelled(cancellation_token)
            if chunk:
                _process_cycle_chunk(
                    chunk,
                    stream=stream,
                    scope=scope,
                    profile=profile,
                    today=today,
                    baseline_workbook=baseline,
                    current_workbook=current,
                    current_deck=current_deck,
                    dependency_graph=dependency_graph,
                )
            part_index += 1
            report_progress(
                on_progress,
                RunPhase.DIFFING_EXCEL,
                processed=part_index,
                total=part_total,
                detail=sheet_name,
            )
        # Excerpts for every finding on this sheet are attached by now;
        # release the cell records on both sides before the next sheet.
        base_sheet.cells.clear()
        curr_sheet.cells.clear()
    return produced


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
    findings: Sequence[Finding] = field(default_factory=list)
    disclosures: list[str] = field(default_factory=list)
    coverage: list[CoverageItem] = field(default_factory=list)
    mapping_coverage: MappingCoverage | None = None
    mapping_suggestions: list[MappingSuggestion] = field(default_factory=list)
    verified_crosschecks: int = 0
    comparison_scope: ComparisonScope = field(default_factory=ComparisonScope)
    alignment_trust: AlignmentTrustManifest | None = None
    package_manifest: PackageManifest | None = None
    #: Filled by the streaming cycle path so `counts` never re-reads the store.
    severity_counts: dict[Severity, int] | None = None

    @property
    def counts(self) -> dict[Severity, int]:
        if self.severity_counts is not None:
            return {
                severity: self.severity_counts.get(severity, 0)
                for severity in Severity
            }
        result = dict.fromkeys(Severity, 0)
        for finding in self.findings:
            if finding.severity is not None:
                result[finding.severity] += 1
        return result


@dataclass(slots=True)
class _SnapshotCapture:
    current_workbook: WorkbookSnapshot | None = None
    current_deck: DeckSnapshot | None = None
    dependency_graph: DependencyGraph | None = None


def _memberize_coverage(
    items: list[CoverageItem],
    member_id: str,
    *,
    artifact: str,
) -> list[CoverageItem]:
    result: list[CoverageItem] = []
    for item in items:
        if item.artifact != artifact:
            continue
        copied = item.model_copy(deep=True)
        copied.artifact_member = member_id
        if member_id != "primary":
            copied.check_id = f"member:{member_id}:{copied.check_id}"
            copied.label = f"{member_id} · {copied.label}"
        result.append(copied)
    return result


def _memberize_excel_result(
    subresult: QCRunResult,
    member_id: str,
) -> tuple[list[Finding], list[CoverageItem], AlignmentTrustManifest | None]:
    findings: list[Finding] = []
    for finding in subresult.findings:
        if finding.artifact != "excel":
            continue
        finding.artifact_member = member_id
        finding.finding_id = ""
        finding.severity_overridden = False
        finding.analyst_comment = ""
        findings.append(finding)
    coverage = _memberize_coverage(
        subresult.coverage,
        member_id,
        artifact="excel",
    )
    trust = subresult.alignment_trust
    if trust is not None and member_id != "primary":
        trust = AlignmentTrustManifest(
            regions=tuple(
                region.model_copy(update={"artifact_member": member_id})
                for region in trust.regions
            ),
            unpaired=tuple(
                region.model_copy(update={"artifact_member": member_id})
                for region in trust.unpaired
            ),
        )
    return findings, coverage, trust


def _merge_alignment_manifests(
    manifests: list[AlignmentTrustManifest],
) -> AlignmentTrustManifest | None:
    if not manifests:
        return None
    return AlignmentTrustManifest(
        regions=tuple(
            region
            for manifest in manifests
            for region in manifest.regions
        ),
        unpaired=tuple(
            region
            for manifest in manifests
            for region in manifest.unpaired
        ),
    )


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


def _run_multi_package(
    *,
    manifest: PackageManifest,
    files: dict[str, Path],
    profile: DeliverableProfile,
    passwords: dict[str, str],
    mode: QCRunMode,
    allow_large_workbooks: bool,
    run_acceptance: NumericTolerance | None,
    compare_sheets: list[str] | None,
    compare_member_sheets: dict[str, tuple[str, ...]],
    compare_slides: list[int] | None,
    cancellation_token: CancellationToken | None,
    on_progress: ProgressCallback | None,
) -> QCRunResult:
    """Run existing single-artifact pipelines sequentially, then merge once."""
    paths_by_member(files, manifest)
    baseline_members = {
        member.member_id: member
        for member in manifest.members_for(
            PackageSide.BASELINE,
            PackageArtifact.EXCEL,
        )
    }
    current_members = {
        member.member_id: member
        for member in manifest.members_for(
            PackageSide.CURRENT,
            PackageArtifact.EXCEL,
        )
    }
    baseline_ppt = manifest.members_for(
        PackageSide.BASELINE,
        PackageArtifact.PPT,
    )
    current_ppt = manifest.members_for(
        PackageSide.CURRENT,
        PackageArtifact.PPT,
    )
    if mode is QCRunMode.CURRENT_FILE_PREFLIGHT:
        if baseline_members or baseline_ppt:
            raise ValueError("current-file preflight does not accept baseline members")
        if not current_members and not current_ppt:
            raise ValueError("current-file preflight needs a current package member")
    elif mode is QCRunMode.CYCLE_COMPARISON:
        excel_pair = bool(baseline_members and current_members)
        ppt_pair = bool(baseline_ppt and current_ppt)
        if not excel_pair and not ppt_pair:
            raise ValueError(
                "cycle comparison needs baseline/current Excel members or PPT decks"
            )
    elif mode is QCRunMode.FINAL_PACKAGE:
        if baseline_members or baseline_ppt:
            raise ValueError("final-package QC does not accept baseline members")
        if not current_members or len(current_ppt) != 1:
            raise ValueError(
                "final-package QC needs current Excel member(s) and PowerPoint"
            )
    current_count = len(current_members)
    coverage: list[CoverageItem] = []
    disclosures: list[str] = []
    trust_manifests: list[AlignmentTrustManifest] = []
    current_deck_snapshot: DeckSnapshot | None = None
    package_reconciler: MultiPackageReconciler | None = None
    package_result = None
    member_events = 0
    stream = _FindingStream()
    today = dt.date.today()

    def member_passwords(*role_pairs: tuple[str, str]) -> dict[str, str]:
        return {
            scalar_role: passwords[source_role]
            for scalar_role, source_role in role_pairs
            if passwords.get(source_role)
        }

    def projected_profile(member_id: str) -> DeliverableProfile:
        projected = profile_for_excel_member(
            profile,
            member_id,
            max(current_count, 1),
        )
        # Waivers are applied once after member identity is attached.
        projected.waivers = []
        projected.crosscheck.mappings = []
        return projected

    def add_findings(items: Iterable[Finding]) -> None:
        batch = list(items)
        if not batch:
            return
        assign_severities(batch, profile, today=today)
        stream.add(batch)

    all_member_ids = sorted(set(baseline_members) | set(current_members))
    unknown_scope_members = sorted(
        set(compare_member_sheets) - set(current_members)
    )
    if unknown_scope_members:
        raise ValueError(
            "unknown workbook members in scope: "
            + ", ".join(unknown_scope_members)
        )

    # Final-package reconciliation needs the deck while each workbook is
    # resident. Load it first, then observe/release one workbook at a time.
    if mode is QCRunMode.FINAL_PACKAGE:
        current = current_ppt[0]
        capture = _SnapshotCapture()
        ppt_result = run_qc(
            current_ppt=files[current.role_key],
            profile=profile.model_copy(update={"waivers": []}, deep=True),
            passwords=member_passwords(("current_ppt", current.role_key)),
            mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
            compare_slides=compare_slides,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
            _snapshot_capture=capture,
        )
        current_deck_snapshot = capture.current_deck
        if current_deck_snapshot is None:
            raise RuntimeError("final-package PowerPoint snapshot was not captured")
        add_findings(
            finding
            for finding in ppt_result.findings
            if finding.artifact == "ppt"
        )
        coverage.extend(
            _memberize_coverage(ppt_result.coverage, "primary", artifact="ppt")
        )
        disclosures.extend(ppt_result.disclosures)
        package_reconciler = MultiPackageReconciler(
            current_deck_snapshot,
            profile.crosscheck,
            current_members,
        )
        report_progress(on_progress, RunPhase.CROSSCHECKING, total=1)

    for member_id in all_member_ids:
        check_cancelled(cancellation_token)
        baseline = baseline_members.get(member_id)
        current = current_members.get(member_id)
        member_scope = list(compare_member_sheets.get(member_id, ())) or None
        if member_id == "primary" and member_scope is None:
            member_scope = compare_sheets

        if mode is QCRunMode.CYCLE_COMPARISON and baseline and current:
            subresult = run_qc(
                baseline_excel=files[baseline.role_key],
                current_excel=files[current.role_key],
                profile=projected_profile(member_id),
                passwords=member_passwords(
                    ("baseline_excel", baseline.role_key),
                    ("current_excel", current.role_key),
                ),
                mode=QCRunMode.CYCLE_COMPARISON,
                allow_large_workbooks=allow_large_workbooks,
                run_acceptance=run_acceptance,
                compare_sheets=member_scope,
                cancellation_token=cancellation_token,
                on_progress=on_progress,
            )
            member_findings, member_coverage, member_trust = (
                _memberize_excel_result(subresult, member_id)
            )
            add_findings(member_findings)
            coverage.extend(member_coverage)
            if member_trust is not None:
                trust_manifests.append(member_trust)
            disclosures.extend(
                f"Excel member {member_id}: {detail}"
                for detail in subresult.disclosures
            )
            continue

        if current is not None:
            capture = (
                _SnapshotCapture() if mode is QCRunMode.FINAL_PACKAGE else None
            )
            subresult = run_qc(
                current_excel=files[current.role_key],
                profile=projected_profile(member_id),
                passwords=member_passwords(("current_excel", current.role_key)),
                mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
                allow_large_workbooks=allow_large_workbooks,
                compare_sheets=member_scope,
                cancellation_token=cancellation_token,
                on_progress=on_progress,
                _snapshot_capture=capture,
            )
            member_findings, member_coverage, _member_trust = (
                _memberize_excel_result(subresult, member_id)
            )
            if mode is QCRunMode.FINAL_PACKAGE:
                if (
                    capture is None
                    or capture.current_workbook is None
                    or package_reconciler is None
                    or current_deck_snapshot is None
                ):
                    raise RuntimeError(
                        f"final-package workbook {member_id!r} was not captured"
                    )
                package_reconciler.observe_member(
                    member_id,
                    capture.current_workbook,
                )
                if capture.dependency_graph is not None:
                    member_profile = CrosscheckProfile(
                        mappings=[
                            mapping
                            for mapping in profile.crosscheck.mappings
                            if mapping.source_member == member_id
                        ],
                        max_candidates=profile.crosscheck.max_candidates,
                    )
                    annotate_ppt_chart_impacts(
                        member_findings,
                        current_deck_snapshot,
                        member_profile,
                        capture.dependency_graph,
                    )
            add_findings(member_findings)
            coverage.extend(member_coverage)
            disclosures.extend(
                f"Excel member {member_id}: {detail}"
                for detail in subresult.disclosures
            )
            if mode is QCRunMode.CYCLE_COMPARISON:
                add_findings(
                    [
                    Finding(
                        artifact="excel",
                        artifact_member=member_id,
                        finding_class=FindingClass.WORKBOOK_ADDED,
                        element=current.display_name,
                        message=(
                            f"Excel workbook member {member_id!r} was added "
                            "in the current package"
                        ),
                    )
                    ]
                )
                member_events += 1
            continue

        if baseline is not None and mode is QCRunMode.CYCLE_COMPARISON:
            add_findings(
                [
                Finding(
                    artifact="excel",
                    artifact_member=member_id,
                    finding_class=FindingClass.WORKBOOK_REMOVED,
                    element=baseline.display_name,
                    message=(
                        f"Excel workbook member {member_id!r} is missing "
                        "from the current package"
                    ),
                )
                ]
            )
            member_events += 1

    if mode is QCRunMode.CYCLE_COMPARISON and baseline_ppt and current_ppt:
        baseline = baseline_ppt[0]
        current = current_ppt[0]
        ppt_result = run_qc(
            baseline_ppt=files[baseline.role_key],
            current_ppt=files[current.role_key],
            profile=profile.model_copy(update={"waivers": []}, deep=True),
            passwords=member_passwords(
                ("baseline_ppt", baseline.role_key),
                ("current_ppt", current.role_key),
            ),
            mode=QCRunMode.CYCLE_COMPARISON,
            compare_slides=compare_slides,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        add_findings(
            finding for finding in ppt_result.findings if finding.artifact == "ppt"
        )
        coverage.extend(
            _memberize_coverage(ppt_result.coverage, "primary", artifact="ppt")
        )
        disclosures.extend(ppt_result.disclosures)
    elif current_ppt and mode is not QCRunMode.FINAL_PACKAGE:
        current = current_ppt[0]
        ppt_result = run_qc(
            current_ppt=files[current.role_key],
            profile=profile.model_copy(update={"waivers": []}, deep=True),
            passwords=member_passwords(("current_ppt", current.role_key)),
            mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
            compare_slides=compare_slides,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        add_findings(
            finding for finding in ppt_result.findings if finding.artifact == "ppt"
        )
        coverage.extend(
            _memberize_coverage(ppt_result.coverage, "primary", artifact="ppt")
        )
        disclosures.extend(ppt_result.disclosures)

    coverage.append(
        CoverageItem(
            check_id="excel-package-members",
            label="Excel package member pairing",
            artifact="package",
            state=CoverageState.CHECKED,
            findings=member_events,
            detail=(
                f"{len(baseline_members)} baseline and {len(current_members)} "
                "current Excel members paired by stable member ID"
            ),
        )
    )
    scoped_sheet_count = sum(len(sheets) for sheets in compare_member_sheets.values())
    unscoped_members = sorted(set(current_members) - set(compare_member_sheets))
    scope_details = [
        f"{len(current_members)} current Excel members loaded fully",
        (
            f"{scoped_sheet_count} explicitly selected member sheets"
            + (
                f"; {len(unscoped_members)} unlisted member(s) unrestricted"
                if unscoped_members
                else ""
            )
            if scoped_sheet_count
            else "all member sheets compared"
        ),
        "package mappings, suggestions, and period reconciliation use the full package",
    ]
    if compare_slides:
        scope_details.append(f"{len(compare_slides)} selected PowerPoint slides")
    coverage.append(
        CoverageItem(
            check_id="comparison-scope",
            label="Validated comparison scope",
            artifact="run",
            state=CoverageState.CHECKED,
            detail="; ".join(scope_details),
        )
    )
    if mode is QCRunMode.FINAL_PACKAGE and package_reconciler is not None:
        package_result = package_reconciler.finish()
        add_findings(package_result.findings)
        coverage.extend(package_result.coverage)
        report_progress(
            on_progress,
            RunPhase.CROSSCHECKING,
            processed=1,
            total=1,
        )
    elif current_ppt:
        coverage.append(
            CoverageItem(
                check_id="excel-ppt-crosscheck",
                label="Excel to PowerPoint mappings",
                artifact="package",
                state=CoverageState.UNAVAILABLE,
                detail="Use final-package mode for Excel-to-PowerPoint QC",
            )
        )

    add_findings(expired_waiver_findings(profile, today))
    sequence, severity_counts = stream.finalize()
    result = QCRunResult(
        profile_name=profile.name,
        mode=mode,
        files={member.role_key: member.display_name for member in manifest.members},
        disclosures=list(dict.fromkeys(disclosures)),
        coverage=coverage,
        package_manifest=manifest,
        comparison_scope=ComparisonScope(
            excel_member_sheets=(compare_member_sheets or None),
            ppt_slide_indices=(
                tuple(compare_slides) if compare_slides is not None else None
            ),
        ),
        alignment_trust=_merge_alignment_manifests(trust_manifests),
        mapping_coverage=(
            package_result.mapping_coverage if package_result is not None else None
        ),
        mapping_suggestions=(
            package_result.suggestions if package_result is not None else []
        ),
        verified_crosschecks=(
            package_result.mapping_coverage.verified
            if package_result is not None
            else 0
        ),
        findings=sequence,
        severity_counts=severity_counts,
    )
    if disclosure := result.comparison_scope.disclosure():
        result.disclosures.append(disclosure)
    return result


def run_qc(
    *,
    baseline_excel: Path | None = None,
    current_excel: Path | None = None,
    baseline_ppt: Path | None = None,
    current_ppt: Path | None = None,
    profile: DeliverableProfile | None = None,
    package_manifest: PackageManifest | None = None,
    package_files: dict[str, Path] | None = None,
    compare_member_sheets: dict[str, tuple[str, ...]] | None = None,
    passwords: dict[str, str] | None = None,
    mode: QCRunMode = QCRunMode.CYCLE_COMPARISON,
    allow_large_workbooks: bool = False,
    run_acceptance: NumericTolerance | None = None,
    compare_sheets: list[str] | None = None,
    compare_slides: list[int] | None = None,
    cancellation_token: CancellationToken | None = None,
    on_progress: ProgressCallback | None = None,
    _snapshot_capture: _SnapshotCapture | None = None,
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
        excel_member_sheets=compare_member_sheets or None,
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
    # Build a legacy manifest from any scalar file args (display names only)
    legacy_roles: dict[str, object] = {}
    for role, path in (
        ("baseline_excel", baseline_excel),
        ("current_excel", current_excel),
        ("baseline_ppt", baseline_ppt),
        ("current_ppt", current_ppt),
    ):
        if path is not None:
            legacy_roles[role] = path
    manifest: PackageManifest | None = None
    if legacy_roles:
        manifest = PackageManifest.from_legacy_files(legacy_roles)
    if package_manifest is not None:
        manifest = package_manifest
        if not package_manifest.is_legacy_projection:
            if package_files is None:
                raise ValueError("multi-workbook execution needs package file paths")
            return _run_multi_package(
                manifest=package_manifest,
                files=package_files,
                profile=profile or default_profile(),
                passwords=passwords or {},
                mode=mode,
                allow_large_workbooks=allow_large_workbooks,
                run_acceptance=run_acceptance,
                compare_sheets=compare_sheets,
                compare_member_sheets=compare_member_sheets or {},
                compare_slides=compare_slides,
                cancellation_token=cancellation_token,
                on_progress=on_progress,
            )
    if mode is QCRunMode.CURRENT_FILE_PREFLIGHT:
        if baseline_excel is not None or baseline_ppt is not None:
            raise ValueError("current-file preflight does not accept baseline files")
        if current_excel is None and current_ppt is None:
            raise ValueError("current-file preflight needs a current Excel or PowerPoint file")
        profile = profile or default_profile()
        passwords = passwords or {}
        result = QCRunResult(profile_name=profile.name, mode=mode)
        result.package_manifest = manifest
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
                password=_password_for(passwords, "current_excel", current_excel),
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
            if _snapshot_capture is not None:
                _snapshot_capture.current_workbook = workbook
                _snapshot_capture.dependency_graph = preflight_dependency_graph
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
                password=_password_for(passwords, "current_ppt", current_ppt),
                phase=RunPhase.LOADING_CURRENT_POWERPOINT,
                cancellation_token=cancellation_token,
                on_progress=on_progress,
            )
            result.files["current_ppt"] = current_ppt.name
            report_progress(on_progress, RunPhase.ANALYZING_POWERPOINT, total=1)
            ppt_preflight = preflight_deck(deck, profile.ppt)
            preflight_deck_snapshot = deck
            if _snapshot_capture is not None:
                _snapshot_capture.current_deck = deck
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
            result.coverage.extend(
                [
                    CoverageItem(
                        check_id="ppt-intrinsic",
                        label="Current PowerPoint intrinsic checks",
                        artifact="ppt",
                        state=CoverageState.UNAVAILABLE,
                        detail="No current deck supplied",
                    ),
                    media_structural_coverage(),
                    CoverageItem(
                        check_id="ppt-media-visual",
                        label="Rendered media and visual layout",
                        artifact="ppt",
                        state=CoverageState.UNAVAILABLE,
                        detail="No current deck supplied",
                    ),
                ]
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
            password=_password_for(passwords, "current_excel", current_excel),
            phase=RunPhase.LOADING_CURRENT_EXCEL,
            allow_large_workbooks=allow_large_workbooks,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        deck = _load_powerpoint_file(
            current_ppt,
            password=_password_for(passwords, "current_ppt", current_ppt),
            phase=RunPhase.LOADING_CURRENT_POWERPOINT,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        result = QCRunResult(profile_name=profile.name, mode=mode)
        result.package_manifest = manifest
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
    result.package_manifest = manifest
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
    today = dt.date.today()
    pre_findings: list[Finding] = []
    post_batches: list[list[Finding]] = []
    current_workbook = None
    baseline_workbook: WorkbookSnapshot | None = None
    dependency_graph: DependencyGraph | None = None
    loaded_workbooks: list[WorkbookSnapshot] = []
    loaded_decks: list[DeckSnapshot] = []
    alignment: WorkbookAlignment | None = None
    values_coverage: CoverageItem | None = None

    # Decks load before the Excel part loop so chunk enrichment can attach
    # PPT chart impacts while every sheet's cells are still resident.
    base_deck: DeckSnapshot | None = None
    current_deck: DeckSnapshot | None = None
    if baseline_ppt is not None and current_ppt is not None:
        base_deck = _load_powerpoint_file(
            baseline_ppt,
            password=_password_for(passwords, "baseline_ppt", baseline_ppt),
            phase=RunPhase.LOADING_BASELINE_POWERPOINT,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        current_deck = _load_powerpoint_file(
            current_ppt,
            password=_password_for(passwords, "current_ppt", current_ppt),
            phase=RunPhase.LOADING_CURRENT_POWERPOINT,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        loaded_decks.extend((base_deck, current_deck))

    if baseline_excel is not None and current_excel is not None:
        base_wb = _load_excel_file(
            baseline_excel,
            password=_password_for(passwords, "baseline_excel", baseline_excel),
            phase=RunPhase.LOADING_BASELINE_EXCEL,
            allow_large_workbooks=allow_large_workbooks,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        curr_wb = _load_excel_file(
            current_excel,
            password=_password_for(passwords, "current_excel", current_excel),
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
        pre_findings.extend(risk_findings)
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
        pre_findings.extend(vba_findings)
        result.coverage.append(vba_coverage(base_wb, curr_wb))
        pre_findings.extend(diff_workbook_metadata(base_wb, curr_wb))
        pre_findings.extend(external_connection_findings(curr_wb))
        result.coverage.append(comment_coverage(base_wb, curr_wb))
        result.coverage.append(power_query_coverage(base_wb, curr_wb))
        result.coverage.append(connection_coverage(base_wb, curr_wb))
        if disclosure := _pair_formula_disclosure(base_wb, curr_wb):
            result.disclosures.append(disclosure)
        report_progress(on_progress, RunPhase.ANALYZING_EXCEL, total=1)

        def _align_tick(index: int, total: int, sheet_name: str) -> None:
            report_progress(
                on_progress,
                RunPhase.ANALYZING_EXCEL,
                processed=index,
                total=total + 1,
                detail=sheet_name,
            )

        alignment = align_workbooks(
            base_wb,
            curr_wb,
            profile,
            cancellation_token=cancellation_token,
            on_sheet=_align_tick,
        )
        result.alignment_trust = build_alignment_trust_manifest(alignment)
        alignment_manifest = result.alignment_trust
        low_confidence = sum(
            region.low_confidence for region in alignment_manifest.regions
        )
        unpaired = len(alignment_manifest.unpaired)
        comparable_pairs = sum(
            region.comparable_cell_pairs for region in alignment_manifest.regions
        )
        skipped_cells = sum(
            region.skipped_low_confidence_cells
            for region in alignment_manifest.regions
        )
        details = [
            f"{len(alignment_manifest.regions)} paired regions",
            f"{comparable_pairs} candidate cell pairs",
            f"{skipped_cells} skipped for low-confidence fallback",
            f"{unpaired} unpaired regions",
        ]
        if requested_scope.excel_sheets:
            details.append(
                "finding output is sheet-scoped; this manifest still describes "
                "every detected region"
            )
        result.coverage.append(
            CoverageItem(
                check_id="excel-alignment-trust",
                label="Excel alignment trust",
                artifact="excel",
                state=(
                    CoverageState.CHECKED
                    if low_confidence == 0 and unpaired == 0
                    else CoverageState.DEGRADED
                ),
                findings=low_confidence + unpaired,
                detail="; ".join(details),
            )
        )
        check_cancelled(cancellation_token)
        alignment_detail = (
            "Low-confidence key alignment skipped cell-level comparison for: "
            + ", ".join(alignment.low_confidence_regions)
            if alignment.low_confidence_regions
            else ""
        )
        # The values coverage row keeps its list position; its findings count
        # is set once the region part loop has produced every value finding.
        values_coverage = CoverageItem(
            check_id="excel-values",
            label="Excel values and presentation",
            artifact="excel",
            state=(
                CoverageState.DEGRADED
                if alignment.low_confidence_regions
                else CoverageState.CHECKED
            ),
            findings=0,
            detail=alignment_detail,
        )
        result.coverage.append(values_coverage)
        structure_findings = diff_workbook_structure(
            base_wb,
            curr_wb,
            alignment,
            profile,
        )
        check_cancelled(cancellation_token)
        post_batches.append(structure_findings)
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
                findings=len(structure_findings),
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
        formula_findings: list[Finding] = []
        check_cancelled(cancellation_token)
        # Close the analysis span before formula comparison opens so the
        # recorded phases stay sequential instead of nesting.
        report_progress(on_progress, RunPhase.ANALYZING_EXCEL, processed=1, total=1)
        report_progress(on_progress, RunPhase.COMPARING_FORMULAS, total=1)
        formula_findings = diff_workbook_formulas(
            base_wb,
            curr_wb,
            alignment,
            profile,
            cancellation_token=cancellation_token,
        )
        check_cancelled(cancellation_token)
        post_batches.append(formula_findings)
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
                findings=len(formula_findings),
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
            dependency_state = dependency_graph.coverage_state
            dependency_detail = dependency_graph.coverage_detail
            circular = detect_circular_references(
                dependency_graph,
                cancellation_token=cancellation_token,
            )
            post_batches.append(list(circular.findings))
            circular_coverage = circular.coverage
            report_progress(
                on_progress, RunPhase.INDEXING_DEPENDENCIES, processed=1, total=1
            )
        else:
            dependency_state = CoverageState.UNAVAILABLE
            dependency_detail = "Formula text is unavailable for dependency extraction"
            circular_coverage = CoverageItem(
                check_id="excel-circular-references",
                label="Circular formula references",
                artifact="excel",
                state=CoverageState.UNAVAILABLE,
                detail="Formula text is unavailable for circular-reference detection",
            )
        result.coverage.append(
            CoverageItem(
                check_id="excel-dependencies",
                label="Formula dependency impact tracing",
                artifact="excel",
                state=dependency_state,
                detail=dependency_detail,
            )
        )
        result.coverage.append(circular_coverage)
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
        post_batches.append(controls.findings)
        if controls.coverage is not None:
            result.coverage.append(controls.coverage)
        baseline_workbook = base_wb
        check_cancelled(cancellation_token)

    if (
        base_deck is not None
        and current_deck is not None
        and baseline_ppt is not None
        and current_ppt is not None
    ):
        result.files["baseline_ppt"] = baseline_ppt.name
        result.files["current_ppt"] = current_ppt.name
        report_progress(on_progress, RunPhase.ANALYZING_POWERPOINT, total=1)
        matching = match_slides(base_deck, current_deck, profile.ppt)
        check_cancelled(cancellation_token)
        ppt_findings = diff_decks(matching, profile.ppt)
        post_batches.append(ppt_findings)
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
                findings=len(ppt_findings),
                detail=(
                    "Complete native chart extraction unavailable for one or both decks"
                    if not base_deck.charts_available
                    or not current_deck.charts_available
                    else "Complete semantic slide-element comparison"
                ),
            )
        )
        result.coverage.append(
            media_structural_coverage(
                base_deck,
                current_deck,
                findings=sum(
                    finding.finding_class is FindingClass.PPT_MEDIA_CHANGED
                    for finding in ppt_findings
                ),
                ambiguous_shapes=sum(
                    len(
                        match_slide_elements(
                            baseline_slide, current_slide
                        ).ambiguous_current_media
                    )
                    for baseline_slide, current_slide in matching.pairs
                ),
            )
        )
        result.coverage.append(
            CoverageItem(
                check_id="ppt-media-visual",
                label="Rendered media and visual layout",
                artifact="ppt",
                state=CoverageState.UNAVAILABLE,
                detail=(
                    "Embedded bytes are checked structurally; pixels, OCR text, "
                    "and rendered layout are not inspected"
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
                media_structural_coverage(),
                CoverageItem(
                    check_id="ppt-media-visual",
                    label="Rendered media and visual layout",
                    artifact="ppt",
                    state=CoverageState.UNAVAILABLE,
                    detail="PowerPoint pair not supplied",
                ),
            ]
        )

    if current_deck is not None and current_workbook is not None and profile.crosscheck.mappings:
        report_progress(on_progress, RunPhase.CROSSCHECKING, total=1)
        crosscheck = verify_mappings(current_deck, current_workbook, profile.crosscheck)
        check_cancelled(cancellation_token)
        post_batches.append(list(crosscheck.findings))
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
                findings=len(crosscheck.findings),
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

    validated_scope = requested_scope.validate_loaded(
        workbooks=tuple(loaded_workbooks),
        decks=tuple(loaded_decks),
    )
    result.comparison_scope = validated_scope
    result.coverage.append(
        validated_scope.coverage_item(
            workbooks=tuple(loaded_workbooks),
            decks=tuple(loaded_decks),
        )
    )
    if disclosure := validated_scope.disclosure():
        result.disclosures.append(disclosure)
    pre_findings = validated_scope.filter_findings(pre_findings)
    post_batches = [
        validated_scope.filter_findings(batch) for batch in post_batches
    ]

    # Non-values findings are bounded; assign severities and enrich them
    # BEFORE the part loop so their excerpts see every sheet's cells.
    report_progress(on_progress, RunPhase.QUERYING_IMPACTS, total=1)
    held = [*pre_findings, *(f for batch in post_batches for f in batch)]
    assign_severities(held, profile, today=today)
    _enrich_retained_findings(
        held,
        baseline_workbook=baseline_workbook,
        current_workbook=current_workbook,
        current_deck=current_deck,
        dependency_graph=dependency_graph,
        crosscheck=profile.crosscheck,
    )
    report_progress(on_progress, RunPhase.QUERYING_IMPACTS, processed=1, total=1)

    stream = _FindingStream()
    try:
        stream.add(pre_findings)
        if (
            alignment is not None
            and baseline_workbook is not None
            and current_workbook is not None
        ):
            produced = _run_value_parts(
                stream=stream,
                baseline=baseline_workbook,
                current=current_workbook,
                alignment=alignment,
                profile=profile,
                scope=validated_scope,
                run_acceptance=run_acceptance,
                today=today,
                current_deck=current_deck,
                dependency_graph=dependency_graph,
                cancellation_token=cancellation_token,
                on_progress=on_progress,
            )
            if values_coverage is not None:
                values_coverage.findings = produced
        for batch in post_batches:
            stream.add(batch)
        expired = expired_waiver_findings(profile, today)
        assign_severities(expired, profile, today=today)
        stream.add(expired)
        report_progress(on_progress, RunPhase.BUILDING_REVIEW, total=1)
        sequence, severity_counts = stream.finalize()
    except BaseException:
        stream.abort()
        raise
    result.findings = sequence
    result.severity_counts = severity_counts
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


def compare_findings(
    previous: Sequence[Finding], current: Sequence[Finding]
) -> FindingsDelta:
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
