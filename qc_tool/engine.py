"""QC run orchestrator: load, align, diff, cross-check, triage.

One `run_qc` call is one QC run — a baseline/current workbook pair and/or
a baseline/current deck pair, an optional profile, and per-file passwords.
The result carries triaged findings plus capability disclosures, including
whether XLSB formula text was enriched or only formula presence was checked.
"""

import contextlib
import datetime as dt
import hashlib
import logging
import shutil
import tempfile
import time
import weakref
from collections import Counter
from collections.abc import Iterable, Iterator, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal

from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple

from qc_tool.availability import (
    availability_coverage,
    excel_availability_issues,
    ppt_availability_issues,
)
from qc_tool.config.profile import (
    CrosscheckProfile,
    DeliverableProfile,
    NumericTolerance,
    ResolvedOutputPolicy,
    default_profile,
    profile_for_excel_member,
    resolve_output_policy,
)
from qc_tool.coverage import (
    CoverageItem,
    CoverageState,
    FindingOutputMode,
    MappingCoverage,
    QCRunMode,
)
from qc_tool.crosscheck.package import MultiPackageReconciler, reconcile_package
from qc_tool.crosscheck.trace import (
    MappingSuggestion,
    annotate_ppt_chart_impacts,
    verify_mappings,
)
from qc_tool.excel.align import (
    AlignmentRegionTrustV2,
    AlignmentTrustManifest,
    AlignmentTrustManifestV2,
    AlignmentTrustPayload,
    WorkbookAlignment,
    align_workbooks,
    build_alignment_trust_manifest,
    promote_region_trust_to_v2,
)
from qc_tool.excel.charts import annotate_chart_impacts, chart_reference_coverage
from qc_tool.excel.complexity import (
    WorkbookComplexity,
    assess_workbook_complexity,
    dependency_index_skip_reason,
)
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
from qc_tool.excel.formulas import (
    FormulaComparisonTelemetry,
    FormulaPairAnalysisMemo,
    PairKeyTelemetry,
    diff_workbook_formulas,
    formula_text_compatible,
    to_r1c1,
)
from qc_tool.excel.interaction import (
    conditional_style_coverage,
    interaction_rule_coverage,
)
from qc_tool.excel.population import (
    CandidateSpill,
    ClassPopulationStats,
    PopulationTelemetry,
    finalize_populations,
)
from qc_tool.excel.preflight import defined_name_scope_coverage, preflight_workbook
from qc_tool.excel.prerequisites import check_comparison_prerequisites
from qc_tool.excel.ranked_identity import detect_ranked_table_candidate
from qc_tool.excel.regions import internal_period_band_suggestions
from qc_tool.excel.workbook_risks import (
    external_link_reachability_coverage,
    workbook_risk_findings,
)
from qc_tool.findings import Finding, FindingClass, Severity
from qc_tool.findings_store import (
    BLOCK_FINDINGS,
    FindingSequence,
    SpillWriter,
    finding_payload,
    merge_spill,
    write_finding_blocks,
)
from qc_tool.io.formula_cache import FormulaExtractionCache
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
from qc_tool.review import requeue_identity_key
from qc_tool.run_action import (
    MAX_RUN_ACTION_ITEMS,
    RankedTableEvidence,
    RunActionItem,
    RunActionReason,
    RunActionRequired,
    RunBlockedError,
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
    formula_cache: FormulaExtractionCache | None = None,
    formula_engine: Literal["native", "excel", "libreoffice", "auto"] = "auto",
    _native_compat_mode: bool = False,
    _xlsb_values_engine: Literal["pyxlsb", "native", "auto"] = "auto",
) -> WorkbookSnapshot:
    check_cancelled(cancellation_token)
    report_progress(on_progress, phase, total=1, detail=path.name)
    workbook = load_workbook_snapshot(
        path,
        password=password,
        allow_large_workbook=allow_large_workbooks,
        cancellation_token=cancellation_token,
        formula_cache=formula_cache,
        formula_engine=formula_engine,
        _native_formula_compat_mode=_native_compat_mode,
        _xlsb_values_engine=_xlsb_values_engine,
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


def _population_coverage(
    finding_class: FindingClass, stats: ClassPopulationStats
) -> CoverageItem:
    """Per-class group-first disclosure: candidates, populations, replays."""
    detail = (
        f"{stats.candidates:,} cells summarised as {stats.populations:,} "
        f"population(s)"
    )
    if stats.replayed:
        detail += (
            f"; {stats.replayed:,} replayed as atomic findings "
            f"({stats.below_threshold:,} below threshold, "
            f"{stats.over_cap:,} over the geometry cap, "
            f"{stats.heterogeneous_evidence:,} with mixed evidence)"
        )
    return CoverageItem(
        check_id=f"excel-population-{finding_class.value}",
        label=f"Group-first populations ({finding_class.value})",
        artifact="excel",
        state=CoverageState.CHECKED,
        findings=stats.candidates,
        detail=detail,
    )


def _enrich_population_samples(
    populations: list[Finding],
    *,
    current_workbook: WorkbookSnapshot | None,
    dependency_graph: DependencyGraph | None,
) -> None:
    """Compute impacts for each population's <= 5 samples only, labelled with
    their sample coordinate -- never the population's full downstream set.

    Excerpts are deliberately not attached here (on-demand "Expand members"
    is a later step); only impacts, which do not need cell records to be
    resident, are computed.
    """
    for population in populations:
        evidence = population.population
        if evidence is None or not evidence.samples:
            continue
        sample_findings = [
            Finding(
                artifact=population.artifact,
                artifact_member=population.artifact_member,
                finding_class=population.finding_class,
                sheet=population.sheet,
                location=sample.current_location,
                message="",
            )
            for sample in evidence.samples
        ]
        if dependency_graph is not None:
            accumulator: ImpactAccumulator | None = ImpactAccumulator(dependency_graph)
            accumulator.annotate(sample_findings)
        else:
            accumulator = None
        if current_workbook is not None:
            annotate_chart_impacts(sample_findings, current_workbook, dependency_graph)
        if accumulator is not None:
            accumulator.finalize(sample_findings)
        else:
            limit_impacts(sample_findings)
        impacts = [
            f"{sample.current_location}: {impact}"
            for sample, finding in zip(evidence.samples, sample_findings, strict=True)
            for impact in finding.impacts
        ]
        if impacts:
            # `Finding.impacts` stays empty for population findings -- only
            # the typed, explicitly-sampled field carries this evidence, so
            # no renderer/story-evidence/priority signal can mistake a
            # bounded 5-member sample for the population's full downstream
            # set (Criterion 6).
            population.population = evidence.model_copy(
                update={"sampled_impacts": tuple(impacts)}
            )


def _finding_batches(
    findings: Sequence[Finding], *, batch_size: int = BLOCK_FINDINGS
) -> Iterator[list[Finding]]:
    """Bounded batches from list or lazy spill-backed finding sequences."""
    source = (
        findings.iter_trusted()
        if isinstance(findings, FindingSequence)
        else iter(findings)
    )
    batch: list[Finding] = []
    for finding in source:
        batch.append(finding)
        if len(batch) >= batch_size:
            yield batch
            batch = []
    if batch:
        yield batch


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


def _drain_cycle_batch(
    batch: list[Finding],
    *,
    stream: _FindingStream,
    profile: DeliverableProfile,
    today: dt.date,
    baseline_workbook: WorkbookSnapshot | None,
    current_workbook: WorkbookSnapshot | None,
    current_deck: DeckSnapshot | None,
    dependency_graph: DependencyGraph | None,
) -> None:
    """Enrich and spill a non-value batch without retaining its full payload.

    Formula findings were historically small, but partial XLSB text can make
    them as large as the workbook's formula population. Clear each processed
    slice in place so aliases such as ``formula_findings`` release their
    Finding objects as soon as the bounded chunk reaches the spill.
    """
    for start in range(0, len(batch), BLOCK_FINDINGS):
        end = min(start + BLOCK_FINDINGS, len(batch))
        chunk = batch[start:end]
        assign_severities(chunk, profile, today=today)
        _enrich_retained_findings(
            chunk,
            baseline_workbook=baseline_workbook,
            current_workbook=current_workbook,
            current_deck=current_deck,
            dependency_graph=dependency_graph,
            crosscheck=profile.crosscheck,
        )
        stream.add(chunk)
        for index in range(start, end):
            batch[index] = None  # type: ignore[assignment]
    batch.clear()


def _buffer_cycle_batches(
    batches: list[list[Finding]],
    *,
    profile: DeliverableProfile,
    today: dt.date,
    baseline_workbook: WorkbookSnapshot | None,
    current_workbook: WorkbookSnapshot | None,
    current_deck: DeckSnapshot | None,
    dependency_graph: DependencyGraph | None,
) -> tuple[FindingSequence | None, Path | None]:
    """Enrich post-value findings now, replay them later in original order.

    Clears every input batch in place after its payload reaches the compact
    buffer; callers must not inspect ``batches`` afterward. This is the memory
    contract that releases formula Finding objects before value production.
    """
    if not any(batches):
        return None, None
    directory = Path(tempfile.mkdtemp(prefix="qc-post-findings-"))

    def payloads() -> Iterator[object]:
        for batch in batches:
            for start in range(0, len(batch), BLOCK_FINDINGS):
                end = min(start + BLOCK_FINDINGS, len(batch))
                chunk = batch[start:end]
                assign_severities(chunk, profile, today=today)
                _enrich_retained_findings(
                    chunk,
                    baseline_workbook=baseline_workbook,
                    current_workbook=current_workbook,
                    current_deck=current_deck,
                    dependency_graph=dependency_graph,
                    crosscheck=profile.crosscheck,
                )
                for finding in chunk:
                    yield finding_payload(finding)
                for index in range(start, end):
                    batch[index] = None  # type: ignore[assignment]
            batch.clear()

    try:
        container = write_finding_blocks(directory / "post.qcfb", payloads())
    except BaseException:
        shutil.rmtree(directory, ignore_errors=True)
        raise
    return FindingSequence(container), directory


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
    candidate_sink: CandidateSpill | None = None,
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
        for region in regions:
            check_cancelled(cancellation_token)
            ignore, refresh, value_only_ignore = region_range_sets(sheet_profile, region)
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
                value_only_ignore=value_only_ignore,
                candidate_sink=candidate_sink,
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
    #: Run-level finding-output contract (plan-20260910): which population
    #: policy this run actually used, and why. ``PROFILE``/``None`` for any
    #: result built outside the versioned output-mode contract (e.g. a
    #: directly-constructed test fixture).
    requested_output_mode: FindingOutputMode = FindingOutputMode.PROFILE
    resolved_output_policy: ResolvedOutputPolicy | None = None
    files: dict[str, str] = field(default_factory=dict)  # role -> file name
    #: Resolved formula-engine/adapter-fingerprint string (e.g.
    #: "native-biff12:1.2.3") per excel role, when formula enrichment ran.
    #: Persisted so Re-QC/carry-forward can disclose a cross-run engine
    #: change instead of silently comparing evidence from two engines.
    formula_engines: dict[str, str] = field(default_factory=dict)
    #: Resolved cached-values decoder per excel role, including any fallback.
    values_engines: dict[str, str] = field(default_factory=dict)
    findings: Sequence[Finding] = field(default_factory=list)
    disclosures: list[str] = field(default_factory=list)
    coverage: list[CoverageItem] = field(default_factory=list)
    mapping_coverage: MappingCoverage | None = None
    mapping_suggestions: list[MappingSuggestion] = field(default_factory=list)
    verified_crosschecks: int = 0
    comparison_scope: ComparisonScope = field(default_factory=ComparisonScope)
    alignment_trust: AlignmentTrustPayload | None = None
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
) -> tuple[list[Finding], list[CoverageItem], AlignmentTrustPayload | None]:
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
        if isinstance(trust, AlignmentTrustManifestV2):
            trust = AlignmentTrustManifestV2(
                regions=tuple(
                    region.model_copy(update={"artifact_member": member_id})
                    for region in trust.regions
                ),
                unpaired=tuple(
                    region.model_copy(update={"artifact_member": member_id})
                    for region in trust.unpaired
                ),
            )
        else:
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
    manifests: list[AlignmentTrustPayload],
) -> AlignmentTrustPayload | None:
    if not manifests:
        return None
    if not any(isinstance(manifest, AlignmentTrustManifestV2) for manifest in manifests):
        # Every manifest is V1: merge unchanged, byte-identical to before
        # this feature existed.
        v1_manifests = [
            manifest
            for manifest in manifests
            if isinstance(manifest, AlignmentTrustManifest)
        ]
        return AlignmentTrustManifest(
            regions=tuple(
                region
                for manifest in v1_manifests
                for region in manifest.regions
            ),
            unpaired=tuple(
                region
                for manifest in v1_manifests
                for region in manifest.unpaired
            ),
        )
    # At least one member applied a confirmed identity rule: the merged
    # manifest is V2, promoting any plain V1 member's regions to V2 shape
    # with "unused" identity fields so every region lives in one place.
    return AlignmentTrustManifestV2(
        regions=tuple(
            region
            if isinstance(region, AlignmentRegionTrustV2)
            else promote_region_trust_to_v2(region)
            for manifest in manifests
            for region in manifest.regions
        ),
        unpaired=tuple(
            region
            for manifest in manifests
            for region in manifest.unpaired
        ),
    )


def _ranked_table_suggestions(
    alignment: WorkbookAlignment,
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
) -> list[RunActionItem]:
    """Bounded, value-free evidence for every unconfigured positional block
    region that looks like a ranked/sorted table under raw position.

    Only regions that fell back to plain positional row alignment are
    screened: a region with a confirmed ``RowIdentityRule`` already aligns by
    ``"keys"`` (see ``qc_tool.excel.align._align_rows_by_identity``), so it is
    never re-suggested once confirmed.
    """
    items: list[RunActionItem] = []
    for sheet_name, regions in alignment.regions.items():
        base_sheet = baseline.sheet(sheet_name)
        curr_sheet = current.sheet(sheet_name)
        for region in regions:
            if region.current.orientation != "block":
                continue
            if region.rows.method != "positional" or region.low_confidence:
                continue
            candidate = detect_ranked_table_candidate(
                base_sheet, curr_sheet, region.baseline, region.current
            )
            if candidate is None:
                continue
            anchor_cell = (
                f"{get_column_letter(region.current.min_col)}{region.current.min_row}"
            )
            available_columns = tuple(
                get_column_letter(column)
                for column in range(region.current.min_col, region.current.max_col + 1)
            )
            data_row_count = region.current.max_row - region.current.min_row + 1
            items.append(
                RunActionItem(
                    sheet=sheet_name,
                    cell=anchor_cell,
                    label=(
                        "Possible ranked/sorted table: columns "
                        + "+".join(candidate.column_letters)
                    ),
                    ranked_table_evidence=RankedTableEvidence(
                        sheet=sheet_name,
                        current_range=region.current.cell_range,
                        data_row_count=data_row_count,
                        available_columns=available_columns,
                        suggested_identity_columns=candidate.column_letters,
                        suggested_ordinal_columns=candidate.ordinal_column_letters,
                        non_blank_coverage=candidate.non_blank_coverage,
                        unique_ratio=candidate.unique_ratio,
                        key_overlap=candidate.key_overlap,
                        formula_ratio=candidate.formula_ratio,
                        displaced_ratio=candidate.displaced_ratio,
                        mismatch_reduction=candidate.mismatch_reduction,
                        projected_positional_mismatches=(
                            candidate.projected_positional_mismatches
                        ),
                        projected_avoided_mismatches=(
                            candidate.projected_avoided_mismatches
                        ),
                    ),
                )
            )
    return items


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


def _memberize_run_action(
    action: RunActionRequired,
    member_id: str,
) -> RunActionRequired:
    items = [
        item.model_copy(
            update={
                "member_id": member_id,
                "ranked_table_evidence": (
                    item.ranked_table_evidence.model_copy(
                        update={"member_id": member_id}
                    )
                    if item.ranked_table_evidence is not None
                    else None
                ),
            }
        )
        for item in action.items
    ]
    return RunActionRequired(
        version=action.version,
        reason=action.reason,
        items=items,
        message=action.message,
        omitted_items=action.omitted_items,
    )


def _aggregate_ranked_package_actions(
    actions: list[tuple[str, RunActionRequired]],
) -> RunActionRequired:
    member_actions = [
        _memberize_run_action(action, member_id)
        for member_id, action in actions
    ]
    selected: list[RunActionItem] = []
    max_items = max((len(action.items) for action in member_actions), default=0)
    for item_index in range(max_items):
        for action in member_actions:
            if item_index < len(action.items):
                selected.append(action.items[item_index])
                if len(selected) == MAX_RUN_ACTION_ITEMS:
                    break
        if len(selected) == MAX_RUN_ACTION_ITEMS:
            break
    known_items = sum(len(action.items) for action in member_actions)
    omitted_items = sum(action.omitted_items for action in member_actions)
    omitted_items += known_items - len(selected)
    member_count = len(member_actions)
    return RunActionRequired(
        version=2,
        reason=RunActionReason.ROW_IDENTITY_CONFIRMATION_REQUIRED,
        items=selected,
        omitted_items=omitted_items,
        message=(
            f"Review row matching for {member_count} package workbook "
            f"member{'s' if member_count != 1 else ''}, save the rules, and "
            "Re-QC."
        ),
    )


def _run_multi_package(
    *,
    manifest: PackageManifest,
    files: dict[str, Path],
    profile: DeliverableProfile,
    passwords: dict[str, str],
    mode: QCRunMode,
    output_mode: FindingOutputMode,
    allow_large_workbooks: bool,
    allow_dependency_indexing: bool = False,
    run_acceptance: NumericTolerance | None,
    compare_sheets: list[str] | None,
    compare_member_sheets: dict[str, tuple[str, ...]],
    compare_slides: list[int] | None,
    cancellation_token: CancellationToken | None,
    on_progress: ProgressCallback | None,
    formula_cache: FormulaExtractionCache | None = None,
    formula_telemetry: FormulaComparisonTelemetry | None = None,
    pair_key_telemetry: PairKeyTelemetry | None = None,
    population_telemetry: PopulationTelemetry | None = None,
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
    trust_manifests: list[AlignmentTrustPayload] = []
    formula_engines: dict[str, str] = {}
    values_engines: dict[str, str] = {}
    current_deck_snapshot: DeckSnapshot | None = None
    package_reconciler: MultiPackageReconciler | None = None
    package_result = None
    member_events = 0
    ranked_blocks: list[tuple[str, RunActionRequired]] = []
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
            output_mode=output_mode,
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
            try:
                subresult = run_qc(
                    baseline_excel=files[baseline.role_key],
                    current_excel=files[current.role_key],
                    profile=projected_profile(member_id),
                    passwords=member_passwords(
                        ("baseline_excel", baseline.role_key),
                        ("current_excel", current.role_key),
                    ),
                    mode=QCRunMode.CYCLE_COMPARISON,
                    output_mode=output_mode,
                    allow_large_workbooks=allow_large_workbooks,
                    allow_dependency_indexing=allow_dependency_indexing,
                    run_acceptance=run_acceptance,
                    compare_sheets=member_scope,
                    cancellation_token=cancellation_token,
                    on_progress=on_progress,
                    formula_cache=formula_cache,
                    _formula_telemetry=formula_telemetry,
                    _pair_key_telemetry=pair_key_telemetry,
                    _population_telemetry=population_telemetry,
                )
            except RunBlockedError as blocked:
                if (
                    blocked.action_required.reason
                    is RunActionReason.ROW_IDENTITY_CONFIRMATION_REQUIRED
                ):
                    ranked_blocks.append((member_id, blocked.action_required))
                    continue
                blocked.action_required = _memberize_run_action(
                    blocked.action_required,
                    member_id,
                )
                raise
            member_findings, member_coverage, member_trust = (
                _memberize_excel_result(subresult, member_id)
            )
            add_findings(member_findings)
            coverage.extend(member_coverage)
            if member_trust is not None:
                trust_manifests.append(member_trust)
            if engine := subresult.formula_engines.get("baseline_excel"):
                formula_engines[baseline.role_key] = engine
            if engine := subresult.formula_engines.get("current_excel"):
                formula_engines[current.role_key] = engine
            if engine := subresult.values_engines.get("baseline_excel"):
                values_engines[baseline.role_key] = engine
            if engine := subresult.values_engines.get("current_excel"):
                values_engines[current.role_key] = engine
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
                output_mode=output_mode,
                allow_large_workbooks=allow_large_workbooks,
                compare_sheets=member_scope,
                cancellation_token=cancellation_token,
                on_progress=on_progress,
                _snapshot_capture=capture,
                formula_cache=formula_cache,
                _formula_telemetry=formula_telemetry,
                _pair_key_telemetry=pair_key_telemetry,
                _population_telemetry=population_telemetry,
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
            if engine := subresult.formula_engines.get("current_excel"):
                formula_engines[current.role_key] = engine
            if engine := subresult.values_engines.get("current_excel"):
                values_engines[current.role_key] = engine
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

    if ranked_blocks:
        raise RunBlockedError(_aggregate_ranked_package_actions(ranked_blocks))

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
            output_mode=output_mode,
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
            output_mode=output_mode,
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
                state=CoverageState.NOT_INCLUDED,
                detail="Use final-package mode for Excel-to-PowerPoint QC",
            )
        )

    report_progress(on_progress, RunPhase.FINALIZING_FINDINGS, total=1)
    add_findings(expired_waiver_findings(profile, today))
    sequence, severity_counts = stream.finalize()
    report_progress(
        on_progress,
        RunPhase.FINALIZING_FINDINGS,
        processed=1,
        total=1,
    )
    result = QCRunResult(
        profile_name=profile.name,
        mode=mode,
        requested_output_mode=output_mode,
        resolved_output_policy=resolve_output_policy(output_mode, profile.review_policy),
        files={member.role_key: member.display_name for member in manifest.members},
        formula_engines=formula_engines,
        values_engines=values_engines,
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
    output_mode: FindingOutputMode = FindingOutputMode.PROFILE,
    allow_large_workbooks: bool = False,
    allow_dependency_indexing: bool = False,
    run_acceptance: NumericTolerance | None = None,
    compare_sheets: list[str] | None = None,
    compare_slides: list[int] | None = None,
    cancellation_token: CancellationToken | None = None,
    on_progress: ProgressCallback | None = None,
    _snapshot_capture: _SnapshotCapture | None = None,
    formula_cache: FormulaExtractionCache | None = None,
    _native_compat_mode: bool = False,
    _xlsb_values_engine: Literal["pyxlsb", "native", "auto"] = "auto",
    _formula_telemetry: FormulaComparisonTelemetry | None = None,
    _pair_key_telemetry: PairKeyTelemetry | None = None,
    _population_telemetry: PopulationTelemetry | None = None,
) -> QCRunResult:
    """Run a full QC comparison. ``passwords`` is keyed by file name.

    ``run_acceptance`` is the analyst's run-level threshold: numeric value
    differences within either bound stay visible as within-tolerance Info
    findings (never suppressed). ``compare_sheets`` / ``compare_slides``
    narrow which sheets (by name) and slides (1-based index) may produce
    findings; files still load fully so cross-references keep resolving.
    ``allow_dependency_indexing`` forces full dependency-graph indexing
    (circular detection, formula/chart/PPT-chart impacts) above the
    documented size policy that otherwise skips it with a disclosed
    coverage reason; distinct from ``allow_large_workbooks``.
    ``_native_compat_mode`` is a private, oracle-verification-only switch
    (plan Criterion 13(a)): when the resolved XLSB formula engine is
    ``native``, restricts its returned text to exactly the coordinates a
    legacy engine also covers, for comparing against a legacy-engine oracle.
    Never set by production callers; ignored for non-native engines.
    ``_xlsb_values_engine`` selects the values-decoding engine for XLSB
    workbooks (plan-20260908-phase-b-guest-performance-followup.md /
    plan-20260909 Step 11): ``"auto"`` (the production default) prefers the
    native kernel when the optional extension is importable and falls back
    to ``pyxlsb`` with a disclosed, content-free reason on any native
    runtime failure; ``"native"``/``"pyxlsb"`` force one engine explicitly
    and are diagnostic-only (a forced ``"native"`` raises instead of falling
    back). ``_formula_telemetry`` is a private, diagnostic-only hook that
    accumulates ``FormulaComparisonTelemetry``
    counters -- including the ``RunPhase.COMPARING_FORMULAS``-scoped
    ``assess_workbook_complexity()`` cost -- for the same follow-up plan;
    never set by production callers. ``output_mode`` (plan-20260910) is the
    run-level finding-output contract: ``profile`` (default) resolves
    population output exactly as the profile's own ``review_policy``
    persists it; ``decision`` forces population output on (the profile's
    own explicit policy when it already enables one, else a versioned
    conservative built-in policy); ``atomic`` forces population output off
    regardless of profile. The resolved policy is recorded on the returned
    result as ``resolved_output_policy``; ``profile_sha256`` is unaffected.
    """
    check_cancelled(cancellation_token)
    mode = QCRunMode(mode)
    output_mode = FindingOutputMode(output_mode)
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
                output_mode=output_mode,
                allow_large_workbooks=allow_large_workbooks,
                allow_dependency_indexing=allow_dependency_indexing,
                run_acceptance=run_acceptance,
                compare_sheets=compare_sheets,
                compare_member_sheets=compare_member_sheets or {},
                compare_slides=compare_slides,
                cancellation_token=cancellation_token,
                on_progress=on_progress,
                formula_cache=formula_cache,
                formula_telemetry=_formula_telemetry,
                pair_key_telemetry=_pair_key_telemetry,
                population_telemetry=_population_telemetry,
            )
    if mode is QCRunMode.CURRENT_FILE_PREFLIGHT:
        if baseline_excel is not None or baseline_ppt is not None:
            raise ValueError("current-file preflight does not accept baseline files")
        if current_excel is None and current_ppt is None:
            raise ValueError("current-file preflight needs a current Excel or PowerPoint file")
        profile = profile or default_profile()
        passwords = passwords or {}
        result = QCRunResult(profile_name=profile.name, mode=mode)
        result.requested_output_mode = output_mode
        result.resolved_output_policy = resolve_output_policy(
            output_mode, profile.review_policy
        )
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
                formula_cache=formula_cache,
                formula_engine=profile.excel.formula_engine,
                _native_compat_mode=_native_compat_mode,
                _xlsb_values_engine=_xlsb_values_engine,
            )
            result.files["current_excel"] = current_excel.name
            if workbook.formula_source is not None:
                result.formula_engines["current_excel"] = workbook.formula_source
            if workbook.values_source is not None:
                result.values_engines["current_excel"] = workbook.values_source
            if workbook.values_engine_fallback_detail:
                result.disclosures.append(
                    f"current_excel: {workbook.values_engine_fallback_detail}"
                )
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
                    state=CoverageState.NOT_INCLUDED,
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
                        state=CoverageState.NOT_INCLUDED,
                        detail="No current deck supplied",
                    ),
                    media_structural_coverage(),
                    CoverageItem(
                        check_id="ppt-media-visual",
                        label="Rendered media and visual layout",
                        artifact="ppt",
                        state=CoverageState.NOT_INCLUDED,
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
                state=CoverageState.NOT_INCLUDED,
                detail="No baseline workbook supplied",
            ),
            CoverageItem(
                check_id="ppt-cycle-comparison",
                label="Historical PowerPoint changes",
                artifact="ppt",
                state=CoverageState.NOT_INCLUDED,
                detail="No baseline deck supplied",
            ),
            CoverageItem(
                check_id="excel-ppt-crosscheck",
                label="Excel to PowerPoint mappings",
                artifact="package",
                state=CoverageState.NOT_INCLUDED,
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
            formula_cache=formula_cache,
            formula_engine=profile.excel.formula_engine,
            _native_compat_mode=_native_compat_mode,
            _xlsb_values_engine=_xlsb_values_engine,
        )
        deck = _load_powerpoint_file(
            current_ppt,
            password=_password_for(passwords, "current_ppt", current_ppt),
            phase=RunPhase.LOADING_CURRENT_POWERPOINT,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
        )
        result = QCRunResult(profile_name=profile.name, mode=mode)
        result.requested_output_mode = output_mode
        result.resolved_output_policy = resolve_output_policy(
            output_mode, profile.review_policy
        )
        result.package_manifest = manifest
        if acceptance_requested:
            logger.warning("acceptance threshold ignored in %s", mode.value)
            result.disclosures.append(acceptance_not_applied)
        result.files = {
            "current_excel": current_excel.name,
            "current_ppt": current_ppt.name,
        }
        if workbook.formula_source is not None:
            result.formula_engines["current_excel"] = workbook.formula_source
        if workbook.values_source is not None:
            result.values_engines["current_excel"] = workbook.values_source
        if workbook.values_engine_fallback_detail:
            result.disclosures.append(
                f"current_excel: {workbook.values_engine_fallback_detail}"
            )
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
                state=CoverageState.NOT_INCLUDED,
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
    resolved_output_policy = resolve_output_policy(output_mode, profile.review_policy)
    result.requested_output_mode = output_mode
    result.resolved_output_policy = resolved_output_policy
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
    population_policy = resolved_output_policy.populations
    candidate_sink = (
        CandidateSpill(profile, today, telemetry=_population_telemetry)
        if population_policy.enabled
        else None
    )
    # One bounded cache per run_qc() call, shared across every sheet/region so
    # a formula pattern repeated across the workbook is classified once --
    # see FormulaPairAnalysisMemo's own docstring for the safety evidence
    # (plan-20260910, Step 5).
    pair_analysis_memo = FormulaPairAnalysisMemo()
    pre_findings: list[Finding] = []
    post_batches: list[list[Finding]] = []
    current_workbook = None
    baseline_workbook: WorkbookSnapshot | None = None
    dependency_graph: DependencyGraph | None = None
    dependency_skip_reason: str | None = None
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
            formula_cache=formula_cache,
            formula_engine=profile.excel.formula_engine,
            _native_compat_mode=_native_compat_mode,
            _xlsb_values_engine=_xlsb_values_engine,
        )
        curr_wb = _load_excel_file(
            current_excel,
            password=_password_for(passwords, "current_excel", current_excel),
            phase=RunPhase.LOADING_CURRENT_EXCEL,
            allow_large_workbooks=allow_large_workbooks,
            cancellation_token=cancellation_token,
            on_progress=on_progress,
            formula_cache=formula_cache,
            formula_engine=profile.excel.formula_engine,
            _native_compat_mode=_native_compat_mode,
            _xlsb_values_engine=_xlsb_values_engine,
        )
        current_workbook = curr_wb
        loaded_workbooks.extend((base_wb, curr_wb))
        if mode is QCRunMode.CYCLE_COMPARISON and profile is not None:
            mismatches = check_comparison_prerequisites(
                base_wb,
                curr_wb,
                profile.excel.comparison_prerequisites,
            )
            if mismatches:
                raise RunBlockedError(
                    RunActionRequired(
                        reason=RunActionReason.COMPARISON_PREREQUISITE_MISMATCH,
                        items=mismatches,
                        message=(
                            "Select the same scenario, fully recalculate, save, "
                            "and Re-QC before comparing."
                        ),
                    )
                )
        result.files["baseline_excel"] = baseline_excel.name
        result.files["current_excel"] = current_excel.name
        if base_wb.formula_source is not None:
            result.formula_engines["baseline_excel"] = base_wb.formula_source
        if curr_wb.formula_source is not None:
            result.formula_engines["current_excel"] = curr_wb.formula_source
        if base_wb.values_source is not None:
            result.values_engines["baseline_excel"] = base_wb.values_source
        if curr_wb.values_source is not None:
            result.values_engines["current_excel"] = curr_wb.values_source
        if base_wb.values_engine_fallback_detail:
            result.disclosures.append(
                f"baseline_excel: {base_wb.values_engine_fallback_detail}"
            )
        if curr_wb.values_engine_fallback_detail:
            result.disclosures.append(
                f"current_excel: {curr_wb.values_engine_fallback_detail}"
            )
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
        result.coverage.append(external_link_reachability_coverage(base_wb, curr_wb))
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
        period_suggestions = []
        for sheet_name, regions in alignment.regions.items():
            period_suggestions.extend(
                internal_period_band_suggestions(
                    curr_wb.sheet(sheet_name),
                    [region.current for region in regions],
                )
            )
        suggestion_detail = "; ".join(
            f"{item.sheet}!{item.region_range}: pin a {item.axis} period region "
            f"at {'row' if item.axis == 'columns' else 'column'} {item.anchor} "
            f"({item.period_count} periods)"
            for item in period_suggestions[:8]
        )
        if len(period_suggestions) > 8:
            suggestion_detail += (
                f"; and {len(period_suggestions) - 8} more internal period bands"
            )
        result.coverage.append(
            CoverageItem(
                check_id="excel-period-axis-suggestions",
                label="Internal period-axis profile suggestions",
                artifact="excel",
                state=(
                    CoverageState.DEGRADED
                    if period_suggestions
                    else CoverageState.CHECKED
                ),
                detail=suggestion_detail,
            )
        )
        check_cancelled(cancellation_token)
        if mode is QCRunMode.CYCLE_COMPARISON:
            ranked_items = _ranked_table_suggestions(alignment, base_wb, curr_wb)
            if ranked_items:
                raise RunBlockedError(
                    RunActionRequired(
                        version=2,
                        reason=RunActionReason.ROW_IDENTITY_CONFIRMATION_REQUIRED,
                        items=ranked_items,
                        message=(
                            "One or more sheets look like a ranked or sorted "
                            "table compared by raw position. Confirm a row "
                            "identity (one or more columns) for each in the "
                            "profile, or leave it unconfigured to keep "
                            "comparing positionally, then re-run QC."
                        ),
                    )
                )
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
            candidate_sink=candidate_sink,
            telemetry=_formula_telemetry,
            pair_key_telemetry=_pair_key_telemetry,
            pair_analysis_memo=pair_analysis_memo,
        )
        check_cancelled(cancellation_token)
        post_batches.append(formula_findings)
        # Measured while COMPARING_FORMULAS is still open (not
        # INDEXING_DEPENDENCIES): on a large real workbook this cost-driver
        # scan over every formula cell can itself run for minutes, dwarfing
        # what remains of indexing once the size policy below skips the
        # actual dependency-graph build -- keeping it here is what lets
        # INDEXING_DEPENDENCIES measure only the work the policy can skip.
        complexity_scan_start = time.perf_counter()
        complexity = (
            assess_workbook_complexity(
                curr_wb,
                allow_complex_workbook=allow_large_workbooks,
                cancellation_token=cancellation_token,
            )
            if curr_wb.formulas_available
            else None
        )
        if _formula_telemetry is not None:
            _formula_telemetry.complexity_assessment_seconds += (
                time.perf_counter() - complexity_scan_start
            )
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
        if complexity is not None:
            result.coverage.append(_complexity_coverage(complexity))
            report_progress(on_progress, RunPhase.INDEXING_DEPENDENCIES, total=1)
            dependency_skip_reason = dependency_index_skip_reason(
                complexity,
                allow_dependency_indexing=allow_dependency_indexing,
            )
            if dependency_skip_reason is None:
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
            else:
                dependency_state = CoverageState.DEGRADED
                dependency_detail = (
                    "Dependency indexing, formula impact tracing, and "
                    f"chart-impact tracing skipped by size policy: "
                    f"{dependency_skip_reason}"
                )
                circular_coverage = CoverageItem(
                    check_id="excel-circular-references",
                    label="Circular formula references",
                    artifact="excel",
                    state=CoverageState.DEGRADED,
                    detail=(
                        f"Circular-reference detection skipped by size policy: "
                        f"{dependency_skip_reason}"
                    ),
                )
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
                    state=CoverageState.NOT_INCLUDED,
                    detail="Excel pair not supplied",
                ),
                CoverageItem(
                    check_id="excel-values",
                    label="Excel values and presentation",
                    artifact="excel",
                    state=CoverageState.NOT_INCLUDED,
                    detail="Excel pair not supplied",
                ),
                CoverageItem(
                    check_id="excel-structure",
                    label="Excel workbook structure",
                    artifact="excel",
                    state=CoverageState.NOT_INCLUDED,
                    detail="Excel pair not supplied",
                ),
                CoverageItem(
                    check_id="excel-formulas",
                    label="Excel formulas",
                    artifact="excel",
                    state=CoverageState.NOT_INCLUDED,
                    detail="Excel pair not supplied",
                ),
                CoverageItem(
                    check_id="excel-dependencies",
                    label="Formula dependency impact tracing",
                    artifact="excel",
                    state=CoverageState.NOT_INCLUDED,
                    detail="Excel pair not supplied",
                ),
                CoverageItem(
                    check_id="excel-interaction-rules",
                    label="Data validation and conditional-format rules",
                    artifact="excel",
                    state=CoverageState.NOT_INCLUDED,
                    detail="Excel pair not supplied",
                ),
                CoverageItem(
                    check_id="excel-conditional-format-styles",
                    label="Conditional-format differential styles",
                    artifact="excel",
                    state=CoverageState.NOT_INCLUDED,
                    detail="Excel pair not supplied",
                ),
                CoverageItem(
                    check_id="excel-availability",
                    label="Availability boundaries",
                    artifact="excel",
                    state=CoverageState.NOT_INCLUDED,
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
                    state=CoverageState.NOT_INCLUDED,
                    detail="PowerPoint pair not supplied",
                ),
                CoverageItem(
                    check_id="ppt-availability",
                    label="Availability boundaries",
                    artifact="ppt",
                    state=CoverageState.NOT_INCLUDED,
                    detail="PowerPoint pair not supplied",
                ),
                media_structural_coverage(),
                CoverageItem(
                    check_id="ppt-media-visual",
                    label="Rendered media and visual layout",
                    artifact="ppt",
                    state=CoverageState.NOT_INCLUDED,
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
                    CoverageState.DEGRADED
                    if not current_deck.charts_available or dependency_skip_reason
                    else CoverageState.CHECKED
                ),
                findings=len(crosscheck.findings),
                detail=(
                    f"{len(crosscheck.verified)} mappings verified"
                    + (
                        ""
                        if current_deck.charts_available
                        else "; visible native chart labels unavailable"
                    )
                    + (
                        f"; PowerPoint chart-impact tracing skipped by size "
                        f"policy: {dependency_skip_reason}"
                        if dependency_skip_reason
                        else ""
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
                state=(
                    CoverageState.NOT_INCLUDED
                    if current_deck is None or current_workbook is None
                    else CoverageState.UNAVAILABLE
                ),
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
    pre_findings[:] = validated_scope.filter_findings(pre_findings)
    for batch in post_batches:
        batch[:] = validated_scope.filter_findings(batch)

    stream = _FindingStream()
    buffered_post: FindingSequence | None = None
    buffered_post_dir: Path | None = None
    try:
        # Enrich non-values before the part loop so excerpts still see every
        # sheet, but drain in bounded chunks: partial XLSB formula findings are
        # no longer assumed to be a small population.
        report_progress(on_progress, RunPhase.QUERYING_IMPACTS, total=1)
        _drain_cycle_batch(
            pre_findings,
            stream=stream,
            profile=profile,
            today=today,
            baseline_workbook=baseline_workbook,
            current_workbook=current_workbook,
            current_deck=current_deck,
            dependency_graph=dependency_graph,
        )
        buffered_post, buffered_post_dir = _buffer_cycle_batches(
            post_batches,
            profile=profile,
            today=today,
            baseline_workbook=baseline_workbook,
            current_workbook=current_workbook,
            current_deck=current_deck,
            dependency_graph=dependency_graph,
        )
        report_progress(
            on_progress, RunPhase.QUERYING_IMPACTS, processed=1, total=1
        )
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
                candidate_sink=candidate_sink,
            )
            if values_coverage is not None:
                values_coverage.findings = produced
        report_progress(on_progress, RunPhase.FINALIZING_FINDINGS, total=1)
        if candidate_sink is not None:
            outcome = finalize_populations(
                candidate_sink,
                population_policy,
                validated_scope,
                telemetry=_population_telemetry,
            )
            for finding_class, class_stats in outcome.stats.items():
                result.coverage.append(_population_coverage(finding_class, class_stats))
            if outcome.replay_findings:
                for replay_batch in _finding_batches(outcome.replay_findings):
                    assign_severities(replay_batch, profile, today=today)
                    _enrich_retained_findings(
                        replay_batch,
                        baseline_workbook=baseline_workbook,
                        current_workbook=current_workbook,
                        current_deck=current_deck,
                        dependency_graph=dependency_graph,
                        crosscheck=profile.crosscheck,
                    )
                    stream.add(replay_batch)
            if outcome.population_findings:
                for population_batch in _finding_batches(
                    outcome.population_findings
                ):
                    _enrich_population_samples(
                        population_batch,
                        current_workbook=current_workbook,
                        dependency_graph=dependency_graph,
                    )
                    stream.add(population_batch)
        if buffered_post is not None:
            stream.add(buffered_post.iter_trusted())
            if buffered_post_dir is not None:
                shutil.rmtree(buffered_post_dir, ignore_errors=True)
                buffered_post_dir = None
        expired = expired_waiver_findings(profile, today)
        assign_severities(expired, profile, today=today)
        stream.add(expired)
        report_progress(
            on_progress,
            RunPhase.FINALIZING_FINDINGS,
            processed=1,
            total=1,
        )
        report_progress(on_progress, RunPhase.BUILDING_REVIEW, total=1)
        sequence, severity_counts = stream.finalize()
    except BaseException:
        stream.abort()
        if candidate_sink is not None:
            candidate_sink.abort()
        if buffered_post_dir is not None:
            shutil.rmtree(buffered_post_dir, ignore_errors=True)
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


PopulationRepresentationKey = tuple[object, ...]


def _output_policy_representation_key(
    mode: FindingOutputMode,
    policy: ResolvedOutputPolicy | None,
) -> tuple[object, ...]:
    """Canonical effective representation contract for cross-run comparison."""
    if policy is None:
        # Historical PROFILE rows predate explicit output policies and used
        # populations-disabled behavior.
        return (mode.value, False) if mode is FindingOutputMode.PROFILE else (mode.value, None)
    populations = policy.populations
    if not populations.enabled:
        return (mode.value, False)
    return (
        mode.value,
        True,
        populations.threshold,
        tuple(sorted(item.value for item in populations.classes)),
        populations.max_rectangles,
        populations.max_explicit_pairs,
    )


def _population_scope_key(finding: Finding) -> tuple[str, ...]:
    return (
        finding.artifact,
        finding.artifact_member,
        finding.finding_class.value,
        finding.sheet or "",
    )


def _baseline_translation_mode_for_representation(finding: Finding) -> str:
    if finding.population is not None:
        membership = finding.population.membership
        if membership.baseline_mode == "shift":
            return "in-place" if membership.shift == (0, 0) else "translated"
        return "translated"
    if finding.baseline_location is None:
        return "no-baseline"
    try:
        baseline = coordinate_to_tuple(finding.baseline_location)
        current = coordinate_to_tuple(finding.location or "")
    except ValueError:
        return "marker"
    return "in-place" if baseline == current else "translated"


def _population_representation_key(
    finding: Finding,
) -> PopulationRepresentationKey | None:
    """Comparable group shape for an eligible population or atomic finding."""
    if finding.finding_class not in {
        FindingClass.FORMULA_LOGIC_CHANGED,
        FindingClass.NUMBER_FORMAT_CHANGED,
    }:
        return None
    if finding.population is not None:
        shape_before = finding.population.shape_before_digest
        shape_after = finding.population.shape_after_digest
    elif finding.finding_class is FindingClass.NUMBER_FORMAT_CHANGED:
        shape_before = finding.baseline_value or ""
        shape_after = finding.current_value or ""
    else:
        baseline_formula = finding.baseline_value
        current_formula = finding.current_value
        baseline_location = finding.baseline_location
        current_location = finding.location
        if (
            not isinstance(baseline_formula, str)
            or not isinstance(current_formula, str)
            or not isinstance(baseline_location, str)
            or not isinstance(current_location, str)
        ):
            return None
        try:
            baseline_row, baseline_column = coordinate_to_tuple(baseline_location)
            current_row, current_column = coordinate_to_tuple(current_location)
            baseline_r1c1 = to_r1c1(
                baseline_formula,
                baseline_row,
                baseline_column,
            )
            current_r1c1 = to_r1c1(
                current_formula,
                current_row,
                current_column,
            )
        except (TokenizerError, TypeError, ValueError):
            return None
        shape_before = hashlib.sha256(baseline_r1c1.encode("utf-8")).hexdigest()
        shape_after = hashlib.sha256(current_r1c1.encode("utf-8")).hexdigest()
    return (
        *_population_scope_key(finding),
        finding.severity.value if finding.severity is not None else "",
        finding.expected_reason.value if finding.expected_reason is not None else "",
        finding.provenance.value if finding.provenance is not None else "",
        finding.subtype.value if finding.subtype is not None else "",
        finding.materiality.value if finding.materiality is not None else "",
        finding.temporal_context.value if finding.temporal_context is not None else "",
        tuple(sorted(tag.value for tag in finding.evidence_tags)),
        finding.event_key,
        shape_before,
        shape_after,
        _baseline_translation_mode_for_representation(finding),
        finding.waiver_reason,
        finding.waiver_expires,
    )


def _population_scope_from_representation(
    key: PopulationRepresentationKey,
) -> tuple[str, ...]:
    return tuple(str(value) for value in key[:4])


def output_representations_compatible(
    previous_mode: FindingOutputMode,
    previous_policy: ResolvedOutputPolicy | None,
    previous: Sequence[Finding],
    current_mode: FindingOutputMode,
    current_policy: ResolvedOutputPolicy | None,
    current: Sequence[Finding],
) -> bool:
    """Whether a location/population identity comparison is demonstrably safe."""
    if _output_policy_representation_key(
        previous_mode, previous_policy
    ) != _output_policy_representation_key(current_mode, current_policy):
        return False

    def keys(
        findings: Sequence[Finding],
    ) -> tuple[
        set[PopulationRepresentationKey],
        set[PopulationRepresentationKey],
        set[tuple[str, ...]],
    ]:
        populations: set[PopulationRepresentationKey] = set()
        atomics: set[PopulationRepresentationKey] = set()
        unknown_atomic_scopes: set[tuple[str, ...]] = set()
        for finding in findings:
            key = _population_representation_key(finding)
            if finding.population is not None:
                if key is not None:
                    populations.add(key)
            elif finding.finding_class in {
                FindingClass.FORMULA_LOGIC_CHANGED,
                FindingClass.NUMBER_FORMAT_CHANGED,
            }:
                if key is None:
                    unknown_atomic_scopes.add(_population_scope_key(finding))
                else:
                    atomics.add(key)
        return populations, atomics, unknown_atomic_scopes

    previous_populations, previous_atomics, previous_unknown = keys(previous)
    current_populations, current_atomics, current_unknown = keys(current)
    if previous_populations & current_atomics or current_populations & previous_atomics:
        return False
    if any(
        _population_scope_from_representation(key) in current_unknown
        for key in previous_populations
    ):
        return False
    return not any(
        _population_scope_from_representation(key) in previous_unknown
        for key in current_populations
    )


def compare_findings(
    previous: Sequence[Finding], current: Sequence[Finding]
) -> FindingsDelta:
    """Match findings by identity (not by id) to compute a fix-progress delta.

    Uses `requeue_identity_key`, which pairs populations by their shape
    digest rather than location -- member-set churn (a row inserted or
    removed between runs) does not turn a persisting population into a
    spurious resolved+new pair (Criterion 5).

    Uses multiset (Counter) semantics, not set membership: two distinct
    findings that happen to share one identity key (e.g. two populations
    split only by waiver/severity, which `population_identity_digest`
    deliberately excludes) must not collapse into a single set entry --
    losing one of two same-identity findings between runs must still count
    as one resolved, not be hidden because the key was still present
    (Criterion 8).

    Caller contract: `previous` and `current` must pass
    `output_representations_compatible` first. Populations and atomics key
    into disjoint identity spaces here, and representation can change because
    of effective policy or threshold crossings even when the requested mode
    stays the same.
    """
    previous_counts = Counter(
        requeue_identity_key(f) for f in previous if f.severity is not Severity.EXPECTED
    )
    current_counts = Counter(
        requeue_identity_key(f) for f in current if f.severity is not Severity.EXPECTED
    )
    resolved = 0
    persisting = 0
    for key, previous_count in previous_counts.items():
        current_count = current_counts.get(key, 0)
        resolved += max(0, previous_count - current_count)
        persisting += min(previous_count, current_count)
    new = 0
    for key, current_count in current_counts.items():
        new += max(0, current_count - previous_counts.get(key, 0))
    return FindingsDelta(resolved=resolved, new=new, persisting=persisting)
