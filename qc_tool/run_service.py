"""Run one QC comparison, write both reports, and record history.

This layer is deliberately free of UI and server imports so the CLI and an
owned worker process can execute a run without loading NiceGUI.
"""

import datetime as dt
import logging
import shutil
import tempfile
import time
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from qc_tool.config.profile import DeliverableProfile, NumericTolerance
from qc_tool.coverage import FindingOutputMode, QCRunMode
from qc_tool.engine import (
    FindingsDelta,
    QCRunResult,
    compare_findings,
    output_representations_compatible,
    run_qc,
)
from qc_tool.excel.formulas import FormulaComparisonTelemetry, PairKeyTelemetry
from qc_tool.excel.population import PopulationTelemetry
from qc_tool.history.store import RunHistory
from qc_tool.io.formula_cache import FormulaExtractionCache
from qc_tool.package import PackageManifest, paths_by_member
from qc_tool.progress import (
    CancellationToken,
    ProgressCallback,
    RunPhase,
    check_cancelled,
    report_progress,
)
from qc_tool.report.excel_report import write_excel_report
from qc_tool.report.html_report import write_html_report
from qc_tool.run_preflight import (
    hash_run_files,
    reject_duplicate_bytes,
    verify_run_file_hashes,
)
from qc_tool.security import private_directory

logger = logging.getLogger(__name__)

#: Reports are on-demand everywhere (user decision 2026-08-09): the queue
#: worker never writes them at run time — the run page generates and stores
#: them on request. Only explicit ``write_reports=True`` callers (the CLI,
#: whose report files ARE the output) write eagerly. This constant remains
#: the budget for run-completion extras such as the re-QC delta.
REPORT_DEFER_FINDINGS = 50_000


@dataclass(slots=True)
class RunArtifacts:
    run_id: int
    result: QCRunResult
    report_paths: dict[str, Path]
    rerun_of: int | None = None
    delta: FindingsDelta | None = None


@dataclass(slots=True)
class PerformRunTelemetry:
    """Aggregate-only wall-clock accounting for one ``perform_run()`` call
    (plan-20260910): initial hashing, the QC pipeline itself, source rehash,
    report writing, and history recording, plus a named residual covering
    everything else (e.g. package-manifest validation, deferred-report
    bookkeeping). Only elapsed seconds are ever recorded -- never a path,
    filename, or finding.
    """

    initial_hash_seconds: float = 0.0
    qc_seconds: float = 0.0
    rehash_seconds: float = 0.0
    reports_seconds: float = 0.0
    history_seconds: float = 0.0
    total_seconds: float = 0.0

    @property
    def residual_seconds(self) -> float:
        named = (
            self.initial_hash_seconds
            + self.qc_seconds
            + self.rehash_seconds
            + self.reports_seconds
            + self.history_seconds
        )
        return max(0.0, self.total_seconds - named)


def perform_run(
    work_dir: Path,
    files: dict[str, Path],
    passwords: dict[str, str],
    profile: DeliverableProfile,
    *,
    mode: QCRunMode = QCRunMode.CYCLE_COMPARISON,
    output_mode: FindingOutputMode = FindingOutputMode.PROFILE,
    rerun_of: int | None = None,
    allow_large_workbooks: bool = False,
    allow_dependency_indexing: bool = False,
    acceptance_absolute: float = 0.0,
    acceptance_relative: float = 0.0,
    compare_sheets: list[str] | None = None,
    compare_slides: list[int] | None = None,
    cancellation_token: CancellationToken | None = None,
    on_progress: ProgressCallback | None = None,
    package_manifest: PackageManifest | dict[str, object] | None = None,
    compare_member_sheets: dict[str, tuple[str, ...]] | None = None,
    write_reports: bool = False,
    on_subphase: Callable[[str, float], None] | None = None,
    formula_cache_enabled: bool = True,
    _perform_run_telemetry: PerformRunTelemetry | None = None,
    _formula_telemetry: FormulaComparisonTelemetry | None = None,
    _pair_key_telemetry: PairKeyTelemetry | None = None,
    _population_telemetry: PopulationTelemetry | None = None,
) -> RunArtifacts:
    """Run QC, record the run in history, and defer reports to on-demand.

    Reports are generated from the run page when actually needed;
    ``write_reports=True`` (the CLI) writes them eagerly at run time.
    ``on_subphase(name, elapsed_seconds)`` -- when given -- is a diagnostic
    hook passed straight through to ``RunHistory.record_run``; production
    callers never need it. ``formula_cache_enabled`` (default True) owns a
    private, bounded XLSB formula-extraction cache at
    ``work_dir/formula-cache``; disabling it changes no finding, only
    whether repeat external-engine extraction is skipped. ``output_mode``
    (plan-20260910) is the run-level finding-output contract: ``profile``
    (default, legacy-compatible) resolves population output exactly as the
    profile's own ``review_policy`` persists it; ``decision`` forces
    population output on; ``atomic`` forces it off. The resolved policy is
    recorded on the result and persisted with the run.
    ``_perform_run_telemetry`` is a private, diagnostic-only hook
    (plan-20260910) that accumulates ``PerformRunTelemetry`` timings around
    initial hashing, the QC pipeline, source rehash, report writing, and
    history recording; never set by production callers.
    ``_formula_telemetry`` is a private, diagnostic-only hook
    (plan-20260910, Step 7 precondition evidence) passed straight through to
    ``run_qc()``'s own ``_formula_telemetry`` param; never set by production
    callers.
    """
    _telemetry_start = (
        time.perf_counter() if _perform_run_telemetry is not None else 0.0
    )
    if acceptance_absolute < 0 or acceptance_relative < 0:
        raise ValueError("acceptance thresholds cannot be negative")
    run_acceptance = (
        NumericTolerance(absolute=acceptance_absolute, relative=acceptance_relative)
        if acceptance_absolute > 0 or acceptance_relative > 0
        else None
    )
    credentials = {
        role: password
        for role, password in passwords.items()
        if role in files and password
    }
    manifest_obj = (
        PackageManifest.from_role_files(files)
        if package_manifest is None
        else (
            package_manifest
            if isinstance(package_manifest, PackageManifest)
            else PackageManifest.model_validate(package_manifest)
        )
    )
    paths_by_member(files, manifest_obj)
    _hash_start = time.perf_counter() if _perform_run_telemetry is not None else 0.0
    file_hashes = hash_run_files(files, cancellation_token=cancellation_token)
    if _perform_run_telemetry is not None:
        _perform_run_telemetry.initial_hash_seconds += time.perf_counter() - _hash_start
    reject_duplicate_bytes(mode, files, file_hashes)

    formula_cache = (
        FormulaExtractionCache(work_dir / "formula-cache")
        if formula_cache_enabled
        else None
    )
    _qc_start = time.perf_counter() if _perform_run_telemetry is not None else 0.0
    result = run_qc(
        baseline_excel=files.get("baseline_excel"),
        current_excel=files.get("current_excel"),
        baseline_ppt=files.get("baseline_ppt"),
        current_ppt=files.get("current_ppt"),
        profile=profile,
        passwords=credentials,
        mode=mode,
        output_mode=output_mode,
        allow_large_workbooks=allow_large_workbooks,
        allow_dependency_indexing=allow_dependency_indexing,
        run_acceptance=run_acceptance,
        compare_sheets=compare_sheets,
        compare_slides=compare_slides,
        package_manifest=manifest_obj,
        package_files=files,
        compare_member_sheets={
            key: tuple(value)
            for key, value in (compare_member_sheets or {}).items()
        },
        cancellation_token=cancellation_token,
        on_progress=on_progress,
        formula_cache=formula_cache,
        _formula_telemetry=_formula_telemetry,
        _pair_key_telemetry=_pair_key_telemetry,
        _population_telemetry=_population_telemetry,
    )
    if _perform_run_telemetry is not None:
        _perform_run_telemetry.qc_seconds += time.perf_counter() - _qc_start
    check_cancelled(cancellation_token)
    _rehash_start = time.perf_counter() if _perform_run_telemetry is not None else 0.0
    verify_run_file_hashes(
        files,
        file_hashes,
        cancellation_token=cancellation_token,
    )
    if _perform_run_telemetry is not None:
        _perform_run_telemetry.rehash_seconds += time.perf_counter() - _rehash_start
    defer_reports = not write_reports
    report_paths: dict[str, Path] = {}
    run_dir: Path | None = None
    if not defer_reports:
        runs_dir = private_directory(work_dir / "runs")
        stamp = dt.datetime.now(dt.UTC).strftime("%Y%m%dT%H%M%S.%fZ-")
        run_dir = Path(tempfile.mkdtemp(prefix=stamp, dir=runs_dir))
        private_directory(run_dir)
        report_paths = {
            "excel": run_dir / "qc_report.xlsx",
            "html": run_dir / "qc_report.html",
        }
    recorded = False
    try:
        if defer_reports:
            # Reports generate on demand from the run page; the phase is
            # reported complete so progress consumers see every stage.
            report_progress(
                on_progress, RunPhase.WRITING_REPORTS, processed=2, total=2
            )
        else:
            _reports_start = (
                time.perf_counter() if _perform_run_telemetry is not None else 0.0
            )
            report_progress(on_progress, RunPhase.WRITING_REPORTS, total=2)
            write_excel_report(result, report_paths["excel"])
            check_cancelled(cancellation_token)
            report_progress(
                on_progress,
                RunPhase.WRITING_REPORTS,
                processed=1,
                total=2,
            )
            write_html_report(result, report_paths["html"])
            check_cancelled(cancellation_token)
            report_progress(
                on_progress,
                RunPhase.WRITING_REPORTS,
                processed=2,
                total=2,
            )
            if _perform_run_telemetry is not None:
                _perform_run_telemetry.reports_seconds += (
                    time.perf_counter() - _reports_start
                )

        report_progress(on_progress, RunPhase.RECORDING_HISTORY, total=1)
        history = RunHistory(work_dir / "history.sqlite3")
        delta: FindingsDelta | None = None
        if rerun_of is not None:
            try:
                previous = history.get_run(rerun_of)
                # The delta banner walks both runs finding by finding;
                # monster re-QCs skip it (the run page says so too). A naive
                # comparison also only makes sense within one output mode --
                # `requeue_identity_key` keys populations and atomics in
                # disjoint spaces, so a mode change between runs would make
                # every population look "resolved" and every atomic member it
                # covered look "new" (plan-20260910 Criterion 5).
                if (
                    len(previous.findings) <= REPORT_DEFER_FINDINGS
                    and len(result.findings) <= REPORT_DEFER_FINDINGS
                    and output_representations_compatible(
                        previous.requested_output_mode,
                        previous.resolved_output_policy,
                        previous.findings,
                        result.requested_output_mode,
                        result.resolved_output_policy,
                        result.findings,
                    )
                ):
                    delta = compare_findings(previous.findings, result.findings)
            except KeyError:
                logger.warning("re-QC referenced missing run %s", rerun_of)
                rerun_of = None
        _history_start = (
            time.perf_counter() if _perform_run_telemetry is not None else 0.0
        )
        run_id = history.record_run(
            result,
            file_hashes=file_hashes,
            report_paths={kind: str(path) for kind, path in report_paths.items()},
            file_paths={role: str(path) for role, path in files.items()},
            rerun_of=rerun_of,
            profile_snapshot=profile,
            on_subphase=on_subphase,
        )
        recorded = True
        if _perform_run_telemetry is not None:
            _perform_run_telemetry.history_seconds += (
                time.perf_counter() - _history_start
            )
        report_progress(
            on_progress,
            RunPhase.RECORDING_HISTORY,
            processed=1,
            total=1,
        )
        report_progress(on_progress, RunPhase.COMPLETE, processed=1, total=1)
    except BaseException:
        if not recorded and run_dir is not None:
            shutil.rmtree(run_dir, ignore_errors=True)
        raise
    if _perform_run_telemetry is not None:
        _perform_run_telemetry.total_seconds += time.perf_counter() - _telemetry_start
    return RunArtifacts(
        run_id=run_id,
        result=result,
        report_paths=report_paths,
        rerun_of=rerun_of,
        delta=delta,
    )
