"""Run one QC comparison, write both reports, and record history.

This layer is deliberately free of UI and server imports so the CLI and an
owned worker process can execute a run without loading NiceGUI.
"""

import datetime as dt
import logging
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path

from qc_tool.config.profile import DeliverableProfile, NumericTolerance
from qc_tool.coverage import QCRunMode
from qc_tool.engine import FindingsDelta, QCRunResult, compare_findings, run_qc
from qc_tool.focus.model import FocusTargetSeed
from qc_tool.focus.targets import build_focus_targets
from qc_tool.history.store import RunHistory, sha256_file
from qc_tool.progress import (
    CancellationToken,
    ProgressCallback,
    RunPhase,
    check_cancelled,
    report_progress,
)
from qc_tool.report.excel_report import write_excel_report
from qc_tool.report.html_report import write_html_report
from qc_tool.security import private_directory

logger = logging.getLogger(__name__)


@dataclass(slots=True)
class RunArtifacts:
    run_id: int
    result: QCRunResult
    report_paths: dict[str, Path]
    rerun_of: int | None = None
    delta: FindingsDelta | None = None


def _focus_targets(
    result: QCRunResult, file_hashes: dict[str, str]
) -> dict[str, tuple[FocusTargetSeed, ...]]:
    """Private desktop-focus locators; a failure here never fails the QC run."""
    try:
        return build_focus_targets(
            result.findings, mode=result.mode, file_hashes=file_hashes
        )
    except Exception:
        # Fixed code only: never log locators, paths, values, or exception text.
        logger.warning("focus-target-generation-failed")
        return {}


def perform_run(
    work_dir: Path,
    files: dict[str, Path],
    passwords: dict[str, str],
    profile: DeliverableProfile,
    *,
    mode: QCRunMode = QCRunMode.CYCLE_COMPARISON,
    rerun_of: int | None = None,
    allow_large_workbooks: bool = False,
    acceptance_absolute: float = 0.0,
    acceptance_relative: float = 0.0,
    compare_sheets: list[str] | None = None,
    compare_slides: list[int] | None = None,
    cancellation_token: CancellationToken | None = None,
    on_progress: ProgressCallback | None = None,
) -> RunArtifacts:
    """Run QC, write both reports, and record the run in history."""
    if acceptance_absolute < 0 or acceptance_relative < 0:
        raise ValueError("acceptance thresholds cannot be negative")
    run_acceptance = (
        NumericTolerance(absolute=acceptance_absolute, relative=acceptance_relative)
        if acceptance_absolute > 0 or acceptance_relative > 0
        else None
    )
    password_by_file = {
        files[role].name: password
        for role, password in passwords.items()
        if role in files and password
    }
    result = run_qc(
        baseline_excel=files.get("baseline_excel"),
        current_excel=files.get("current_excel"),
        baseline_ppt=files.get("baseline_ppt"),
        current_ppt=files.get("current_ppt"),
        profile=profile,
        passwords=password_by_file,
        mode=mode,
        allow_large_workbooks=allow_large_workbooks,
        run_acceptance=run_acceptance,
        compare_sheets=compare_sheets,
        compare_slides=compare_slides,
        cancellation_token=cancellation_token,
        on_progress=on_progress,
    )
    check_cancelled(cancellation_token)
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

        report_progress(on_progress, RunPhase.RECORDING_HISTORY, total=1)
        history = RunHistory(work_dir / "history.sqlite3")
        delta: FindingsDelta | None = None
        if rerun_of is not None:
            try:
                previous = history.get_run(rerun_of)
                delta = compare_findings(previous.findings, result.findings)
            except KeyError:
                logger.warning("re-QC referenced missing run %s", rerun_of)
                rerun_of = None
        file_hashes: dict[str, str] = {}
        for role, path in files.items():
            check_cancelled(cancellation_token)
            file_hashes[role] = sha256_file(path)
        check_cancelled(cancellation_token)
        run_id = history.record_run(
            result,
            file_hashes=file_hashes,
            report_paths={kind: str(path) for kind, path in report_paths.items()},
            file_paths={role: str(path) for role, path in files.items()},
            rerun_of=rerun_of,
            focus_targets=_focus_targets(result, file_hashes),
        )
        recorded = True
        report_progress(
            on_progress,
            RunPhase.RECORDING_HISTORY,
            processed=1,
            total=1,
        )
        report_progress(on_progress, RunPhase.COMPLETE, processed=1, total=1)
    except BaseException:
        if not recorded:
            shutil.rmtree(run_dir, ignore_errors=True)
        raise
    return RunArtifacts(
        run_id=run_id,
        result=result,
        report_paths=report_paths,
        rerun_of=rerun_of,
        delta=delta,
    )
