"""Windows ext4-guest LARGE_WORKBOOK throughput acceptance: phase timings, CPU, and
combined process-tree peak memory for a real perform_run() invocation, with
a persistent formula-extraction cache to distinguish cold/one-hit/two-hit
scenarios (plan-20260904-large_workbook-load-and-formula-compare.md, Step 4).

Never prints or persists a filename, sheet name, formula, coordinate, or
defined name -- only counts, timings, hashes, aggregate memory figures, and
the caller-supplied ``--label``, which is restricted to a short, safe,
generic token (lowercase alphanumeric/hyphen only) so it structurally cannot
carry a filename or path. Run the safe pilot first, then the real LARGE_WORKBOOK pair,
on the same ext4-backed win11 QEMU guest, once for the pre-change code and
once for the candidate code (see the runbook in docs/HANDOFF.md for the
exact invocation sequence).

Usage (PowerShell or cmd, inside the project's Windows environment):

    python scripts\\windows_large_workbook_throughput_acceptance.py ^
        --baseline-excel Z:\\path\\to\\baseline.xlsb ^
        --current-excel Z:\\path\\to\\current.xlsb ^
        --work-dir C:\\QC-Pilot\\throughput-work ^
        --label pre-change-cold ^
        --output Z:\\QC_Tool\\windows-return\\pre-change-cold.json

Run it up to three times against the SAME --work-dir (so the formula cache
persists) to observe cold, then two-hit (same pair again -- both sides
cache-hit), scenarios. A one-hit (single new file, one cached side) scenario
needs a third distinct file and is not reproducible with only the two LARGE_WORKBOOK
files; if unavailable, its saving is inferred as approximately half of the
two-hit saving.

Scope: a single baseline/current Excel pair only. Multi-package (member
-qualified) scenarios are intentionally unsupported here -- not partially
threaded -- since this harness's whole purpose is one-pair throughput
acceptance; add a member-aware variant separately if that scope is ever
actually needed rather than half-wiring it into this script's flags.
"""

from __future__ import annotations

import argparse
import contextlib
import cProfile
import hashlib
import json
import logging
import multiprocessing as mp
import queue
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Protocol

import psutil

# Import qc_tool lazily after sys.path is set so this script can run from a
# plain `python scripts\...` invocation without an editable install.
_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_ROOT))

#: `--label` is reported verbatim in the output JSON, so it is restricted to
#: this safe, bounded pattern -- short, lowercase, hyphen-separated tokens
#: only. This cannot hold a filename, path, sheet name, or other free text
#: (which would contain characters like `/`, `\\`, `.`, or spaces), closing
#: the gap between this script's privacy claim and what it actually enforces.
_SAFE_LABEL_RE = re.compile(r"^[a-z0-9]+(-[a-z0-9]+){0,7}$")
_MAX_LABEL_LENGTH = 64


def _safe_label(value: str) -> str:
    if len(value) > _MAX_LABEL_LENGTH or not _SAFE_LABEL_RE.match(value):
        raise argparse.ArgumentTypeError(
            "label must be 1-64 lowercase alphanumeric/hyphen segments "
            "(e.g. 'large_workbook-pre-change-cold'), never a filename or path"
        )
    return value


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


_OFFICE_PROCESS_NAMES = frozenset(
    {"excel.exe", "soffice", "soffice.bin", "soffice.exe", "libreoffice"}
)


class _MonitoredProcess(Protocol):
    @property
    def pid(self) -> int | None: ...

    @property
    def exitcode(self) -> int | None: ...

    def is_alive(self) -> bool: ...

    def join(self, timeout: float | None = None) -> None: ...

    def terminate(self) -> None: ...

    def kill(self) -> None: ...


@dataclass(frozen=True, slots=True)
class ProcessSupervision:
    sampled_peak_rss_bytes: int
    owned_office_process_peak: int
    timed_out: bool
    exit_code: int | None


def _process_tree_metrics(process_id: int) -> tuple[int, int]:
    """Current RSS and owned Office-process count for one live process tree."""
    try:
        root = psutil.Process(process_id)
        processes = [root, *root.children(recursive=True)]
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return 0, 0
    total = 0
    office = 0
    for proc in processes:
        with contextlib.suppress(
            psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess
        ):
            total += proc.memory_info().rss
            if proc.name().casefold() in _OFFICE_PROCESS_NAMES:
                office += 1
    return total, office


def _system_office_process_count() -> int:
    count = 0
    for process in psutil.process_iter(("name",)):
        with contextlib.suppress(
            psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess
        ):
            if str(process.info.get("name") or "").casefold() in _OFFICE_PROCESS_NAMES:
                count += 1
    return count


def _terminate_process_tree(process: _MonitoredProcess) -> None:
    descendants: list[psutil.Process] = []
    if process.pid is not None:
        with contextlib.suppress(psutil.NoSuchProcess, psutil.AccessDenied):
            descendants = psutil.Process(process.pid).children(recursive=True)
    for descendant in descendants:
        with contextlib.suppress(psutil.NoSuchProcess, psutil.AccessDenied):
            descendant.terminate()
    if process.is_alive():
        process.terminate()
    process.join(timeout=2.0)
    alive = []
    if descendants:
        _gone, alive = psutil.wait_procs(descendants, timeout=2.0)
    for descendant in alive:
        with contextlib.suppress(psutil.NoSuchProcess, psutil.AccessDenied):
            descendant.kill()
    if process.is_alive():
        process.kill()
        process.join(timeout=2.0)


def _monitor_process(
    process: _MonitoredProcess,
    *,
    deadline_seconds: float,
    sample_interval_seconds: float = 0.05,
) -> ProcessSupervision:
    """Monitor an already-started child and enforce its wall-clock deadline."""
    if process.pid is None:
        raise ValueError("cannot monitor a process that has not started")
    deadline = time.monotonic() + deadline_seconds
    peak_rss = 0
    office_peak = 0
    timed_out = False
    while process.is_alive():
        rss, office = _process_tree_metrics(process.pid)
        peak_rss = max(peak_rss, rss)
        office_peak = max(office_peak, office)
        if time.monotonic() >= deadline:
            timed_out = True
            _terminate_process_tree(process)
            break
        process.join(timeout=sample_interval_seconds)
    rss, office = _process_tree_metrics(process.pid)
    peak_rss = max(peak_rss, rss)
    office_peak = max(office_peak, office)
    return ProcessSupervision(
        sampled_peak_rss_bytes=peak_rss,
        owned_office_process_peak=office_peak,
        timed_out=timed_out,
        exit_code=process.exitcode,
    )


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-excel", type=Path, required=True)
    parser.add_argument("--current-excel", type=Path, required=True)
    parser.add_argument(
        "--work-dir",
        type=Path,
        required=True,
        help="Persistent directory (reused across invocations) that owns "
        "the formula-extraction cache and history for this acceptance run.",
    )
    parser.add_argument(
        "--label",
        required=True,
        type=_safe_label,
        help="Short generic scenario token, e.g. 'large_workbook-pre-change-cold' "
        "(lowercase alphanumeric/hyphen only, max 64 chars -- never a "
        "filename or path).",
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--profile", type=Path, default=None, help="Optional named profile YAML."
    )
    parser.add_argument(
        "--output-mode",
        choices=["profile", "decision", "atomic"],
        default="decision",
        help="Run-level finding-output contract (plan-20260910); this "
        "harness's own purpose is measuring the compact decision-mode "
        "production lane, so 'decision' is the default -- pass 'profile' "
        "or 'atomic' explicitly only for a deliberate comparison.",
    )
    parser.add_argument(
        "--deadline-seconds",
        type=float,
        default=900.0,
        help="Hard wall-clock ceiling (plan-20260910 Criterion 14): the run "
        "is cancelled with a content-free TimeoutError past this many "
        "seconds, comfortably above the release gate so a genuinely "
        "runaway measurement cannot run unbounded.",
    )
    parser.add_argument(
        "--cprofile-output",
        type=Path,
        default=None,
        help="Optional path to dump cProfile stats (function name/file/line/"
        "timing only -- never argument values or finding content) for the "
        "perform_run() call, to locate a phase's real hotspot.",
    )
    return parser


def _execute_acceptance(args: argparse.Namespace) -> dict[str, object]:
    """Run one acceptance workload inside the supervised child process."""
    from qc_tool.config.profile import default_profile, load_profile
    from qc_tool.coverage import FindingOutputMode, QCRunMode
    from qc_tool.excel.formulas import FormulaComparisonTelemetry, PairKeyTelemetry
    from qc_tool.excel.population import PopulationTelemetry
    from qc_tool.history.review_state import finding_evidence_digest
    from qc_tool.history.store import RunHistory
    from qc_tool.progress import (
        CancellationToken,
        PhaseTelemetry,
        ProgressEvent,
        RunPhase,
    )
    from qc_tool.run_service import PerformRunTelemetry, perform_run

    files = {
        "baseline_excel": args.baseline_excel,
        "current_excel": args.current_excel,
    }
    profile = load_profile(args.profile) if args.profile is not None else default_profile()

    process = psutil.Process()
    phase_telemetry = PhaseTelemetry()
    formula_telemetry = FormulaComparisonTelemetry()
    pair_key_telemetry = PairKeyTelemetry()
    population_telemetry = PopulationTelemetry()
    perform_telemetry = PerformRunTelemetry()
    recording_history_subphases: dict[str, float] = {}
    cancel_flag = CancellationToken()

    profiler = cProfile.Profile() if args.cprofile_output is not None else None
    profiling_active = False

    def on_progress(event: ProgressEvent) -> None:
        nonlocal profiling_active
        if profiler is not None:
            if event.phase is RunPhase.RECORDING_HISTORY and not profiling_active:
                profiler.enable()
                profiling_active = True
            elif profiling_active and event.phase is not RunPhase.RECORDING_HISTORY:
                profiler.disable()
                profiling_active = False
        phase_telemetry(event)

    started = time.perf_counter()
    cpu_before = process.cpu_times()
    artifacts = perform_run(
        args.work_dir,
        files,
        {},
        profile,
        mode=QCRunMode.CYCLE_COMPARISON,
        output_mode=FindingOutputMode(args.output_mode),
        allow_large_workbooks=True,
        write_reports=False,
        cancellation_token=cancel_flag,
        on_progress=on_progress,
        on_subphase=recording_history_subphases.__setitem__,
        _perform_run_telemetry=perform_telemetry,
        _formula_telemetry=formula_telemetry,
        _pair_key_telemetry=pair_key_telemetry,
        _population_telemetry=population_telemetry,
    )
    if profiler is not None:
        if profiling_active:
            profiler.disable()
        profiler.dump_stats(str(args.cprofile_output))
    elapsed = time.perf_counter() - started
    cpu_after = process.cpu_times()

    result = artifacts.result
    digests = [finding_evidence_digest(f) for f in result.findings]
    ordered_digest = hashlib.sha256("\n".join(digests).encode("utf-8")).hexdigest()
    record = RunHistory(args.work_dir / "history.sqlite3").get_run(artifacts.run_id)
    phases = phase_telemetry.as_payload()
    named_phase_seconds = 0.0
    for item in phases:
        elapsed_value = item.get("elapsed_seconds") if isinstance(item, dict) else None
        if isinstance(elapsed_value, int | float):
            named_phase_seconds += float(elapsed_value)

    # After native integration these two Python timers contain only rows that
    # requested fallback. Report that cost separately, then add the native FFI
    # time for the complete current classification cost. The legacy
    # rust_addressable_* keys remain for before/after report compatibility,
    # but now truthfully include both execution paths.
    python_fallback_delta_seconds = (
        formula_telemetry.extension_seconds + formula_telemetry.wrapper_reference_seconds
    )
    rust_addressable_narrow_seconds = (
        formula_telemetry.native_delta_seconds + python_fallback_delta_seconds
    )
    rust_addressable_broad_seconds = (
        rust_addressable_narrow_seconds + formula_telemetry.normalization_seconds
    )

    report = {
        "label": args.label,
        "run_id": artifacts.run_id,
        "elapsed_seconds": elapsed,
        # Worker-process CPU only. The supervising parent separately samples
        # the complete worker/descendant process tree for RSS and Office use.
        "cpu_user_seconds": cpu_after.user - cpu_before.user,
        "cpu_system_seconds": cpu_after.system - cpu_before.system,
        "findings": len(result.findings),
        "ordered_evidence_digest": ordered_digest,
        "phases": phases,
        "named_phase_seconds": named_phase_seconds,
        "phase_residual_seconds": max(0.0, elapsed - named_phase_seconds),
        "recording_history_subphases": recording_history_subphases,
        "review_counts": dict(record.review_counts),
        "pattern_review_counts": dict(record.pattern_review_counts),
        "story_counts": dict(record.story_counts),
        "perform_run_telemetry": {
            "initial_hash_seconds": perform_telemetry.initial_hash_seconds,
            "qc_seconds": perform_telemetry.qc_seconds,
            "rehash_seconds": perform_telemetry.rehash_seconds,
            "reports_seconds": perform_telemetry.reports_seconds,
            "history_seconds": perform_telemetry.history_seconds,
            "total_seconds": perform_telemetry.total_seconds,
            "residual_seconds": perform_telemetry.residual_seconds,
        },
        "population_telemetry": {
            "construction_seconds": population_telemetry.construction_seconds,
            "spill_seconds": population_telemetry.spill_seconds,
            "finalize_seconds": population_telemetry.finalize_seconds,
        },
        "pair_key_telemetry": {
            "changed_pairs": pair_key_telemetry.changed_pairs,
            "distinct_pair_keys": pair_key_telemetry.distinct_pair_keys,
            "classification_conflicts": pair_key_telemetry.classification_conflicts,
            "frequency_histogram": pair_key_telemetry.frequency_histogram(),
        },
        # plan-20260910 Step 7 precondition evidence -- aggregate-only
        # counts/durations, no formula/sheet/coordinate (see
        # FormulaComparisonTelemetry's own docstring).
        "formula_comparison_telemetry": {
            "extension_seconds": formula_telemetry.extension_seconds,
            "wrapper_reference_seconds": formula_telemetry.wrapper_reference_seconds,
            "normalization_seconds": formula_telemetry.normalization_seconds,
            "complexity_assessment_seconds": (
                formula_telemetry.complexity_assessment_seconds
            ),
            "finding_construction_seconds": (
                formula_telemetry.finding_construction_seconds
            ),
            "paired_traversal_seconds": formula_telemetry.paired_traversal_seconds,
            "consistency_seconds": formula_telemetry.consistency_seconds,
            "error_scan_seconds": formula_telemetry.error_scan_seconds,
            "pair_analysis_memo_hits": formula_telemetry.pair_analysis_memo_hits,
            "pair_analysis_memo_misses": formula_telemetry.pair_analysis_memo_misses,
            "native_delta_seconds": formula_telemetry.native_delta_seconds,
            "native_delta_batches": formula_telemetry.native_delta_batches,
            "native_delta_batch_failures": (
                formula_telemetry.native_delta_batch_failures
            ),
            "native_delta_api_failures": (
                formula_telemetry.native_delta_api_failures
            ),
            "native_delta_protocol_failures": (
                formula_telemetry.native_delta_protocol_failures
            ),
            "native_delta_runtime_failures": (
                formula_telemetry.native_delta_runtime_failures
            ),
            "native_delta_supported_pairs": (
                formula_telemetry.native_delta_supported_pairs
            ),
            "native_delta_fallback_pairs": (
                formula_telemetry.native_delta_fallback_pairs
            ),
            "native_delta_declared_unsupported_pairs": (
                formula_telemetry.native_delta_declared_unsupported_pairs
            ),
            "native_delta_invalid_output_pairs": (
                formula_telemetry.native_delta_invalid_output_pairs
            ),
            "native_delta_oversized_pairs": (
                formula_telemetry.native_delta_oversized_pairs
            ),
            "python_fallback_delta_seconds": python_fallback_delta_seconds,
            "formula_delta_classification_seconds": (
                rust_addressable_narrow_seconds
            ),
            "rust_addressable_narrow_seconds": rust_addressable_narrow_seconds,
            "rust_addressable_broad_seconds": rust_addressable_broad_seconds,
        },
        "requested_output_mode": result.requested_output_mode.value,
        "resolved_output_policy": (
            result.resolved_output_policy.model_dump(mode="json")
            if result.resolved_output_policy is not None
            else None
        ),
        "formula_engines": dict(result.formula_engines),
        "values_engines": dict(result.values_engines),
        # Raw disclosures can legitimately carry selected sheet names. Keep
        # only their count in this private-workload aggregate report.
        "disclosure_count": len(result.disclosures),
    }
    return report


def _acceptance_child(args: argparse.Namespace, result_queue: Any) -> None:
    logging.disable(logging.CRITICAL)
    try:
        result_queue.put(("ok", _execute_acceptance(args)))
    except BaseException as exc:
        result_queue.put(("error", type(exc).__name__))


def _write_report(path: Path, report: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(report, indent=2, sort_keys=True), encoding="utf-8")


def main() -> int:
    # Disabled before parent or child touches a private path. Kept inside the
    # entry point so importing this module never silences another test/logger.
    logging.disable(logging.CRITICAL)
    args = _parser().parse_args()
    if args.deadline_seconds <= 0:
        _parser().error("--deadline-seconds must be positive")
    files = {
        "baseline_excel": args.baseline_excel,
        "current_excel": args.current_excel,
    }
    before_hashes = {role: _sha256(path) for role, path in files.items()}
    office_before = _system_office_process_count()
    context = mp.get_context("spawn")
    result_queue = context.Queue()
    process = context.Process(target=_acceptance_child, args=(args, result_queue))
    supervised_started = time.perf_counter()
    process.start()
    supervision = _monitor_process(
        process,
        deadline_seconds=args.deadline_seconds,
    )
    supervised_elapsed = time.perf_counter() - supervised_started
    after_hashes = {role: _sha256(path) for role, path in files.items()}
    office_after = _system_office_process_count()
    common = {
        "label": args.label,
        "supervised_elapsed_seconds": supervised_elapsed,
        "sampled_process_tree_peak_rss_bytes": supervision.sampled_peak_rss_bytes,
        # Backward-compatible field name, now backed by periodic sampling.
        "combined_process_tree_peak_rss_bytes": supervision.sampled_peak_rss_bytes,
        "office_process_counts": {
            "system_before": office_before,
            "owned_peak": supervision.owned_office_process_peak,
            "system_after": office_after,
        },
        "source_hashes_unchanged": before_hashes == after_hashes,
    }
    if supervision.timed_out:
        report = {**common, "status": "timeout"}
        _write_report(args.output, report)
        print(
            f"aborted: exceeded the {args.deadline_seconds:g}s hard diagnostic deadline",
            file=sys.stderr,
        )
        return 4
    try:
        status, payload = result_queue.get(timeout=5.0)
    except queue.Empty:
        status, payload = "error", "MissingChildResult"
    finally:
        result_queue.close()
        result_queue.join_thread()
    if supervision.exit_code != 0 or status != "ok" or not isinstance(payload, dict):
        report = {
            **common,
            "status": "error",
            "error_type": str(payload) if status == "error" else "ChildProcessError",
        }
        _write_report(args.output, report)
        print(f"aborted: acceptance child failed ({report['error_type']})", file=sys.stderr)
        return 2
    report = {**payload, **common, "status": "complete"}
    _write_report(args.output, report)
    print(
        json.dumps(
            {
                key: value
                for key, value in report.items()
                if key not in {"phases", "recording_history_subphases"}
            },
            indent=2,
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
