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
import re
import sys
import time
from pathlib import Path

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


def _tree_rss(process: psutil.Process) -> int:
    """Combined RSS of this process plus every live child (e.g. Excel)."""
    total = 0
    procs = [process]
    with contextlib.suppress(psutil.NoSuchProcess, psutil.AccessDenied):
        procs.extend(process.children(recursive=True))
    for proc in procs:
        with contextlib.suppress(
            psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess
        ):
            total += proc.memory_info().rss
    return total


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


def main() -> int:
    # A degraded-path warning inside qc_tool (e.g. a failed formula-worker
    # fallback) can embed the real source filename via %s formatting; Python's
    # default root-logger handler would otherwise print it straight to stderr.
    # Disabled before this script ever touches a real path, matching every
    # other private-data probe in this project. Deliberately NOT at module
    # level: importing this module for a future test (with only synthetic
    # fixtures) must not globally silence logging for the rest of the
    # pytest session.
    logging.disable(logging.CRITICAL)
    args = _parser().parse_args()

    from qc_tool.config.profile import default_profile, load_profile
    from qc_tool.coverage import FindingOutputMode, QCRunMode
    from qc_tool.excel.formulas import FormulaComparisonTelemetry
    from qc_tool.history.review_state import finding_evidence_digest
    from qc_tool.progress import (
        CancellationToken,
        PhaseTelemetry,
        ProgressEvent,
        RunCancelled,
        RunPhase,
    )
    from qc_tool.run_service import perform_run

    files = {
        "baseline_excel": args.baseline_excel,
        "current_excel": args.current_excel,
    }
    before_hashes = {role: _sha256(path) for role, path in files.items()}
    profile = load_profile(args.profile) if args.profile is not None else default_profile()

    process = psutil.Process()
    peak_tree_rss = _tree_rss(process)
    telemetry = PhaseTelemetry()
    formula_telemetry = FormulaComparisonTelemetry()
    recording_history_subphases: dict[str, float] = {}
    cancel_flag = CancellationToken()
    deadline_started = time.perf_counter()

    # Scoped to just RECORDING_HISTORY: profiling the whole perform_run() call
    # re-pays cProfile's per-call overhead on already-understood phases
    # (comparing_formulas, querying_impacts, building_review) for nothing.
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
        # Hard diagnostic deadline (Criterion 14): a content-free, bounded
        # cancellation past --deadline-seconds -- never an unbounded private
        # measurement run.
        if time.perf_counter() - deadline_started > args.deadline_seconds:
            cancel_flag.cancel()
        telemetry(event)

    started = time.perf_counter()
    cpu_before = process.cpu_times()
    try:
        artifacts = perform_run(
            args.work_dir,
            files,
            {},
            profile,
            mode=QCRunMode.CYCLE_COMPARISON,
            output_mode=FindingOutputMode(args.output_mode),
            # This acceptance harness exists to measure large-workbook (LARGE_WORKBOOK
            # -class) throughput; it always opts in rather than exposing a
            # toggle that would just make the script refuse its own purpose
            # (Criterion 18).
            allow_large_workbooks=True,
            write_reports=False,
            cancellation_token=cancel_flag,
            on_progress=on_progress,
            # Diagnostic only (see RunHistory.record_run's docstring): breaks
            # the single recording_history phase down into main_pass/story_
            # classify_and_replay/sqlite_write/storage_measurement/total so a
            # large-N cost concentration is attributable, not just a total.
            on_subphase=recording_history_subphases.__setitem__,
            # plan-20260910 Step 7 precondition evidence: aggregate-only
            # per-pair classification timings, never a formula/sheet/
            # coordinate (see FormulaComparisonTelemetry's own docstring).
            _formula_telemetry=formula_telemetry,
        )
    except RunCancelled:
        print(
            f"aborted: exceeded the {args.deadline_seconds:g}s hard diagnostic "
            "deadline",
            file=sys.stderr,
        )
        return 4
    if profiler is not None:
        if profiling_active:
            profiler.disable()
        profiler.dump_stats(str(args.cprofile_output))
    elapsed = time.perf_counter() - started
    cpu_after = process.cpu_times()
    peak_tree_rss = max(peak_tree_rss, _tree_rss(process))

    after_hashes = {role: _sha256(path) for role, path in files.items()}
    result = artifacts.result
    digests = [finding_evidence_digest(f) for f in result.findings]
    ordered_digest = hashlib.sha256("\n".join(digests).encode("utf-8")).hexdigest()

    # plan-20260910 Step 7 precondition evidence: the NARROW reading is only
    # the per-pair classification work Step 7's own scope proposes moving to
    # Rust (extension detection, wrapper detection, and the reference-delta
    # scan) -- excludes normalization (to_r1c1, already native-kernel-
    # addressable via a separate existing surface) and complexity_
    # assessment_seconds/finding_construction_seconds (both stay Python per
    # the plan's own explicit scope). The BROADER reading additionally
    # includes normalization for a conservative upper bound. Both are
    # reported so the Step 7 GO/NO-GO decision isn't pre-judged by this
    # script's own interpretation.
    rust_addressable_narrow_seconds = (
        formula_telemetry.extension_seconds + formula_telemetry.wrapper_reference_seconds
    )
    rust_addressable_broad_seconds = (
        rust_addressable_narrow_seconds + formula_telemetry.normalization_seconds
    )

    report = {
        "label": args.label,
        "run_id": artifacts.run_id,
        "elapsed_seconds": elapsed,
        # Parent-process CPU only -- a short-lived Excel COM child's own CPU
        # time is not attributable here without continuous polling across
        # its lifetime; elapsed_seconds and combined_process_tree_peak_rss_
        # bytes remain the authoritative wall-time/memory evidence.
        "cpu_user_seconds": cpu_after.user - cpu_before.user,
        "cpu_system_seconds": cpu_after.system - cpu_before.system,
        "findings": len(result.findings),
        "ordered_evidence_digest": ordered_digest,
        "phases": telemetry.as_payload(),
        "recording_history_subphases": recording_history_subphases,
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
            "rust_addressable_narrow_seconds": rust_addressable_narrow_seconds,
            "rust_addressable_broad_seconds": rust_addressable_broad_seconds,
        },
        "combined_process_tree_peak_rss_bytes": peak_tree_rss,
        "source_hashes_unchanged": before_hashes == after_hashes,
        # plan-20260910 Criterion 8: the run-level output-mode contract and
        # per-role resolved formula-engine identities, both purely
        # aggregate/informational -- a values-engine fallback (if any) is
        # already covered by the existing `disclosures` list below.
        "requested_output_mode": result.requested_output_mode.value,
        "resolved_output_policy": (
            result.resolved_output_policy.model_dump(mode="json")
            if result.resolved_output_policy is not None
            else None
        ),
        "formula_engines": dict(result.formula_engines),
        "disclosures": list(result.disclosures),
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2, sort_keys=True))
    print(json.dumps({k: v for k, v in report.items() if k != "phases"}, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
