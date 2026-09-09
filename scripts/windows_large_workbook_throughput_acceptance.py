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
"""

from __future__ import annotations

import argparse
import contextlib
import hashlib
import json
import logging
import re
import sys
import time
from pathlib import Path

import psutil

# A degraded-path warning inside qc_tool (e.g. a failed formula-worker
# fallback) can embed the real source filename via %s formatting; Python's
# default root-logger handler would otherwise print it straight to stderr.
# Disabled before this script ever touches a real path, matching every
# other private-data probe in this project.
logging.disable(logging.CRITICAL)

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
    parser.add_argument("--allow-large-workbooks", action="store_true", default=True)
    return parser


def main() -> int:
    args = _parser().parse_args()

    from qc_tool.config.profile import default_profile, load_profile
    from qc_tool.coverage import QCRunMode
    from qc_tool.history.review_state import finding_evidence_digest
    from qc_tool.progress import PhaseTelemetry
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

    started = time.perf_counter()
    cpu_before = process.cpu_times()
    artifacts = perform_run(
        args.work_dir,
        files,
        {},
        profile,
        mode=QCRunMode.CYCLE_COMPARISON,
        allow_large_workbooks=args.allow_large_workbooks,
        write_reports=False,
        on_progress=telemetry,
    )
    elapsed = time.perf_counter() - started
    cpu_after = process.cpu_times()
    peak_tree_rss = max(peak_tree_rss, _tree_rss(process))

    after_hashes = {role: _sha256(path) for role, path in files.items()}
    result = artifacts.result
    digests = [finding_evidence_digest(f) for f in result.findings]
    ordered_digest = hashlib.sha256("\n".join(digests).encode("utf-8")).hexdigest()

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
        "combined_process_tree_peak_rss_bytes": peak_tree_rss,
        "source_hashes_unchanged": before_hashes == after_hashes,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2, sort_keys=True))
    print(json.dumps({k: v for k, v in report.items() if k != "phases"}, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
