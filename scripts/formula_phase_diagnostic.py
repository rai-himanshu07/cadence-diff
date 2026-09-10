"""Formula-phase attribution + values-engine diagnostic
(plan-20260908-phase-b-guest-performance-followup.md, Step 1).

Surfaces ``qc_tool.excel.formulas.FormulaComparisonTelemetry`` -- including
the ``assess_workbook_complexity()`` cost that also lives inside the
``RunPhase.COMPARING_FORMULAS`` boundary but is not part of
``diff_workbook_formulas()`` itself -- for a real ``run_qc()`` invocation.
Optionally selects the XLSB values-decoding engine (``--xlsb-values-engine``)
so the same pair can be measured with ``pyxlsb`` (the shipped default) vs.
``native``/``auto``, to quantify B1's native values kernel's effect on load
time now that it can be requested end to end. Optionally engages the private
formula-extraction cache (``--formula-cache-dir``) and records the cache's
schema version plus whether the directory already held entries before this
run (``cold``/``warm``/``disabled``) -- never the directory path itself --
so a cold-before/cold-after or warm-before/warm-after comparison (plan
Criterion 4) can be told apart from a mixed-state measurement.

Never prints or persists a filename, sheet name, formula, coordinate, or
defined name -- only counts, timings, hashes, and the caller-supplied
``--label`` (restricted to a short, safe, generic token). Calls ``run_qc()``
directly rather than ``perform_run()``, mirroring this project's own
established precedent for private/diagnostic-only parameters (e.g. the
``_native_compat_mode`` oracle scripts) -- ``perform_run()`` deliberately has
no passthrough for these switches.

Scope: a single baseline/current Excel pair only. Multi-package (member
-qualified) scenarios are intentionally unsupported here -- not partially
threaded -- since this diagnostic's whole purpose is phase attribution for
one pair; add a member-aware variant separately if that scope is ever
actually needed rather than half-wiring it into this script's flags.

Usage:

    python scripts/formula_phase_diagnostic.py \\
        --baseline-excel /path/to/baseline.xlsb \\
        --current-excel /path/to/current.xlsb \\
        --xlsb-values-engine auto \\
        --label large_workbook-values-auto \\
        --output /path/to/result.json
"""

from __future__ import annotations

import argparse
import dataclasses
import hashlib
import json
import logging
import re
import sys
import time
from pathlib import Path

# Import qc_tool lazily after sys.path is set so this script can run from a
# plain `python scripts/...` invocation without an editable install.
_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(_ROOT))

#: `--label` is reported verbatim in the output JSON, so it is restricted to
#: this safe, bounded pattern -- it cannot hold a filename, path, sheet name,
#: or other free text.
_SAFE_LABEL_RE = re.compile(r"^[a-z0-9]+(-[a-z0-9]+){0,7}$")
_MAX_LABEL_LENGTH = 64


def _safe_label(value: str) -> str:
    if len(value) > _MAX_LABEL_LENGTH or not _SAFE_LABEL_RE.match(value):
        raise argparse.ArgumentTypeError(
            "label must be 1-64 lowercase alphanumeric/hyphen segments "
            "(e.g. 'large_workbook-values-auto'), never a filename or path"
        )
    return value


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-excel", type=Path, required=True)
    parser.add_argument("--current-excel", type=Path, required=True)
    parser.add_argument(
        "--xlsb-values-engine",
        choices=("pyxlsb", "native", "auto"),
        default="pyxlsb",
        help="Values-decoding engine for XLSB inputs (default matches the "
        "shipped production default; 'auto'/'native' measure the native "
        "kernel's effect). Ignored for xlsx/xlsm inputs.",
    )
    parser.add_argument(
        "--label",
        required=True,
        type=_safe_label,
        help="Short generic scenario token, e.g. 'large_workbook-values-auto' "
        "(lowercase alphanumeric/hyphen only, max 64 chars -- never a "
        "filename or path).",
    )
    parser.add_argument(
        "--formula-cache-dir",
        type=Path,
        default=None,
        help="Optional formula-extraction cache directory. Omitted means no "
        "cache is used (matches this script's original behavior). The report "
        "discloses the cache's schema version and whether it already held "
        "entries before this run -- never the directory path.",
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--profile", type=Path, default=None, help="Optional named profile YAML."
    )
    return parser


def main() -> int:
    # Disabled before this script ever touches a real path, matching every
    # other private-data probe in this project (a degraded-path warning
    # inside qc_tool can embed the real source filename via %s formatting).
    # Deliberately NOT at module level: importing this module for a future
    # test (with only synthetic fixtures) must not globally silence logging
    # for the rest of the pytest session.
    logging.disable(logging.CRITICAL)
    args = _parser().parse_args()

    from qc_tool.config.profile import default_profile, load_profile
    from qc_tool.coverage import QCRunMode
    from qc_tool.engine import run_qc
    from qc_tool.excel.formulas import FormulaComparisonTelemetry
    from qc_tool.io.formula_cache import CACHE_SCHEMA_VERSION, FormulaExtractionCache

    files = {
        "baseline_excel": args.baseline_excel,
        "current_excel": args.current_excel,
    }
    before_hashes = {role: _sha256(path) for role, path in files.items()}
    profile = (
        load_profile(args.profile) if args.profile is not None else default_profile()
    )

    cache: FormulaExtractionCache | None = None
    cache_state = "disabled"
    if args.formula_cache_dir is not None:
        cache_state = "warm" if args.formula_cache_dir.is_dir() else "cold"
        cache = FormulaExtractionCache(args.formula_cache_dir)

    telemetry = FormulaComparisonTelemetry()
    started = time.perf_counter()
    result = run_qc(
        baseline_excel=args.baseline_excel,
        current_excel=args.current_excel,
        profile=profile,
        mode=QCRunMode.CYCLE_COMPARISON,
        # This diagnostic exists to measure large-workbook formula-phase
        # throughput; it always opts in rather than exposing a toggle that
        # would just make the script refuse its own purpose (Criterion 18).
        allow_large_workbooks=True,
        formula_cache=cache,
        _formula_telemetry=telemetry,
        _xlsb_values_engine=args.xlsb_values_engine,
    )
    elapsed = time.perf_counter() - started

    after_hashes = {role: _sha256(path) for role, path in files.items()}

    report = {
        "label": args.label,
        "xlsb_values_engine": args.xlsb_values_engine,
        "cache_state": cache_state,
        "cache_schema_version": CACHE_SCHEMA_VERSION,
        "elapsed_seconds": elapsed,
        "findings": len(result.findings),
        "formula_comparison_telemetry": dataclasses.asdict(telemetry),
        "source_hashes_unchanged": before_hashes == after_hashes,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2, sort_keys=True))
    print(json.dumps(report, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
