"""XLSB load-locality diagnostic
(plan-20260908-phase-b-guest-performance-followup.md, Step 2).

Measures how much of a workbook's load time is attributable to reading its
bytes from a possibly-slow shared/network path, versus decoding those bytes
once they are resident on a local disk. Loads the SAME source file twice via
``load_workbook_snapshot()`` -- once from wherever ``--source-excel`` points
(e.g. a VirtioFS-shared drive), once from a local staged copy under
``--work-dir`` -- and reports only timings, a copy-correctness boolean, and
a source-unchanged boolean. Also accepts ``--xlsb-values-engine`` so the
same probe can additionally measure the effect of routing XLSB values
through the native kernel instead of pyxlsb (plan-20260908's other open
question), on either side of the locality comparison.

Never prints or persists the source path, filename, sheet name, formula, or
cell content -- the staged local copy is written under a generic name
(``staged_source<suffix>``), never the source's own filename, and is deleted
after measurement unless ``--keep-staged-copy`` is given. Disables logging
before touching any path, matching every other private-data probe in this
project.

Scope: a single source Excel file only. Multi-package (member-qualified)
scenarios are intentionally unsupported here -- not partially threaded --
since this diagnostic's whole purpose is one-file load-locality attribution;
add a member-aware variant separately if that scope is ever actually needed
rather than half-wiring it into this script's flags.

Usage:

    python scripts/xlsb_load_locality_diagnostic.py \\
        --source-excel Z:\\path\\to\\current.xlsb \\
        --work-dir C:\\QC-Pilot\\load-locality-work \\
        --xlsb-values-engine auto \\
        --label large_workbook-current-shared-vs-local \\
        --output Z:\\QC_Tool\\windows-return\\load-locality-current.json
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import re
import shutil
import sys
import time
from pathlib import Path
from typing import Literal

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
            "(e.g. 'large_workbook-current-shared-vs-local'), never a filename or path"
        )
    return value


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def stage_local_copy(source: Path, work_dir: Path) -> Path:
    """Copy ``source`` into ``work_dir`` under a generic name (never the
    source's own filename) and return the staged path.
    """
    work_dir.mkdir(parents=True, exist_ok=True)
    staged = work_dir / f"staged_source{source.suffix}"
    shutil.copyfile(source, staged)
    return staged


def measure_load_seconds(
    path: Path, *, xlsb_values_engine: Literal["pyxlsb", "native", "auto"]
) -> float:
    from qc_tool.io.loader import load_workbook_snapshot

    started = time.perf_counter()
    load_workbook_snapshot(path, _xlsb_values_engine=xlsb_values_engine)
    return time.perf_counter() - started


def run_probe(
    source: Path,
    work_dir: Path,
    *,
    xlsb_values_engine: Literal["pyxlsb", "native", "auto"],
    keep_staged_copy: bool,
) -> dict[str, object]:
    source_hash_before = _sha256(source)
    staged = stage_local_copy(source, work_dir)
    try:
        copy_matches_source = _sha256(staged) == source_hash_before
        shared_path_load_seconds = measure_load_seconds(
            source, xlsb_values_engine=xlsb_values_engine
        )
        local_copy_load_seconds = measure_load_seconds(
            staged, xlsb_values_engine=xlsb_values_engine
        )
    finally:
        if not keep_staged_copy:
            staged.unlink(missing_ok=True)
    source_hash_unchanged = _sha256(source) == source_hash_before

    return {
        "xlsb_values_engine": xlsb_values_engine,
        "copy_matches_source": copy_matches_source,
        "source_hash_unchanged": source_hash_unchanged,
        "shared_path_load_seconds": shared_path_load_seconds,
        "local_copy_load_seconds": local_copy_load_seconds,
        "speedup_ratio_shared_over_local": (
            shared_path_load_seconds / local_copy_load_seconds
            if local_copy_load_seconds > 0
            else None
        ),
    }


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-excel", type=Path, required=True)
    parser.add_argument(
        "--work-dir",
        type=Path,
        required=True,
        help="Local directory to stage a same-bytes copy of --source-excel "
        "into (created if missing). Should be on local disk, not the shared "
        "path, so the comparison is meaningful.",
    )
    parser.add_argument(
        "--xlsb-values-engine",
        choices=("pyxlsb", "native", "auto"),
        default="pyxlsb",
        help="Values-decoding engine for XLSB inputs (default matches the "
        "shipped production default). Ignored for xlsx/xlsm inputs.",
    )
    parser.add_argument(
        "--label",
        required=True,
        type=_safe_label,
        help="Short generic scenario token, e.g. 'large_workbook-current-shared-vs-local' "
        "(lowercase alphanumeric/hyphen only, max 64 chars -- never a "
        "filename or path).",
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--keep-staged-copy",
        action="store_true",
        default=False,
        help="Keep the staged local copy after measurement (default: delete "
        "it) -- never useful across invocations since it is unnamed and "
        "would be overwritten by the next run anyway.",
    )
    return parser


def main() -> int:
    # Disabled before this script ever touches a real path, matching every
    # other private-data probe in this project (a degraded-path warning
    # inside qc_tool can embed the real source filename via %s formatting).
    # Deliberately NOT at module level: importing this module for tests
    # (with only synthetic fixtures) must not globally silence logging for
    # the rest of the pytest session.
    logging.disable(logging.CRITICAL)
    args = _parser().parse_args()

    report = run_probe(
        args.source_excel,
        args.work_dir,
        xlsb_values_engine=args.xlsb_values_engine,
        keep_staged_copy=args.keep_staged_copy,
    )
    report = {"label": args.label, **report}

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2, sort_keys=True))
    print(json.dumps(report, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
