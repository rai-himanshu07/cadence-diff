"""Memory-behavior probe for the part-processed diff pipeline plan.

Runs a deterministic monster pair through ``run_qc`` in a subprocess with an
optional address-space limit and records per-phase peak RSS. Evidence goes to
ignored ``artifacts/``; the fixture is generated on demand and never
committed.

Usage::

    conda run -n py311 python -m tests.memory_probe \
        --rows 40000 --columns 8 --limit-gb 4 --output artifacts/monster-probe
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
from pathlib import Path

_CHILD = "child"


def _vm_hwm_bytes() -> int:
    text = Path("/proc/self/status").read_text(encoding="ascii")
    for line in text.splitlines():
        if line.startswith("VmHWM:"):
            return int(line.split()[1]) * 1024
    return 0


def _child(arguments: argparse.Namespace) -> int:
    import importlib

    if arguments.limit_gb > 0:
        resource = importlib.import_module("resource")
        limit = int(arguments.limit_gb * 1024**3)
        resource.setrlimit(resource.RLIMIT_AS, (limit, limit))

    from qc_tool.engine import run_qc
    from qc_tool.progress import ProgressEvent

    phases: list[dict[str, object]] = []
    started = time.perf_counter()

    def on_progress(event: ProgressEvent) -> None:
        phases.append(
            {
                "phase": event.phase.value,
                "processed": event.processed,
                "total": event.total,
                "elapsed": round(time.perf_counter() - started, 2),
                "peak_rss_bytes": _vm_hwm_bytes(),
            }
        )

    result = run_qc(
        baseline_excel=Path(arguments.baseline),
        current_excel=Path(arguments.current),
        allow_large_workbooks=True,
        on_progress=on_progress,
    )
    payload = {
        "findings": len(result.findings),
        "severity": {
            severity.value: count for severity, count in result.counts.items()
        },
        "elapsed_seconds": round(time.perf_counter() - started, 2),
        "peak_rss_bytes": _vm_hwm_bytes(),
        "phases": phases,
    }
    print(json.dumps(payload))
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rows", type=int, default=40_000)
    parser.add_argument("--columns", type=int, default=8)
    parser.add_argument("--limit-gb", type=float, default=0.0)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--baseline", default="")
    parser.add_argument("--current", default="")
    parser.add_argument("--role", default="", help=argparse.SUPPRESS)
    arguments = parser.parse_args(argv)

    if arguments.role == _CHILD:
        return _child(arguments)

    from tests.fixtures.monster import generate_monster_pair

    fixture_dir = Path("tests/fixtures/generated/monster")
    baseline, current = generate_monster_pair(
        fixture_dir, rows=arguments.rows, columns=arguments.columns
    )
    arguments.output.mkdir(parents=True, exist_ok=True)

    command = [
        sys.executable,
        "-m",
        "tests.memory_probe",
        "--role",
        _CHILD,
        "--rows",
        str(arguments.rows),
        "--columns",
        str(arguments.columns),
        "--limit-gb",
        str(arguments.limit_gb),
        "--output",
        str(arguments.output),
        "--baseline",
        str(baseline),
        "--current",
        str(current),
    ]
    started = time.perf_counter()
    completed = subprocess.run(command, capture_output=True, text=True)
    record: dict[str, object] = {
        "rows": arguments.rows,
        "columns": arguments.columns,
        "expected_value_diffs": arguments.rows * (arguments.columns - 1),
        "limit_gb": arguments.limit_gb,
        "exit_code": completed.returncode,
        "wall_seconds": round(time.perf_counter() - started, 2),
    }
    if completed.returncode == 0 and completed.stdout.strip():
        record["result"] = json.loads(completed.stdout.strip().splitlines()[-1])
    else:
        record["stderr_tail"] = completed.stderr.strip().splitlines()[-8:]
    label = f"probe-r{arguments.rows}-l{arguments.limit_gb:g}.json"
    (arguments.output / label).write_text(
        json.dumps(record, indent=1), encoding="utf-8"
    )
    print(json.dumps(record, indent=1))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
