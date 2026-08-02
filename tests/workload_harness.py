"""Isolated aggregate-only harness for private representative workloads.

Results belong under ignored ``artifacts/``. The harness never serializes input
paths, file names, hashes, formulas, values, labels, coordinates, or finding
messages. Example::

    conda run -n py311 python -m tests.workload_harness run \
        --workload-id cycle-pair --output artifacts/real-workload-step3 \
        --baseline-excel /private/baseline.xlsx \
        --current-excel /private/current.xlsx --repeats 3
"""

from __future__ import annotations

import argparse
import contextlib
import hashlib
import json
import re
import sqlite3
import statistics
import subprocess
import sys
import time
from collections import Counter
from pathlib import Path
from typing import Any, Literal

import psutil

from qc_tool.config.profile import default_profile, load_profile
from qc_tool.coverage import QCRunMode
from qc_tool.engine import run_qc
from qc_tool.excel.dependency import build_dependency_graph
from qc_tool.excel.formula_tokens import (
    FormulaPrecedentKind,
    extract_formula_precedents,
    tokenize_formula,
)
from qc_tool.excel.references import (
    ReferenceStatus,
    build_reference_index,
    reference_reason_code,
    resolve_reference,
)
from qc_tool.findings import FindingClass
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.progress import ProgressEvent
from qc_tool.review import (
    build_pattern_groups,
    build_review_groups,
    count_pattern_groups,
    review_counts,
)
from qc_tool.story import build_stories

_WORKLOAD_ID_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
_ROOT = Path(__file__).resolve().parents[1]
_ARTIFACTS = (_ROOT / "artifacts").resolve()
_AXIS_CLASSES = frozenset(
    {
        FindingClass.ROW_DELETED,
        FindingClass.ROW_INSERTED,
        FindingClass.ROW_GROWTH,
        FindingClass.ROW_KEY_CHANGED,
        FindingClass.COLUMN_DELETED,
        FindingClass.COLUMN_INSERTED,
        FindingClass.COLUMN_GROWTH,
        FindingClass.COLUMN_KEY_CHANGED,
    }
)
ReferenceMode = Literal["lexical", "raw-token"]


def _sha256(path: Path) -> bytes:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.digest()


def _raw_reference_values(formula: str) -> tuple[str, ...]:
    values: list[str] = []
    dynamic_stack: list[str | None] = []
    for token in tokenize_formula(formula):
        if token.type == "FUNC" and token.subtype == "OPEN":
            function = token.value[:-1].casefold()
            if function.endswith("anchorarray"):
                dynamic_stack.append("spill")
            elif function.endswith("single"):
                dynamic_stack.append("implicit")
            else:
                dynamic_stack.append(None)
            continue
        if token.type == "FUNC" and token.subtype == "CLOSE":
            if dynamic_stack:
                dynamic_stack.pop()
            continue
        if token.type != "OPERAND" or token.subtype != "RANGE":
            continue
        dynamic = dynamic_stack[-1] if dynamic_stack else None
        if dynamic == "spill":
            values.append(f"{token.value}#")
        elif dynamic == "implicit":
            values.append(f"@{token.value}")
        else:
            values.append(token.value)
    return tuple(values)


def _reference_census(path: Path, mode: ReferenceMode) -> dict[str, object]:
    workbook = load_workbook_snapshot(path, allow_large_workbook=True)
    reference_index = build_reference_index(workbook)
    statuses: Counter[str] = Counter()
    reasons: Counter[str] = Counter()
    formula_count = 0
    lexical_formula_count = 0
    range_membership = 0
    projected_edges = 0
    operand_count = 0
    for sheet in workbook.sheets:
        for (row, column), cell in sheet.cells.items():
            if cell.formula is None:
                continue
            formula_count += 1
            lowered = cell.formula.casefold()
            lexical_formula_count += "let(" in lowered or "lambda(" in lowered
            if mode == "raw-token":
                precedents = _raw_reference_values(cell.formula)
            else:
                precedents = extract_formula_precedents(cell.formula).precedents
            for operand in precedents:
                operand_count += 1
                if isinstance(operand, str):
                    value = operand
                elif operand.kind is FormulaPrecedentKind.LOCAL_SYMBOL:
                    statuses[FormulaPrecedentKind.LOCAL_SYMBOL.value] += 1
                    continue
                elif operand.kind is FormulaPrecedentKind.UNSUPPORTED:
                    statuses[ReferenceStatus.UNSUPPORTED.value] += 1
                    reasons[f"unsupported:{operand.reason}"] += 1
                    continue
                else:
                    value = operand.value
                resolution = resolve_reference(
                    workbook,
                    value,
                    host_sheet=sheet.name,
                    host_cell=(row, column),
                    require_within_sheet=False,
                    index=reference_index,
                )
                statuses[resolution.status.value] += 1
                if resolution.status is not ReferenceStatus.RESOLVED:
                    reason = reference_reason_code(resolution.status, resolution.detail)
                    reasons[f"{resolution.status.value}:{reason}"] += 1
                    continue
                for resolved in resolution.ranges:
                    range_membership += resolved.size
                    self_edge = int(
                        resolved.sheet == sheet.name
                        and resolved.min_row <= row <= resolved.max_row
                        and resolved.min_col <= column <= resolved.max_col
                    )
                    projected_edges += resolved.size - self_edge
    dependency_index = build_dependency_graph(workbook)
    return {
        "dependency_index": {
            "direct_edges": dependency_index.direct_edge_count,
            "materialized_member_edges": 0,
            "range_descriptors": dependency_index.range_descriptor_count,
            "symbolic_references": len(dependency_index.symbolic_references),
        },
        "formula_count": formula_count,
        "lexical_formula_count": lexical_formula_count,
        "operand_count": operand_count,
        "projected_concrete_edges": projected_edges,
        "range_membership": range_membership,
        "reference_reasons": dict(sorted(reasons.items())),
        "reference_statuses": dict(sorted(statuses.items())),
    }


class _PhaseRecorder:
    def __init__(self) -> None:
        self._process = psutil.Process()
        self._started: dict[str, float] = {}
        self._start_rss: dict[str, int] = {}
        self.metrics: dict[str, dict[str, float | int]] = {}

    def __call__(self, event: ProgressEvent) -> None:
        phase = event.phase.value
        now = time.perf_counter()
        rss = self._process.memory_info().rss
        self._started.setdefault(phase, now)
        self._start_rss.setdefault(phase, rss)
        metric = self.metrics.setdefault(
            phase,
            {
                "elapsed_seconds": 0.0,
                "peak_rss_bytes": rss,
                "rss_increment_bytes": 0,
                "start_rss_bytes": rss,
                "processed": 0,
                "total": 0,
            },
        )
        metric["elapsed_seconds"] = now - self._started[phase]
        metric["peak_rss_bytes"] = max(int(metric["peak_rss_bytes"]), rss)
        metric["rss_increment_bytes"] = max(
            0,
            int(metric["peak_rss_bytes"]) - self._start_rss[phase],
        )
        metric["processed"] = event.processed
        metric["total"] = event.total


def _mode(args: argparse.Namespace) -> QCRunMode:
    if args.baseline_excel is not None or args.baseline_ppt is not None:
        return QCRunMode.CYCLE_COMPARISON
    if args.current_excel is not None and args.current_ppt is not None:
        return QCRunMode.FINAL_PACKAGE
    return QCRunMode.CURRENT_FILE_PREFLIGHT


def _worker(args: argparse.Namespace) -> int:
    inputs = [
        path
        for path in (
            args.baseline_excel,
            args.current_excel,
            args.baseline_ppt,
            args.current_ppt,
        )
        if path is not None
    ]
    before = [_sha256(path) for path in inputs]
    profile = load_profile(args.profile) if args.profile is not None else default_profile()
    phases = _PhaseRecorder()
    started = time.perf_counter()
    result = run_qc(
        baseline_excel=args.baseline_excel,
        current_excel=args.current_excel,
        baseline_ppt=args.baseline_ppt,
        current_ppt=args.current_ppt,
        profile=profile,
        mode=_mode(args),
        allow_large_workbooks=True,
        on_progress=phases,
    )
    elapsed = time.perf_counter() - started
    grouped = review_counts(build_review_groups(result.findings))
    patterned = build_pattern_groups(result.findings)
    pattern_counts = count_pattern_groups(patterned)
    mixed_pattern_groups = sum(
        1
        for group in patterned
        if len(
            {
                (
                    member.finding_class,
                    member.severity,
                    member.expected_growth,
                    member.provenance,
                    member.subtype,
                )
                for member in group.members
            }
        )
        > 1
    )
    census = (
        _reference_census(args.current_excel, args.reference_mode)
        if args.current_excel is not None
        else {
            "dependency_index": {
                "direct_edges": 0,
                "materialized_member_edges": 0,
                "range_descriptors": 0,
                "symbolic_references": 0,
            },
            "formula_count": 0,
            "lexical_formula_count": 0,
            "operand_count": 0,
            "projected_concrete_edges": 0,
            "range_membership": 0,
            "reference_reasons": {},
            "reference_statuses": {},
        }
    )
    after = [_sha256(path) for path in inputs]
    finding_classes = Counter(finding.finding_class.value for finding in result.findings)
    severities = Counter(
        finding.severity.value
        for finding in result.findings
        if finding.severity is not None
    )
    provenance: Counter[str] = Counter()
    subtypes: Counter[str] = Counter()
    materiality: Counter[str] = Counter()
    temporal_context: Counter[str] = Counter()
    materiality_temporal: Counter[str] = Counter()
    expected_reasons: Counter[str] = Counter()
    evidence_tags: Counter[str] = Counter()
    axis_events: set[str] = set()
    object_events: set[str] = set()
    for finding in result.findings:
        key = f"{finding.finding_class.value}:"
        provenance[
            key + (finding.provenance.value if finding.provenance is not None else "unset")
        ] += 1
        subtypes[
            key + (finding.subtype.value if finding.subtype is not None else "unset")
        ] += 1
        materiality[
            key
            + (finding.materiality.value if finding.materiality is not None else "unset")
        ] += 1
        temporal_context[
            key
            + (
                finding.temporal_context.value
                if finding.temporal_context is not None
                else "unset"
            )
        ] += 1
        materiality_temporal[
            ":".join(
                (
                    finding.finding_class.value,
                    (
                        finding.materiality.value
                        if finding.materiality is not None
                        else "unset"
                    ),
                    (
                        finding.temporal_context.value
                        if finding.temporal_context is not None
                        else "unset"
                    ),
                )
            )
        ] += 1
        expected_reasons[
            key
            + (
                finding.expected_reason.value
                if finding.expected_reason is not None
                else "unset"
            )
        ] += 1
        for tag in finding.evidence_tags:
            evidence_tags[f"{key}{tag.value}"] += 1
        if finding.event_key:
            target = (
                axis_events if finding.finding_class in _AXIS_CLASSES else object_events
            )
            target.add(finding.event_key)
    stories = build_stories(result.findings)
    story_kinds: dict[str, dict[str, int]] = {}
    for story in stories:
        bucket = story_kinds.setdefault(story.kind.value, {"stories": 0, "members": 0})
        bucket["stories"] += 1
        bucket["members"] += story.member_count
    coverage = Counter(item.state.value for item in result.coverage)
    payload = {
        "atomic_findings": len(result.findings),
        "coverage_states": dict(sorted(coverage.items())),
        "elapsed_seconds": elapsed,
        "event_groups": {"axis": len(axis_events), "object": len(object_events)},
        "finding_classes": dict(sorted(finding_classes.items())),
        "finding_materiality": dict(sorted(materiality.items())),
        "finding_temporal_context": dict(sorted(temporal_context.items())),
        "finding_materiality_temporal": dict(sorted(materiality_temporal.items())),
        "finding_expected_reasons": dict(sorted(expected_reasons.items())),
        "finding_evidence_tags": dict(sorted(evidence_tags.items())),
        "finding_provenance": dict(sorted(provenance.items())),
        "finding_subtypes": dict(sorted(subtypes.items())),
        "mixed_pattern_groups": mixed_pattern_groups,
        "story_kinds": dict(sorted(story_kinds.items())),
        "pattern_review_counts": {
            severity.value: count
            for severity, count in sorted(
                pattern_counts.review_items.items(),
                key=lambda item: item[0].value,
            )
        },
        "phase_metrics": phases.metrics,
        "review_counts": {
            severity.value: count
            for severity, count in sorted(
                grouped.review_items.items(),
                key=lambda item: item[0].value,
            )
        },
        "severity": dict(sorted(severities.items())),
        "source_hashes_unchanged": before == after,
        **census,
    }
    print(json.dumps(payload, sort_keys=True))
    return 0


def _worker_command(args: argparse.Namespace) -> list[str]:
    command = [sys.executable, "-m", "tests.workload_harness", "worker"]
    for option in (
        "baseline_excel",
        "current_excel",
        "baseline_ppt",
        "current_ppt",
        "profile",
    ):
        value = getattr(args, option)
        if value is not None:
            command.extend((f"--{option.replace('_', '-')}", str(value)))
    command.extend(("--reference-mode", args.reference_mode))
    return command


def _sample_worker(args: argparse.Namespace) -> dict[str, Any]:
    process = psutil.Popen(
        _worker_command(args),
        cwd=_ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    peak_rss = 0
    while process.poll() is None:
        with contextlib.suppress(psutil.NoSuchProcess, psutil.AccessDenied):
            peak_rss = max(peak_rss, process.memory_info().rss)
        time.sleep(0.01)
    stdout, _stderr = process.communicate()
    if process.returncode != 0:
        raise RuntimeError(
            f"workload worker failed with {process.returncode}; "
            "private stderr was not persisted"
        )
    lines = [line for line in stdout.splitlines() if line.strip()]
    if not lines:
        raise RuntimeError("workload worker returned no aggregate metrics")
    metrics = json.loads(lines[-1])
    metrics["peak_rss_bytes"] = peak_rss
    return metrics


def _history_inputs(database: Path, run_id: int) -> dict[str, Path]:
    with sqlite3.connect(database) as connection:
        row = connection.execute(
            "SELECT file_paths FROM runs WHERE id = ?",
            (run_id,),
        ).fetchone()
    if row is None:
        raise ValueError("anonymous history run was not found")
    payload = json.loads(row[0] or "{}")
    if not isinstance(payload, dict):
        raise ValueError("anonymous history run has invalid file metadata")
    allowed = {
        "baseline_excel",
        "current_excel",
        "baseline_ppt",
        "current_ppt",
    }
    if not payload or not set(payload).issubset(allowed):
        raise ValueError("anonymous history run has unsupported file roles")
    inputs = {
        role: Path(value)
        for role, value in payload.items()
        if isinstance(role, str) and isinstance(value, str)
    }
    if len(inputs) != len(payload) or not all(path.is_file() for path in inputs.values()):
        raise ValueError("anonymous history run inputs are unavailable")
    return inputs


def _run(args: argparse.Namespace) -> int:
    if _WORKLOAD_ID_RE.fullmatch(args.workload_id) is None:
        raise ValueError("workload ID must be lowercase kebab-case")
    output = args.output.resolve()
    if not output.is_relative_to(_ARTIFACTS):
        raise ValueError("workload output must be beneath the ignored artifacts directory")
    output.mkdir(parents=True, exist_ok=False)
    records: list[dict[str, Any]] = []
    for repeat in range(args.repeats):
        record = {
            "workload_id": args.workload_id,
            "repeat": repeat,
            **_sample_worker(args),
        }
        records.append(record)
        with (output / "metrics.jsonl").open("a", encoding="utf-8") as stream:
            stream.write(json.dumps(record, sort_keys=True) + "\n")
    summary = {
        "workload_id": args.workload_id,
        "repeats": len(records),
        "median_elapsed_seconds": statistics.median(
            float(record["elapsed_seconds"]) for record in records
        ),
        "median_peak_rss_bytes": statistics.median(
            int(record["peak_rss_bytes"]) for record in records
        ),
        "all_sources_unchanged": all(
            bool(record["source_hashes_unchanged"]) for record in records
        ),
    }
    (output / "summary.json").write_text(
        json.dumps(summary, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(summary, indent=2, sort_keys=True))
    return 0


def _add_inputs(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--baseline-excel", type=Path)
    parser.add_argument("--current-excel", type=Path)
    parser.add_argument("--baseline-ppt", type=Path)
    parser.add_argument("--current-ppt", type=Path)
    parser.add_argument("--profile", type=Path)
    parser.add_argument(
        "--reference-mode",
        choices=("lexical", "raw-token"),
        default="lexical",
    )


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    worker = commands.add_parser("worker")
    _add_inputs(worker)
    run = commands.add_parser("run")
    _add_inputs(run)
    run.add_argument("--workload-id", required=True)
    run.add_argument("--output", type=Path, required=True)
    run.add_argument("--repeats", type=int, default=3)
    run.add_argument("--history-run-id", type=int)
    run.add_argument("--history-db", type=Path, default=Path("data/history.sqlite3"))
    return parser


def main() -> int:
    args = _parser().parse_args()
    if args.command == "run" and args.history_run_id is not None:
        if any(
            getattr(args, option) is not None
            for option in (
                "baseline_excel",
                "current_excel",
                "baseline_ppt",
                "current_ppt",
            )
        ):
            raise ValueError("history-run-id cannot be combined with explicit inputs")
        for role, path in _history_inputs(args.history_db, args.history_run_id).items():
            setattr(args, role, path)
    supplied = sum(
        getattr(args, option) is not None
        for option in (
            "baseline_excel",
            "current_excel",
            "baseline_ppt",
            "current_ppt",
        )
    )
    if supplied == 0:
        raise ValueError("at least one workload input is required")
    if args.command == "worker":
        return _worker(args)
    if args.repeats < 1:
        raise ValueError("repeats must be positive")
    return _run(args)


if __name__ == "__main__":
    raise SystemExit(main())
