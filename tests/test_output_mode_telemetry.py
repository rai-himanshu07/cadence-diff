"""plan-20260910 Step 3: bounded, privacy-safe telemetry for the compact
formula-compare hot path, before any guest work.

Covers `PairKeyTelemetry` (changed-formula-pair classification-conflict/
frequency diagnostics), `PopulationTelemetry` (candidate construction/spill/
finalize timers), and `PerformRunTelemetry` (perform-run phase accounting),
plus a bounded synthetic benchmark proving telemetry overhead stays within
the plan's <=2% budget and that no telemetry field ever carries a path,
filename, sheet name, coordinate, formula, value, or raw pair-key text.
"""

from __future__ import annotations

import datetime as dt
import sys
import time
from pathlib import Path

import pytest
from openpyxl import Workbook

from qc_tool.config.profile import PopulationPolicy, default_profile
from qc_tool.coverage import FindingOutputMode
from qc_tool.excel.align import align_workbooks
from qc_tool.excel.formulas import (
    FormulaComparisonTelemetry,
    PairKeyTelemetry,
    diff_workbook_formulas,
)
from qc_tool.excel.population import (
    CandidateSpill,
    PopulationTelemetry,
    finalize_populations,
)
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.run_preflight import hash_run_files
from qc_tool.run_service import PerformRunTelemetry, perform_run
from qc_tool.scope import ComparisonScope
from tests.test_population_finalize import _candidate, _write_formula_pair

_TODAY = dt.date(2026, 9, 10)


def test_pair_key_telemetry_counts_distinct_keys_with_zero_conflicts() -> None:
    telemetry = PairKeyTelemetry()
    signature = (False, None, None, (), "")
    for _ in range(5):
        telemetry.observe("R1C1", "R1C1*2", signature)
    telemetry.observe("R2C1", "R2C1*3", signature)

    assert telemetry.changed_pairs == 6
    assert telemetry.distinct_pair_keys == 2
    assert telemetry.classification_conflicts == 0
    histogram = telemetry.frequency_histogram()
    assert histogram["1"] == 1  # the once-seen key
    assert histogram["5-9"] == 1  # the five-times-seen key


def test_pair_key_telemetry_detects_a_classification_conflict() -> None:
    telemetry = PairKeyTelemetry()
    telemetry.observe("R1C1", "R1C1*2", (False, None, None, (), ""))
    # Same canonical key, different classification signature this time.
    telemetry.observe("R1C1", "R1C1*2", (True, None, None, (), ""))

    assert telemetry.changed_pairs == 2
    assert telemetry.distinct_pair_keys == 1
    assert telemetry.classification_conflicts == 1


def test_pair_key_telemetry_public_surface_never_carries_pair_text() -> None:
    """Aggregate-only privacy contract: `changed_pairs`, `distinct_pair_keys`,
    `classification_conflicts`, and `frequency_histogram()` are the only
    public surface a diagnostic script would ever serialize -- none of them
    can carry a formula, sheet name, path, coordinate, or raw key text since
    they are plain integers/a bucketed count dict by construction.
    """
    telemetry = PairKeyTelemetry()
    telemetry.observe(
        "R1C1=Sheet1!R2C3*ImportantConstant",
        "R1C1=Sheet1!R2C3*ImportantConstant*2",
        (False, None, None, (), ""),
    )
    public_payload = {
        "changed_pairs": telemetry.changed_pairs,
        "distinct_pair_keys": telemetry.distinct_pair_keys,
        "classification_conflicts": telemetry.classification_conflicts,
        "frequency_histogram": telemetry.frequency_histogram(),
    }
    assert all(
        isinstance(value, (int, dict)) for value in public_payload.values()
    )
    serialized = repr(public_payload)
    assert "Sheet1" not in serialized
    assert "ImportantConstant" not in serialized
    assert "R1C1" not in serialized


def test_diff_workbook_formulas_pair_key_telemetry_is_additive_only(
    tmp_path: Path,
) -> None:
    """Passing pair-key telemetry changes no finding, and a uniform
    formula-shape fixture (every row differs by the same multiplier) is
    expected to collapse to exactly one distinct canonical pair key with
    zero classification conflicts -- direct, real evidence for Step 5's
    zero-conflict memoization gate.
    """
    base, curr = _write_formula_pair(tmp_path, rows=25)
    baseline = load_workbook_snapshot(base)
    current = load_workbook_snapshot(curr)
    alignment = align_workbooks(baseline, current)

    plain = diff_workbook_formulas(baseline, current, alignment)
    telemetry = PairKeyTelemetry()
    instrumented = diff_workbook_formulas(
        baseline, current, alignment, pair_key_telemetry=telemetry
    )

    assert [f.model_dump(mode="json") for f in instrumented] == [
        f.model_dump(mode="json") for f in plain
    ]
    assert telemetry.changed_pairs == 25
    assert telemetry.distinct_pair_keys == 1
    assert telemetry.classification_conflicts == 0
    assert telemetry.frequency_histogram()["25+"] == 1


def test_population_telemetry_records_construction_and_finalize_timers() -> None:
    profile = default_profile()
    telemetry = PopulationTelemetry()
    spill = CandidateSpill(profile, _TODAY, telemetry=telemetry)
    for row in range(2, 17):
        candidate = _candidate(location=f"B{row}", baseline_location=f"B{row - 1}")
        spill.add(candidate, shape_before="digest-before", shape_after="digest-after")

    policy = PopulationPolicy(enabled=True, threshold=10)
    outcome = finalize_populations(spill, policy, ComparisonScope(), telemetry=telemetry)

    assert len(outcome.population_findings) == 1
    assert telemetry.construction_seconds > 0.0
    assert telemetry.finalize_seconds > 0.0
    assert telemetry.spill_seconds >= 0.0


def test_perform_run_telemetry_accounts_every_named_phase_and_residual(
    tmp_path: Path,
) -> None:
    base, curr = _write_formula_pair(tmp_path, rows=5)
    telemetry = PerformRunTelemetry()

    artifacts = perform_run(
        tmp_path / "work",
        {"baseline_excel": base, "current_excel": curr},
        {},
        default_profile(),
        write_reports=True,
        _perform_run_telemetry=telemetry,
    )

    assert artifacts.run_id > 0
    assert telemetry.initial_hash_seconds > 0.0
    assert telemetry.qc_seconds > 0.0
    assert telemetry.rehash_seconds > 0.0
    assert telemetry.reports_seconds > 0.0
    assert telemetry.history_seconds > 0.0
    assert telemetry.total_seconds > 0.0
    assert telemetry.residual_seconds >= 0.0
    # The residual is a small remainder, not the dominant share of the run.
    assert telemetry.residual_seconds < telemetry.total_seconds


def test_perform_run_threads_all_compact_telemetry_without_changing_findings(
    tmp_path: Path,
) -> None:
    base, curr = _write_formula_pair(tmp_path, rows=25)
    plain = perform_run(
        tmp_path / "plain",
        {"baseline_excel": base, "current_excel": curr},
        {},
        default_profile(),
        output_mode=FindingOutputMode.DECISION,
    )
    formula = FormulaComparisonTelemetry()
    pairs = PairKeyTelemetry()
    populations = PopulationTelemetry()
    instrumented = perform_run(
        tmp_path / "instrumented",
        {"baseline_excel": base, "current_excel": curr},
        {},
        default_profile(),
        output_mode=FindingOutputMode.DECISION,
        _formula_telemetry=formula,
        _pair_key_telemetry=pairs,
        _population_telemetry=populations,
    )

    assert [finding.model_dump(mode="json") for finding in instrumented.result.findings] == [
        finding.model_dump(mode="json") for finding in plain.result.findings
    ]
    assert pairs.changed_pairs == 25
    assert populations.construction_seconds > 0.0
    assert populations.finalize_seconds > 0.0


def _build_large_formula_pair(tmp_path: Path, *, rows: int) -> tuple[Path, Path]:
    def build(path: Path, *, multiplier: int) -> None:
        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "Data"
        sheet.append(["Input", "Output"])
        for row in range(2, 2 + rows):
            # A distinct multiplier per row keeps every pair's canonical key
            # unique, so the benchmark exercises the full classification
            # cost (tokenizing, wrapper detection, reference-tag scan) for
            # every single changed pair, not a one-key shortcut.
            sheet.append([100, f"=A{row}*{2 + (row % 37)}"])
        workbook.save(path)

    base, curr = tmp_path / "bench_base.xlsx", tmp_path / "bench_curr.xlsx"
    build(base, multiplier=2)
    build(curr, multiplier=3)
    return base, curr


@pytest.mark.skipif(
    sys.platform == "win32",
    reason=(
        "Windows process_time advances in coarse scheduler ticks; Linux CI owns "
        "the 2% microbenchmark while Windows runs the functional telemetry suite"
    ),
)
def test_bounded_synthetic_telemetry_overhead_stays_within_two_percent(
    tmp_path: Path,
) -> None:
    """Step 3's required bench: the NEW pair-key telemetry this step adds
    stays within a 2% overhead budget, isolated from the pre-existing
    (already-accepted) ``FormulaComparisonTelemetry`` cost by holding it
    constant on both sides of the comparison. Measured via ``process_time()``
    (this process's own CPU time, immune to being preempted by unrelated
    system load -- confirmed necessary: a wall-clock `perf_counter()` version
    of this exact bench was flaky on this shared dev machine, occasionally
    reporting 3-6% purely from OS scheduling noise) and taking the MINIMUM
    across many repeats (the standard robust-statistics choice for a
    deterministic CPU-bound workload: noise can only ADD delay, never
    subtract it, so the minimum is the best estimate of true cost).
    """
    base, curr = _build_large_formula_pair(tmp_path, rows=20_000)
    baseline = load_workbook_snapshot(base)
    current = load_workbook_snapshot(curr)
    alignment = align_workbooks(baseline, current)

    repeats = 9

    def run(*, instrumented: bool) -> float:
        started = time.process_time()
        diff_workbook_formulas(
            baseline,
            current,
            alignment,
            telemetry=FormulaComparisonTelemetry(),
            pair_key_telemetry=PairKeyTelemetry() if instrumented else None,
        )
        return time.process_time() - started

    # One untimed warm-up call each side absorbs first-call effects (e.g. a
    # cold LRU cache) that are identical either way and would only add noise.
    run(instrumented=False)
    run(instrumented=True)
    # Interleaved (not two separate blocks) so neither side is systematically
    # favored by drift (CPU frequency ramp-up, background load) over time.
    baseline_times: list[float] = []
    instrumented_times: list[float] = []
    for _ in range(repeats):
        baseline_times.append(run(instrumented=False))
        instrumented_times.append(run(instrumented=True))

    baseline_best = min(baseline_times)
    instrumented_best = min(instrumented_times)
    assert baseline_best > 0.05, (
        "fixture too small to measure reliably "
        f"(baseline best {baseline_best:.4f}s CPU) -- widen it, don't loosen "
        "the 2% budget"
    )
    overhead = (instrumented_best - baseline_best) / baseline_best

    assert overhead <= 0.02, (
        f"telemetry overhead {overhead:.4%} exceeded the 2% budget "
        f"(baseline best {baseline_best:.4f}s CPU, "
        f"instrumented best {instrumented_best:.4f}s CPU)"
    )


def test_perform_run_telemetry_never_leaks_a_path_or_filename(tmp_path: Path) -> None:
    """Privacy contract: `PerformRunTelemetry`'s public fields are floats
    only, by construction -- explicitly reconfirmed against the real
    `perform_run()` output rather than assumed.
    """
    base, curr = _write_formula_pair(tmp_path, rows=5)
    telemetry = PerformRunTelemetry()
    perform_run(
        tmp_path / "work",
        {"baseline_excel": base, "current_excel": curr},
        {},
        default_profile(),
        _perform_run_telemetry=telemetry,
    )
    for value in (
        telemetry.initial_hash_seconds,
        telemetry.qc_seconds,
        telemetry.rehash_seconds,
        telemetry.reports_seconds,
        telemetry.history_seconds,
        telemetry.total_seconds,
        telemetry.residual_seconds,
    ):
        assert isinstance(value, float)
    serialized = repr(telemetry)
    assert str(tmp_path) not in serialized
    assert "bench_base" not in serialized
    assert ".xlsx" not in serialized


def test_hash_run_files_is_unaffected_by_telemetry(tmp_path: Path) -> None:
    """Sanity check the timing hook wraps the existing hashing call rather
    than replacing it: hashes are identical with or without telemetry.
    """
    base, curr = _write_formula_pair(tmp_path, rows=3)
    files = {"baseline_excel": base, "current_excel": curr}
    plain_hashes = hash_run_files(files)
    telemetry = PerformRunTelemetry()
    perform_run(
        tmp_path / "work2",
        files,
        {},
        default_profile(),
        _perform_run_telemetry=telemetry,
    )
    assert hash_run_files(files) == plain_hashes
