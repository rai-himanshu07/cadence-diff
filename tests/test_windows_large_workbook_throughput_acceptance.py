"""Supervision contracts for the private-workload Windows acceptance harness."""

from __future__ import annotations

import json
import multiprocessing as mp
import subprocess
import sys
import time
from pathlib import Path

import pytest

from scripts.windows_large_workbook_throughput_acceptance import _monitor_process
from tests.test_population_finalize import _write_formula_pair


def _sleep_child(seconds: float) -> None:
    time.sleep(seconds)


def _memory_child(mebibytes: int, seconds: float) -> None:
    payload = bytearray(mebibytes * 1024 * 1024)
    for index in range(0, len(payload), 4096):
        payload[index] = 1
    time.sleep(seconds)


def _supervise(target, args: tuple[object, ...], *, deadline: float):
    context = mp.get_context("spawn")
    process = context.Process(target=target, args=args)
    process.start()
    return _monitor_process(
        process,
        deadline_seconds=deadline,
        sample_interval_seconds=0.01,
    )


def test_monitor_enforces_deadline_without_progress_callbacks() -> None:
    started = time.monotonic()

    result = _supervise(_sleep_child, (5.0,), deadline=0.2)

    assert result.timed_out
    assert time.monotonic() - started < 3.0


def test_monitor_observes_a_child_memory_spike() -> None:
    baseline = _supervise(_sleep_child, (0.4,), deadline=5.0)
    spiked = _supervise(_memory_child, (48, 0.4), deadline=5.0)

    assert not baseline.timed_out and not spiked.timed_out
    assert spiked.sampled_peak_rss_bytes >= baseline.sampled_peak_rss_bytes + 32 * 1024**2


def test_supervised_harness_emits_complete_aggregate_report(tmp_path: Path) -> None:
    baseline, current = _write_formula_pair(tmp_path, rows=25)
    output = tmp_path / "acceptance.json"
    completed = subprocess.run(
        [
            sys.executable,
            "scripts/windows_large_workbook_throughput_acceptance.py",
            "--baseline-excel",
            str(baseline),
            "--current-excel",
            str(current),
            "--work-dir",
            str(tmp_path / "work"),
            "--label",
            "synthetic-acceptance",
            "--output",
            str(output),
            "--deadline-seconds",
            "30",
        ],
        cwd=Path(__file__).parents[1],
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stderr
    report = json.loads(output.read_text(encoding="utf-8"))
    assert report["status"] == "complete"
    assert report["source_hashes_unchanged"] is True
    assert report["sampled_process_tree_peak_rss_bytes"] > 0
    assert report["office_process_counts"]["owned_peak"] == 0
    assert report["formula_engines"]
    assert report["values_engines"]
    assert report["review_counts"]
    assert report["pair_key_telemetry"]["changed_pairs"] == 25
    assert report["population_telemetry"]["construction_seconds"] > 0
    assert report["perform_run_telemetry"]["total_seconds"] > 0
    formula_telemetry = report["formula_comparison_telemetry"]
    if formula_telemetry["native_delta_batches"]:
        assert formula_telemetry["native_delta_batches"] == 1
        assert formula_telemetry["native_delta_supported_pairs"] == 25
        assert formula_telemetry["native_delta_fallback_pairs"] == 0
        assert formula_telemetry["native_delta_seconds"] > 0.0
    else:
        assert formula_telemetry["native_delta_supported_pairs"] == 0
        assert formula_telemetry["native_delta_fallback_pairs"] == 0
        assert formula_telemetry["native_delta_seconds"] == 0.0
        assert formula_telemetry["python_fallback_delta_seconds"] > 0.0
    assert formula_telemetry["native_delta_batch_failures"] == 0
    assert formula_telemetry["native_delta_api_failures"] == 0
    assert formula_telemetry["native_delta_protocol_failures"] == 0
    assert formula_telemetry["native_delta_runtime_failures"] == 0
    assert formula_telemetry["native_delta_declared_unsupported_pairs"] == 0
    assert formula_telemetry["native_delta_invalid_output_pairs"] == 0
    assert formula_telemetry["native_delta_oversized_pairs"] == 0
    assert formula_telemetry["formula_delta_classification_seconds"] == pytest.approx(
        formula_telemetry["native_delta_seconds"]
        + formula_telemetry["python_fallback_delta_seconds"]
    )
    assert "disclosures" not in report
