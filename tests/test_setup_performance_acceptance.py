"""Privacy and aggregate-output contracts for setup acceptance."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

import qc_tool.setup.preview_worker as preview_worker_module
from qc_tool.setup.preview_worker import SetupScanOutcome
from scripts.setup_performance_acceptance import (
    _discover_private_pairs,
    _measure_setup,
    _run_pair,
    _safe_label,
)


def _write_workbook(path: Path, value: int) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Synthetic"
    sheet.append(["ID", "Value"])
    sheet.append(["A", value])
    workbook.save(path)


def test_private_pair_discovery_returns_only_hashed_labels(tmp_path: Path) -> None:
    private_name = "private-client-name"
    pair_dir = tmp_path / private_name
    pair_dir.mkdir()
    _write_workbook(pair_dir / "before.xlsx", 1)
    _write_workbook(pair_dir / "after.xlsx", 2)

    pairs = _discover_private_pairs(tmp_path)

    assert len(pairs) == 1
    label, _baseline, _current = pairs[0]
    assert label.startswith("pair-")
    assert private_name not in label
    assert label == _safe_label(f"directory:{private_name}")


def test_setup_measurement_is_aggregate_only_and_sources_unchanged(
    tmp_path: Path,
) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _write_workbook(baseline, 1)
    _write_workbook(current, 2)

    result = _measure_setup("public-test", baseline, current)

    assert result["ok"] is True
    assert result["source_hashes_unchanged"] is True
    assert result["first_inventory_seconds"] is not None
    assert result["first_editable_sheet_seconds"] is not None
    phase_seconds = result["phase_seconds_by_side"]
    assert isinstance(phase_seconds, dict)
    assert set(phase_seconds) == {"baseline", "current"}
    serialized = str(result)
    assert str(baseline) not in serialized
    assert str(current) not in serialized
    assert "Synthetic" not in serialized
    assert "ID" not in serialized


def test_public_acceptance_does_not_require_setup_to_beat_tiny_full_load(
    monkeypatch,
) -> None:
    import argparse

    responses = iter(
        [
            {
                "label": "public-xlsx",
                "ok": True,
                "elapsed_seconds": 0.5,
                "first_inventory_seconds": 0.4,
                "first_editable_sheet_seconds": 0.45,
                "peak_process_tree_rss_bytes": 200_000_000,
                "source_hashes_unchanged": True,
            },
            {
                "label": "public-xlsx",
                "ok": True,
                "elapsed_seconds": 0.03,
                "peak_process_tree_rss_bytes": 100_000_000,
                "source_hashes_unchanged": True,
            },
        ]
    )
    monkeypatch.setattr(
        "scripts.setup_performance_acceptance._supervise",
        lambda _arguments: next(responses),
    )

    result = _run_pair(
        argparse.Namespace(public_format="xlsx", private_index=None)
    )

    assert result["setup_faster_than_full_preparation"] is False
    assert result["acceptance_passed"] is True


def test_setup_measurement_includes_child_process_cpu(
    monkeypatch, tmp_path: Path
) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _write_workbook(baseline, 1)
    _write_workbook(current, 2)
    monkeypatch.setattr(
        preview_worker_module,
        "run_setup_scan_worker",
        lambda *_args, **_kwargs: SetupScanOutcome(
            result_payload={"current_sheets": []},
            worker_cpu_seconds=1.25,
        ),
    )

    result = _measure_setup("public-test", baseline, current)

    assert isinstance(result["cpu_seconds"], float)
    assert result["cpu_seconds"] >= 1.25
