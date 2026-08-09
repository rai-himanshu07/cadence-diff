"""Step 7 checkpoint reconciliation.

The findings output budget this file originally covered was removed on
2026-08-08 (user decision: never bound stored findings). The frozen Step 7
oracle remains valid history: its ``budget_omitted`` was zero.
"""

from __future__ import annotations

import json
from pathlib import Path


def test_step7_checkpoint_reconciles_with_the_delta_ledger() -> None:
    step6 = json.loads(
        (Path(__file__).parent / "oracles" / "real_workload_step6.json").read_text(
            encoding="utf-8"
        )
    )
    step7 = json.loads(
        (Path(__file__).parent / "oracles" / "real_workload_step7.json").read_text(
            encoding="utf-8"
        )
    )
    ledger = json.loads(
        (Path(__file__).parent / "oracles" / "step_delta_ledger.json").read_text(
            encoding="utf-8"
        )
    )["steps"]["7"]

    assert step7["atomic_findings"] - step6["atomic_findings"] == (
        ledger["expected_atomic_delta"]
    )
    assert step7["budget_omitted"] == 0
    assert step7["severity"] == step6["severity"]
    assert step7["review_counts"] == step6["review_counts"]
    assert step7["coverage_states"] == step6["coverage_states"]
    assert step7["dependency_index"] == step6["dependency_index"]
    assert step7["median_peak_rss_mib"] <= step6["median_peak_rss_mib"] * 1.05
    assert step7["median_elapsed_seconds"] <= step6["median_elapsed_seconds"] * 1.10
    assert step7["source_hashes_unchanged"] is True
