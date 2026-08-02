"""Brain-upgrade checkpoint: reconciles the anonymous run-13-pair oracle.

`real_workload_brain.json` freezes the materiality/provenance/wrapper/story
behavior on the representative cycle pair against the previous plan's final
checkpoint (`real_workload_step13.json`). Every delta must be attributable
to a Brain rule, never to accident.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

_ORACLE_DIR = Path(__file__).parent / "oracles"


def _oracle(name: str) -> dict[str, Any]:
    return json.loads((_ORACLE_DIR / name).read_text(encoding="utf-8"))


def test_brain_checkpoint_reconciles_with_step13() -> None:
    before = _oracle("real_workload_step13.json")
    brain = _oracle("real_workload_brain.json")

    assert brain["trials_consistent"] is True
    assert brain["source_hashes_unchanged"] is True
    assert brain["coverage_states"] == before["coverage_states"]

    # 13 in-place pairs merged plus ~200 alignment phantoms resolved (KPI-panel
    # triples, composite-key churn, and run-membership consistency artifacts —
    # the affected column truly holds two uniform section patterns).
    assert before["atomic_findings"] - brain["atomic_findings"] == 213
    assert sum(brain["key_changes"].values()) == 5
    assert brain["row_events"]["row_key_changed"] == 5
    assert brain["formula_inconsistent"] == 2

    # Severity totals move only through the documented rules.
    assert sum(brain["severity"].values()) == brain["atomic_findings"]
    assert brain["severity"]["critical"] <= 118
    assert brain["severity"]["expected"] == 7  # every one genuine cadence
    assert before["severity"]["critical"] == 335


def test_brain_value_tiers_cover_every_value_change() -> None:
    brain = _oracle("real_workload_brain.json")
    tiers = brain["value_materiality"]

    assert sum(tiers.values()) == 273  # every value change carries a tier state
    assert tiers["noise"] == 87  # the proven sub-ULP population, none critical
    assert tiers["recent_restatement"] > 0
    assert tiers["unset"] == 18  # non-numeric replacements have no tier


def test_brain_wrapper_rollout_collapses() -> None:
    brain = _oracle("real_workload_brain.json")

    assert brain["wrapped_formula_changes"] >= 0.95 * 1081
    assert brain["wrapper_skeletons"] <= 6


def test_brain_stories_partition_the_run() -> None:
    brain = _oracle("real_workload_brain.json")
    kinds = brain["story_kinds"]

    total_members = sum(bucket["members"] for bucket in kinds.values())
    total_stories = sum(bucket["stories"] for bucket in kinds.values())
    assert total_members == brain["atomic_findings"]
    assert total_stories <= 8

    residual = kinds["residual"]["members"]
    coverage = (total_members - residual) / total_members
    # Deliberate deviation from the plan's 90% target: conditional-format
    # adds/removes and chart-label changes stay residual because no evidence
    # links them to a driver; honesty beats the round number.
    assert coverage >= 0.88

    assert brain["mixed_pattern_groups"] == 0


def test_standalone_preflight_checkpoints() -> None:
    frozen = _oracle("standalone_preflight.json")

    reference_a = frozen["reference_a"]
    assert reference_a["severity"]["critical"] <= 112  # was 759 before demotion
    assert reference_a["columnar_error_population_atomics"] >= 600
    assert (
        reference_a["severity"]["warning"]
        >= reference_a["columnar_error_population_atomics"]
    )

    reference_b = frozen["reference_b"]
    assert reference_b["external_link_findings"] >= 1  # never a zero-finding 'clean'
    assert reference_b["atomic_findings"] >= reference_b["external_link_findings"]
    assert reference_b["severity"]["critical"] >= 1
