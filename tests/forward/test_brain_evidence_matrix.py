"""Executable reconciliation for the anonymous Brain evidence matrix."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

_MATRIX_PATH = Path(__file__).parents[1] / "oracles" / "brain_evidence_matrix.json"


def _string_values(value: object) -> list[str]:
    if isinstance(value, str):
        return [value]
    if isinstance(value, dict):
        return [item for child in value.values() for item in _string_values(child)]
    if isinstance(value, list):
        return [item for child in value for item in _string_values(child)]
    return []


def test_anonymous_raw_evidence_matrix_reconciles() -> None:
    matrix: dict[str, Any] = json.loads(_MATRIX_PATH.read_text(encoding="utf-8"))
    assert matrix["schema_version"] == 1
    assert matrix["source_hashes_unchanged"] is True
    assert _string_values(matrix) == []

    cycle = matrix["workloads"]["cycle_pair"]
    numeric = cycle["numeric_constant_changes"]
    assert numeric["display_residue"] + numeric["substantive"] == numeric["total"]
    assert (
        numeric["recent_substantive"] + numeric["historical_substantive"]
        == numeric["substantive"]
    )
    formula = cycle["formula_changes"]
    assert formula["formula_to_formula"] + formula["added"] == formula["total"]
    assert formula["exact_baseline_subtree"] <= formula["formula_to_formula"]
    assert cycle["sideways_table_changes"] == {
        "added_columns": 2,
        "count": 1,
        "row_span_unchanged": True,
    }
    assert cycle["kpi_panel_count"] == 6
    assert cycle["changed_chart_part_count"] == 4
    assert cycle["changed_drawing_anchor_count"] == 1

    xlsb_a = matrix["workloads"]["xlsb_a"]
    assert (
        xlsb_a["raw_formula_record_count"]
        == xlsb_a["scanner_formula_coordinate_count"]
    )
    assert xlsb_a["cached_errors"]["na"] + xlsb_a["cached_errors"]["value"] == (
        xlsb_a["cached_errors"]["total"]
    )
    assert xlsb_a["sparse_value_only_mass"] == {
        "count": 1235,
        "formula_record_count": 0,
        "longest_run": 3,
        "run_count": 1139,
        "share_basis_points": 659,
    }
    assert xlsb_a["concentrated_formula_populations"]["count"] == 4

    xlsb_b = matrix["workloads"]["xlsb_b"]
    assert (
        xlsb_b["raw_formula_record_count"]
        == xlsb_b["scanner_formula_coordinate_count"]
    )
    assert xlsb_b["cached_error_count"] == 0
    assert xlsb_b["external_link_part_count"] == 2
    assert xlsb_b["workbook_external_link_relationship_count"] == 1
    assert xlsb_b["external_path_relationship_count"] == 4
