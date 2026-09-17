"""Ranked-table blocked-action parsing for workspace recovery."""

from __future__ import annotations

from qc_tool.ui.ranked_table_dialog import (
    ranked_regions_from_action,
    region_draft_from_item,
)

_V2_ITEM: dict[str, object] = {
    "member_id": "primary",
    "sheet": "Panel",
    "cell": "A1",
    "ranked_table_evidence": {
        "version": 2,
        "member_id": "primary",
        "sheet": "Panel",
        "current_range": "A1:E6001",
        "data_row_count": 6000,
        "header_row": 1,
        "available_columns": ["A", "B", "C", "D", "E"],
        "column_headers": ["Rank", "Record ID", "Value", "Value 2", "Value 3"],
        "suggested_identity_columns": ["B"],
        "suggested_ordinal_columns": ["A"],
        "non_blank_coverage": 0.999,
        "unique_ratio": 0.998,
        "key_overlap": 0.95,
        "formula_ratio": 0.0,
        "displaced_ratio": 1.0,
        "mismatch_reduction": 0.995,
        "projected_positional_mismatches": 500_000,
        "projected_avoided_mismatches": 497_500,
    },
}

_V1_ITEM: dict[str, object] = {
    "member_id": "primary",
    "sheet": "Legacy",
    "cell": "A1",
    "detail": "bounded legacy detail",
    "suggested_identity_columns": ["B"],
    "suggested_ordinal_columns": ["A"],
}


def test_region_draft_from_v2_item_carries_workspace_proposal_fields() -> None:
    region = region_draft_from_item(_V2_ITEM)

    assert region is not None
    assert region.member_id == "primary"
    assert region.sheet == "Panel"
    assert region.anchor_cell == "A1"
    assert region.current_range == "A1:E6001"
    assert region.header_row == 1
    assert region.available_columns == ("A", "B", "C", "D", "E")
    assert region.column_headers == (
        "Rank",
        "Record ID",
        "Value",
        "Value 2",
        "Value 3",
    )
    assert region.identity_columns == ("B",)
    assert region.ordinal_columns == ("A",)
    assert region.data_row_count == 6000
    assert region.projected_avoided_mismatches == 497_500


def test_region_draft_from_v1_item_preserves_bounded_compatibility_fields() -> None:
    region = region_draft_from_item(_V1_ITEM)

    assert region is not None
    assert region.current_range == ""
    assert region.available_columns == ()
    assert region.legacy_detail == "bounded legacy detail"
    assert region.identity_columns == ("B",)
    assert region.ordinal_columns == ("A",)


def test_region_draft_rejects_non_ranked_or_incomplete_items() -> None:
    assert region_draft_from_item({"sheet": "Config", "cell": "B2"}) is None
    assert region_draft_from_item({"sheet": "", "cell": "A1"}) is None
    assert region_draft_from_item({}) is None


def test_ranked_regions_from_action_filters_noise_and_preserves_order() -> None:
    second = dict(_V2_ITEM, member_id="ops", sheet="Panel2")

    regions = ranked_regions_from_action(
        {"items": [_V2_ITEM, "invalid", {"sheet": "Config", "cell": "B2"}, second]}
    )

    assert [(region.member_id, region.sheet) for region in regions] == [
        ("primary", "Panel"),
        ("ops", "Panel2"),
    ]


def test_ranked_regions_from_action_handles_missing_items() -> None:
    assert ranked_regions_from_action({}) == ()
    assert ranked_regions_from_action({"items": "invalid"}) == ()
