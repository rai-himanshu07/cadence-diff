"""Parse ranked-table blocked actions into configuration-workspace proposals."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

DuplicatePolicy = Literal["skip", "occurrence", "position"]


def _str_tuple(value: object) -> tuple[str, ...]:
    return tuple(str(entry) for entry in value) if isinstance(value, list | tuple) else ()


def _as_int(value: object) -> int | None:
    return value if isinstance(value, int) else None


def _as_float(value: object) -> float | None:
    return float(value) if isinstance(value, int | float) else None


@dataclass(frozen=True, slots=True)
class RegionDraft:
    """One bounded ranked-region proposal decoded from a blocked run action."""

    member_id: str
    sheet: str
    anchor_cell: str
    current_range: str = ""
    available_columns: tuple[str, ...] = ()
    column_headers: tuple[str, ...] = ()
    header_row: int | None = None
    manual_review: bool = False
    data_row_count: int | None = None
    non_blank_coverage: float | None = None
    unique_ratio: float | None = None
    key_overlap: float | None = None
    formula_ratio: float | None = None
    displaced_ratio: float | None = None
    mismatch_reduction: float | None = None
    projected_positional_mismatches: int | None = None
    projected_avoided_mismatches: int | None = None
    legacy_detail: str = ""
    identity_columns: tuple[str, ...] = ()
    ordinal_columns: tuple[str, ...] = ()
    duplicate_policy: DuplicatePolicy = "skip"


def region_draft_from_item(item: dict[str, object]) -> RegionDraft | None:
    """Decode one ranked region from a versioned or legacy action item."""
    sheet = item.get("sheet")
    cell = item.get("cell")
    if not isinstance(sheet, str) or not sheet or not isinstance(cell, str) or not cell:
        return None
    raw_member = item.get("member_id")
    member_id = raw_member if isinstance(raw_member, str) and raw_member else "primary"
    evidence = item.get("ranked_table_evidence")
    if isinstance(evidence, dict):
        return RegionDraft(
            member_id=member_id,
            sheet=sheet,
            anchor_cell=cell,
            current_range=str(evidence.get("current_range") or ""),
            available_columns=_str_tuple(evidence.get("available_columns")),
            column_headers=_str_tuple(evidence.get("column_headers")),
            header_row=_as_int(evidence.get("header_row")),
            manual_review=bool(evidence.get("manual_review", False)),
            data_row_count=_as_int(evidence.get("data_row_count")),
            non_blank_coverage=_as_float(evidence.get("non_blank_coverage")),
            unique_ratio=_as_float(evidence.get("unique_ratio")),
            key_overlap=_as_float(evidence.get("key_overlap")),
            formula_ratio=_as_float(evidence.get("formula_ratio")),
            displaced_ratio=_as_float(evidence.get("displaced_ratio")),
            mismatch_reduction=_as_float(evidence.get("mismatch_reduction")),
            projected_positional_mismatches=_as_int(
                evidence.get("projected_positional_mismatches")
            ),
            projected_avoided_mismatches=_as_int(
                evidence.get("projected_avoided_mismatches")
            ),
            identity_columns=_str_tuple(evidence.get("suggested_identity_columns")),
            ordinal_columns=_str_tuple(evidence.get("suggested_ordinal_columns")),
        )
    identity = _str_tuple(item.get("suggested_identity_columns"))
    if not identity:
        return None
    return RegionDraft(
        member_id=member_id,
        sheet=sheet,
        anchor_cell=cell,
        legacy_detail=str(item.get("detail") or ""),
        identity_columns=identity,
        ordinal_columns=_str_tuple(item.get("suggested_ordinal_columns")),
    )


def ranked_regions_from_action(
    action: dict[str, object],
) -> tuple[RegionDraft, ...]:
    """Return every usable ranked-region proposal in one blocked action."""
    raw_items = action.get("items")
    items = (
        [item for item in raw_items if isinstance(item, dict)]
        if isinstance(raw_items, list)
        else []
    )
    return tuple(
        proposal
        for item in items
        if (proposal := region_draft_from_item(item)) is not None
    )
