"""Forward-reference fixture for the ranked-table run-action v2 payload
(plan-20260909-release-hardening-and-ranked-table-review.md, Criterion 11).

`RankedTableEvidence` does not exist as production code yet -- Step 8 adds
it. This module fixes the exact field shape now so Step 8's real model and
Step 9's dialog view-model are built against one agreed contract instead of
re-deriving it from plan prose later.

Field names intentionally match `RankedTableCandidate`'s own attributes
(`qc_tool/excel/ranked_identity.py`) so the real model can be populated by
direct attribute copy. Every field is either a bounded structural label
(member id, sheet, region range, column letters) or a plain aggregate number
-- never a cell value, header string, or raw telemetry sentence.
"""

from __future__ import annotations

from typing import TypedDict


class RankedTableEvidencePayloadV2(TypedDict):
    version: int
    member_id: str
    sheet: str
    current_range: str
    data_row_count: int
    available_columns: tuple[str, ...]
    suggested_identity_columns: tuple[str, ...]
    suggested_ordinal_columns: tuple[str, ...]
    non_blank_coverage: float
    unique_ratio: float
    key_overlap: float
    formula_ratio: float
    displaced_ratio: float
    mismatch_reduction: float
    projected_positional_mismatches: int
    projected_avoided_mismatches: int


def ranked_table_evidence_payload_v2(
    *,
    member_id: str = "primary",
    sheet: str = "Data",
    current_range: str = "A1:E6000",
    data_row_count: int = 5999,
    available_columns: tuple[str, ...] = ("A", "B", "C", "D", "E"),
    suggested_identity_columns: tuple[str, ...] = ("B",),
    suggested_ordinal_columns: tuple[str, ...] = ("A",),
    non_blank_coverage: float = 0.999,
    unique_ratio: float = 0.998,
    key_overlap: float = 0.95,
    formula_ratio: float = 0.0,
    displaced_ratio: float = 1.0,
    mismatch_reduction: float = 0.995,
    projected_positional_mismatches: int = 500_000,
    projected_avoided_mismatches: int | None = None,
) -> RankedTableEvidencePayloadV2:
    """One synthetic, value-free v2 payload for dialog/view-model fixtures.

    Every default describes a plausible ranked-table screen result on a
    bounded synthetic 6,000-row table -- never data read from a real
    workbook.
    """
    return RankedTableEvidencePayloadV2(
        version=2,
        member_id=member_id,
        sheet=sheet,
        current_range=current_range,
        data_row_count=data_row_count,
        available_columns=available_columns,
        suggested_identity_columns=suggested_identity_columns,
        suggested_ordinal_columns=suggested_ordinal_columns,
        non_blank_coverage=non_blank_coverage,
        unique_ratio=unique_ratio,
        key_overlap=key_overlap,
        formula_ratio=formula_ratio,
        displaced_ratio=displaced_ratio,
        mismatch_reduction=mismatch_reduction,
        projected_positional_mismatches=projected_positional_mismatches,
        projected_avoided_mismatches=(
            projected_avoided_mismatches
            if projected_avoided_mismatches is not None
            else round(projected_positional_mismatches * mismatch_reduction)
        ),
    )
