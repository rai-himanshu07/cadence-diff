"""Primitive-only "run cannot proceed" contract shared by the engine, worker,
history, queue, UI, and CLI surfaces.

A blocked run is a terminal, non-active outcome distinct from success,
failure, or cancellation: something about the supplied inputs (not a QC
finding) means the comparison cannot proceed until the analyst takes an
action outside this tool. No formula or path enters this payload. Ranked-table
actions may carry bounded header labels for local analyst confirmation; all
other fields are structural labels, locations, and aggregate counts.
"""

from __future__ import annotations

from enum import StrEnum
from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator, model_validator

from qc_tool.package import MEMBER_ID_PATTERN

MAX_RUN_ACTION_ITEMS = 16
MAX_RUN_ACTION_SHEET_CHARS = 128
MAX_RUN_ACTION_CELL_CHARS = 32
MAX_RUN_ACTION_LABEL_CHARS = 256
MAX_RUN_ACTION_DETAIL_CHARS = 1_024
MAX_RUN_ACTION_MESSAGE_CHARS = 1_024
MAX_RANKED_TABLE_RANGE_CHARS = 64
MAX_RANKED_TABLE_AVAILABLE_COLUMNS = 512
MAX_RANKED_TABLE_HEADER_CHARS = 80


class RunActionReason(StrEnum):
    COMPARISON_PREREQUISITE_MISMATCH = "comparison_prerequisite_mismatch"
    ROW_IDENTITY_CONFIRMATION_REQUIRED = "row_identity_confirmation_required"
    BLANK_IDENTITY_KEY_BLOCKED = "blank_identity_key_blocked"


class RankedTableEvidence(BaseModel):
    """Typed, value-free evidence for one ranked/sorted-table suggestion.

    Most fields are bounded structural labels or aggregate numbers. The one
    deliberate content exception is ``column_headers``: short header labels
    shown only in the local confirmation dialog so an analyst need not reopen
    the workbook. Ordinary row values and raw telemetry never enter this
    payload. Field names
    intentionally mirror ``qc_tool.excel.ranked_identity.RankedTableCandidate``
    so the producer can populate this by direct attribute copy.
    """

    version: Literal[2] = 2
    member_id: str = Field(default="primary", pattern=MEMBER_ID_PATTERN)
    sheet: str = ""
    current_range: str = ""
    data_row_count: int = Field(default=0, ge=0)
    header_row: int | None = Field(default=None, ge=1)
    manual_review: bool = False
    available_columns: tuple[str, ...] = Field(
        default=(), max_length=MAX_RANKED_TABLE_AVAILABLE_COLUMNS
    )
    column_headers: tuple[str, ...] = Field(
        default=(), max_length=MAX_RANKED_TABLE_AVAILABLE_COLUMNS
    )
    suggested_identity_columns: tuple[str, ...] = Field(default=(), max_length=12)
    suggested_ordinal_columns: tuple[str, ...] = Field(default=(), max_length=12)
    non_blank_coverage: float = Field(default=0.0, ge=0.0, le=1.0)
    unique_ratio: float = Field(default=0.0, ge=0.0, le=1.0)
    key_overlap: float = Field(default=0.0, ge=0.0, le=1.0)
    formula_ratio: float = Field(default=0.0, ge=0.0, le=1.0)
    displaced_ratio: float = Field(default=0.0, ge=0.0, le=1.0)
    mismatch_reduction: float = Field(default=0.0, ge=0.0, le=1.0)
    projected_positional_mismatches: int = Field(default=0, ge=0)
    projected_avoided_mismatches: int = Field(default=0, ge=0)

    @field_validator("sheet", "current_range", mode="before")
    @classmethod
    def bound_text(cls, value: object, info: Any) -> str:
        limits = {
            "sheet": MAX_RUN_ACTION_SHEET_CHARS,
            "current_range": MAX_RANKED_TABLE_RANGE_CHARS,
        }
        return str(value or "")[: limits[info.field_name]]

    @field_validator("available_columns", mode="before")
    @classmethod
    def bound_available_columns(cls, value: object) -> object:
        if isinstance(value, list | tuple):
            return tuple(value)[:MAX_RANKED_TABLE_AVAILABLE_COLUMNS]
        return value

    @field_validator("column_headers", mode="before")
    @classmethod
    def bound_column_headers(cls, value: object) -> object:
        if not isinstance(value, list | tuple):
            return value
        return tuple(
            " ".join(str(entry or "").split())[:MAX_RANKED_TABLE_HEADER_CHARS]
            for entry in value[:MAX_RANKED_TABLE_AVAILABLE_COLUMNS]
        )


class RunActionItem(BaseModel):
    """One bounded location; ranked evidence may include short headers."""

    member_id: str = Field(default="primary", pattern=MEMBER_ID_PATTERN)
    sheet: str = ""
    cell: str = ""
    label: str = ""
    detail: str = ""
    suggested_identity_columns: tuple[str, ...] = Field(default=(), max_length=12)
    suggested_ordinal_columns: tuple[str, ...] = Field(default=(), max_length=12)
    #: Typed v2 evidence (Criterion 11); ``None`` for the comparison-
    #: prerequisite reason and for any legacy stored payload recorded
    #: before this field existed -- consumers fall back to the flat
    #: ``detail``/``suggested_*_columns`` fields above in that case.
    ranked_table_evidence: RankedTableEvidence | None = None

    @field_validator("sheet", "cell", "label", "detail", mode="before")
    @classmethod
    def bound_text(cls, value: object, info: Any) -> str:
        limits = {
            "sheet": MAX_RUN_ACTION_SHEET_CHARS,
            "cell": MAX_RUN_ACTION_CELL_CHARS,
            "label": MAX_RUN_ACTION_LABEL_CHARS,
            "detail": MAX_RUN_ACTION_DETAIL_CHARS,
        }
        return str(value or "")[: limits[info.field_name]]


class RunActionRequired(BaseModel):
    """Versioned, bounded, primitive-only reason a run stopped before QC."""

    version: int = 1
    reason: RunActionReason
    items: list[RunActionItem] = Field(default_factory=list)
    message: str = ""
    omitted_items: int = Field(default=0, ge=0)

    @model_validator(mode="before")
    @classmethod
    def bound_payload(cls, value: object) -> object:
        if not isinstance(value, dict):
            return value
        payload = dict(value)
        raw_items = payload.get("items")
        items = list(raw_items) if isinstance(raw_items, list | tuple) else []
        payload["items"] = items[:MAX_RUN_ACTION_ITEMS]
        omitted = len(items) - len(payload["items"])
        raw_omitted = payload.get("omitted_items", 0)
        payload["omitted_items"] = max(
            int(raw_omitted) if isinstance(raw_omitted, int) else 0,
            omitted,
        )
        payload["message"] = str(payload.get("message") or "")[
            :MAX_RUN_ACTION_MESSAGE_CHARS
        ]
        return payload


class RunBlockedError(RuntimeError):
    """The run cannot proceed; carries the bounded action-required payload."""

    def __init__(self, action_required: RunActionRequired) -> None:
        super().__init__(action_required.message or action_required.reason.value)
        self.action_required = action_required
