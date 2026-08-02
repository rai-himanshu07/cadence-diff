"""Validated comparison scope applied after full loading but before budgets."""

from __future__ import annotations

from collections.abc import Iterable
from typing import Self

from pydantic import BaseModel, field_validator

from qc_tool.coverage import CoverageItem, CoverageState
from qc_tool.findings import Finding
from qc_tool.io.model import WorkbookSnapshot
from qc_tool.ppt.model import DeckSnapshot


class ComparisonScope(BaseModel):
    """Optional analyst selection; ``None`` means all loaded content."""

    excel_sheets: tuple[str, ...] | None = None
    ppt_slide_indices: tuple[int, ...] | None = None

    @field_validator("excel_sheets", mode="before")
    @classmethod
    def validate_sheet_selection(cls, value: object) -> object:
        if value is None:
            return None
        values = tuple(str(item).strip() for item in value)  # type: ignore[arg-type]
        if not values or any(not item for item in values):
            raise ValueError("Excel scope cannot be empty")
        return tuple(dict.fromkeys(values))

    @field_validator("ppt_slide_indices", mode="before")
    @classmethod
    def validate_slide_selection(cls, value: object) -> object:
        if value is None:
            return None
        values = tuple(int(item) for item in value)  # type: ignore[arg-type]
        if not values:
            raise ValueError("PowerPoint scope cannot be empty")
        if any(item < 1 for item in values):
            raise ValueError("PowerPoint slide indices must be positive and 1-based")
        return tuple(dict.fromkeys(values))

    def validate_loaded(
        self,
        *,
        workbooks: Iterable[WorkbookSnapshot] = (),
        decks: Iterable[DeckSnapshot] = (),
    ) -> Self:
        sheet_names = {sheet.name for workbook in workbooks for sheet in workbook.sheets}
        slide_indices = {slide.index + 1 for deck in decks for slide in deck.slides}
        if self.excel_sheets is not None:
            unknown = sorted(set(self.excel_sheets) - sheet_names)
            if unknown:
                raise ValueError(
                    "unknown Excel sheet scope: "
                    + ", ".join(repr(item) for item in unknown)
                )
        if self.ppt_slide_indices is not None:
            unknown_slides = sorted(set(self.ppt_slide_indices) - slide_indices)
            if unknown_slides:
                raise ValueError(
                    "unknown PowerPoint slide scope: "
                    + ", ".join(str(item) for item in unknown_slides)
                )
        return self

    def filter_findings(self, findings: list[Finding]) -> list[Finding]:
        selected_sheets = (
            set(self.excel_sheets) if self.excel_sheets is not None else None
        )
        selected_slides = (
            set(self.ppt_slide_indices)
            if self.ppt_slide_indices is not None
            else None
        )
        return [
            finding
            for finding in findings
            if not (
                selected_sheets is not None
                and finding.artifact == "excel"
                and finding.sheet is not None
                and finding.sheet not in selected_sheets
            )
            and not (
                selected_slides is not None
                and finding.artifact == "ppt"
                and finding.slide_index is not None
                and finding.slide_index not in selected_slides
            )
        ]

    def coverage_item(
        self,
        *,
        workbooks: Iterable[WorkbookSnapshot] = (),
        decks: Iterable[DeckSnapshot] = (),
    ) -> CoverageItem:
        sheet_names = {sheet.name for workbook in workbooks for sheet in workbook.sheets}
        slide_indices = {slide.index + 1 for deck in decks for slide in deck.slides}
        parts: list[str] = []
        if sheet_names:
            parts.append(
                f"Excel {len(self.excel_sheets or sheet_names)}/{len(sheet_names)}"
            )
        if slide_indices:
            parts.append(
                "PowerPoint "
                f"{len(self.ppt_slide_indices or slide_indices)}/{len(slide_indices)}"
            )
        parts.append("files loaded fully")
        return CoverageItem(
            check_id="comparison-scope",
            label="Validated comparison scope",
            artifact="run",
            state=CoverageState.CHECKED,
            detail="; ".join(parts),
        )

    def disclosure(self) -> str | None:
        parts: list[str] = []
        if self.excel_sheets is not None:
            parts.append("Excel sheets: " + ", ".join(self.excel_sheets))
        if self.ppt_slide_indices is not None:
            parts.append(
                "PowerPoint slides: "
                + ", ".join(str(item) for item in self.ppt_slide_indices)
            )
        if not parts:
            return None
        return (
            "Comparison scope narrowed by analyst - "
            + "; ".join(parts)
            + ". Findings outside the scope are not reported; files loaded fully."
        )
