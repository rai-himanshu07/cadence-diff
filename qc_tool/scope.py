"""Validated comparison scope applied after full loading but before budgets."""

from __future__ import annotations

import re
from collections.abc import Iterable, Mapping
from typing import Self

from pydantic import BaseModel, field_validator

from qc_tool.coverage import CoverageItem, CoverageState
from qc_tool.findings import Finding
from qc_tool.io.model import WorkbookSnapshot
from qc_tool.package import MEMBER_ID_PATTERN
from qc_tool.ppt.model import DeckSnapshot


class ComparisonScope(BaseModel):
    """Optional analyst selection; ``None`` means all loaded content."""

    excel_sheets: tuple[str, ...] | None = None
    excel_member_sheets: dict[str, tuple[str, ...]] | None = None
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

    @field_validator("excel_member_sheets", mode="before")
    @classmethod
    def validate_member_sheets(cls, value: object) -> object:
        if value is None:
            return None
        if not isinstance(value, dict):
            raise ValueError("excel_member_sheets must be a mapping")
        result: dict[str, tuple[str, ...]] = {}
        for key, sheets in value.items():
            if not isinstance(key, str) or re.fullmatch(MEMBER_ID_PATTERN, key) is None:
                raise ValueError(f"invalid member id {key!r}")
            vals = tuple(str(item).strip() for item in sheets)
            if not vals or any(not item for item in vals):
                raise ValueError("member sheet selection cannot be empty")
            result[key] = tuple(dict.fromkeys(vals))
        return result

    def validate_loaded(
        self,
        *,
        workbooks: Iterable[WorkbookSnapshot] = (),
        decks: Iterable[DeckSnapshot] = (),
        workbooks_by_member: Mapping[str, WorkbookSnapshot] | None = None,
    ) -> Self:
        loaded_workbooks = tuple(workbooks)
        workbook_map = dict(workbooks_by_member or {})
        if not workbook_map and loaded_workbooks:
            if len(loaded_workbooks) > 1 and self.excel_member_sheets is not None:
                raise ValueError(
                    "member-qualified scope needs member-qualified workbooks"
                )
            if len(loaded_workbooks) == 1:
                workbook_map["primary"] = loaded_workbooks[0]
        sheet_names = {
            sheet.name for workbook in loaded_workbooks for sheet in workbook.sheets
        }
        slide_indices = {slide.index + 1 for deck in decks for slide in deck.slides}
        if workbook_map:
            if len(workbook_map) > 1 and self.excel_sheets is not None:
                raise ValueError(
                    "legacy Excel sheet scope is ambiguous with multiple members"
                )
            primary = workbook_map.get("primary")
            sheet_names = (
                {sheet.name for sheet in primary.sheets}
                if primary is not None
                else set()
            )
        if self.excel_sheets is not None:
            unknown = sorted(set(self.excel_sheets) - sheet_names)
            if unknown:
                raise ValueError(
                    "unknown Excel sheet scope: "
                    + ", ".join(repr(item) for item in unknown)
                )
        if self.excel_member_sheets is not None:
            unknown_members = sorted(set(self.excel_member_sheets) - set(workbook_map))
            if unknown_members:
                raise ValueError(
                    "unknown workbook members in scope: "
                    + ", ".join(unknown_members)
                )
            for member_id, sheets in self.excel_member_sheets.items():
                workbook = workbook_map[member_id]
                unknown = sorted(
                    set(sheets) - {sheet.name for sheet in workbook.sheets}
                )
                if unknown:
                    raise ValueError(
                        f"unknown Excel sheet scope for member {member_id}: "
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
        member_sheets = self.excel_member_sheets
        result: list[Finding] = []
        for finding in findings:
            if finding.artifact == "excel":
                member = getattr(finding, "artifact_member", "primary")
                if member_sheets is not None:
                    if member in member_sheets:
                        allowed = set(member_sheets[member])
                        if finding.sheet is not None and finding.sheet not in allowed:
                            continue
                elif selected_sheets is not None:
                    # legacy scope applies only to primary projection
                    if member != "primary":
                        continue
                    if (
                        finding.sheet is not None
                        and finding.sheet not in selected_sheets
                    ):
                        continue
            if (
                selected_slides is not None
                and finding.artifact == "ppt"
                and finding.slide_index is not None
                and finding.slide_index not in selected_slides
            ):
                continue
            result.append(finding)
        return result

    def coverage_item(
        self,
        *,
        workbooks: Iterable[WorkbookSnapshot] = (),
        decks: Iterable[DeckSnapshot] = (),
        workbooks_by_member: Mapping[str, WorkbookSnapshot] | None = None,
    ) -> CoverageItem:
        workbook_map = dict(workbooks_by_member or {})
        sheet_names = {sheet.name for workbook in workbooks for sheet in workbook.sheets}
        slide_indices = {slide.index + 1 for deck in decks for slide in deck.slides}
        parts: list[str] = []
        if workbook_map:
            selected = {
                member_id: (
                    self.excel_member_sheets[member_id]
                    if self.excel_member_sheets is not None
                    and member_id in self.excel_member_sheets
                    else tuple(sheet.name for sheet in workbook.sheets)
                )
                for member_id, workbook in workbook_map.items()
            }
            selected_count = sum(len(sheets) for sheets in selected.values())
            total_count = sum(
                len(workbook.sheets) for workbook in workbook_map.values()
            )
            parts.append(
                f"Excel {selected_count}/{total_count} sheets across "
                f"{len(selected)}/{len(workbook_map)} members"
            )
            if self.excel_member_sheets is not None:
                unlisted = len(set(workbook_map) - set(self.excel_member_sheets))
                if unlisted:
                    parts.append(f"{unlisted} unlisted member(s) unrestricted")
        elif sheet_names:
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
        if self.excel_member_sheets is not None:
            parts.extend(
                f"Excel member {member}: " + ", ".join(sheets)
                for member, sheets in sorted(self.excel_member_sheets.items())
            )
            parts.append("unlisted Excel members use all sheets")
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
            + ". Excel/PPT findings outside the selected scope are not reported; "
            "package analyses remain whole-package; files loaded fully."
        )
