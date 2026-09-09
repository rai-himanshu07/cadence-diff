"""Pure view model and validation for the ranked-table review dialog.

No NiceGUI import: every function here is directly unit-testable without a
browser. ``qc_tool.ui.app`` renders this view model and calls its mutation
and validation helpers from event handlers (plan-20260909, Step 9,
Criteria 12-15). The dialog reviews one ``RunBlockedError``'s
``row_identity_confirmation_required`` action: one or more unconfigured
ranked/sorted-table regions the analyst may confirm a row identity for.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from typing import Literal

from qc_tool.config.profile import (
    DeliverableProfile,
    ExcelMemberProfile,
    RowIdentityRule,
    SheetProfile,
)

DuplicatePolicy = Literal["skip", "occurrence", "position"]
DUPLICATE_POLICIES: tuple[DuplicatePolicy, ...] = ("skip", "occurrence", "position")

#: (label, consequence, risk level) for every duplicate-identity policy --
#: Criterion 12: all three visible with their exact consequences and a
#: clear risk level, "skip" recommended.
DUPLICATE_POLICY_COPY: dict[DuplicatePolicy, tuple[str, str, str]] = {
    "skip": (
        "Skip ambiguous duplicates",
        "Rows sharing a duplicate identity are left unmatched and keep "
        "comparing by position; nothing is silently guessed.",
        "recommended",
    ),
    "occurrence": (
        "Pair duplicates by occurrence",
        "The 1st, 2nd, 3rd... occurrence of a duplicate identity in the "
        "baseline pairs with the same occurrence number in the current file.",
        "moderate risk",
    ),
    "position": (
        "Pair duplicate leftovers by position",
        "Any duplicate identity left unpaired after unique matching falls "
        "back to raw row position, the same as before confirming a match.",
        "higher risk",
    ),
}


def _str_tuple(value: object) -> tuple[str, ...]:
    return tuple(str(entry) for entry in value) if isinstance(value, list | tuple) else ()


def _as_int(value: object) -> int | None:
    return value if isinstance(value, int) else None


def _as_float(value: object) -> float | None:
    return float(value) if isinstance(value, int | float) else None


@dataclass(frozen=True, slots=True)
class RegionDraft:
    """One editable ranked-table region shown in the dialog.

    ``data_row_count`` is ``None`` for a legacy v1 action item (no typed
    ``ranked_table_evidence``) -- the bounded compatibility path required by
    Criterion 11: the region still renders and can still be confirmed, just
    without the rich metrics or chip-selectable ``available_columns`` a v2
    item provides.
    """

    member_id: str
    sheet: str
    anchor_cell: str
    current_range: str = ""
    available_columns: tuple[str, ...] = ()
    data_row_count: int | None = None
    non_blank_coverage: float | None = None
    unique_ratio: float | None = None
    key_overlap: float | None = None
    displaced_ratio: float | None = None
    mismatch_reduction: float | None = None
    projected_positional_mismatches: int | None = None
    projected_avoided_mismatches: int | None = None
    legacy_detail: str = ""
    identity_columns: tuple[str, ...] = ()
    ordinal_columns: tuple[str, ...] = ()
    duplicate_policy: DuplicatePolicy = "skip"

    @property
    def has_typed_evidence(self) -> bool:
        return self.data_row_count is not None

    @property
    def label(self) -> str:
        """Member-qualified label (Criterion 15): ``member / sheet / range``."""
        prefix = f"{self.member_id} / " if self.member_id != "primary" else ""
        if self.current_range:
            return f"{prefix}{self.sheet} / {self.current_range}"
        return f"{prefix}{self.sheet}"

    @property
    def noise_summary(self) -> str:
        """Concise estimated-noise summary for the primary view (Criterion 12)."""
        if (
            self.projected_positional_mismatches is None
            or self.projected_avoided_mismatches is None
        ):
            return self.legacy_detail
        return (
            "Comparing by row position alone projects about "
            f"{self.projected_positional_mismatches:,} extra differences here; "
            "confirming a match could avoid roughly "
            f"{self.projected_avoided_mismatches:,} of them."
        )

    @property
    def why_paused_detail(self) -> str:
        """Human-readable evidence for the "Why QC paused" expansion.

        Plain language, never the raw snake_case telemetry sentence
        (Criterion 12).
        """
        if (
            self.data_row_count is None
            or self.non_blank_coverage is None
            or self.unique_ratio is None
            or self.key_overlap is None
            or self.displaced_ratio is None
            or self.mismatch_reduction is None
        ):
            return self.legacy_detail
        return (
            f"{self.data_row_count:,} data rows. Non-blank coverage "
            f"{self.non_blank_coverage:.0%}, values unique "
            f"{self.unique_ratio:.0%} of the time, {self.key_overlap:.0%} "
            "of identities appear in both files, and "
            f"{self.displaced_ratio:.0%} of matched rows changed position "
            "between the two files. Confirming this identity is estimated "
            f"to remove {self.mismatch_reduction:.0%} of that positional noise."
        )

    def with_columns(
        self, *, identity: tuple[str, ...], ordinal: tuple[str, ...]
    ) -> RegionDraft:
        return replace(self, identity_columns=identity, ordinal_columns=ordinal)

    def with_duplicate_policy(self, policy: DuplicatePolicy) -> RegionDraft:
        return replace(self, duplicate_policy=policy)

    def errors(self) -> tuple[str, ...]:
        """Validation errors for this region (Criterion 13)."""
        issues: list[str] = []
        identity = set(self.identity_columns)
        ordinal = set(self.ordinal_columns)
        if not identity:
            issues.append("Choose at least one column to match rows by.")
        overlap = identity & ordinal
        if overlap:
            issues.append(
                "A column cannot both match identity and be ignored as "
                "order-only: " + ", ".join(sorted(overlap))
            )
        if self.available_columns:
            available = set(self.available_columns)
            outside = (identity | ordinal) - available
            if outside:
                issues.append(
                    "Outside this region's columns: " + ", ".join(sorted(outside))
                )
        return tuple(issues)

    @property
    def is_valid(self) -> bool:
        return not self.errors()

    @property
    def status(self) -> Literal["ready", "error"]:
        return "ready" if self.is_valid else "error"


def region_draft_from_item(item: dict[str, object]) -> RegionDraft | None:
    """Build one region draft from a stored/IPC blocked-action item dict.

    Returns ``None`` when the item does not describe a ranked-table region
    at all (no sheet/cell, or a v1 item with no suggested identity columns)
    -- not every ``RunActionItem`` is one; the comparison-prerequisite
    -mismatch reason never is.
    """
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
            data_row_count=_as_int(evidence.get("data_row_count")),
            non_blank_coverage=_as_float(evidence.get("non_blank_coverage")),
            unique_ratio=_as_float(evidence.get("unique_ratio")),
            key_overlap=_as_float(evidence.get("key_overlap")),
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


@dataclass(frozen=True, slots=True)
class DialogViewModel:
    """Pure state for the ranked-table review dialog (Criteria 12-15).

    ``opened_hash``/``opened_source_existed`` snapshot the destination
    profile's on-disk state exactly once, at dialog-open time, for whatever
    name pre-fills the destination field (``opened_profile_name``) -- so a
    change made elsewhere while this dialog sits open is never silently
    overwritten (Criterion 13). A destination the analyst has retyped to a
    different name was never snapshotted, so it never conflicts.
    """

    regions: tuple[RegionDraft, ...]
    active_index: int = 0
    profile_name: str = ""
    opened_profile_name: str = ""
    opened_hash: str | None = None
    opened_source_existed: bool = False

    @property
    def region_count(self) -> int:
        return len(self.regions)

    @property
    def active_region(self) -> RegionDraft:
        return self.regions[self.active_index]

    def with_active_index(self, index: int) -> DialogViewModel:
        clamped = max(0, min(index, len(self.regions) - 1))
        return replace(self, active_index=clamped)

    def with_region(self, index: int, region: RegionDraft) -> DialogViewModel:
        regions = list(self.regions)
        regions[index] = region
        return replace(self, regions=tuple(regions))

    def with_active_region(self, region: RegionDraft) -> DialogViewModel:
        return self.with_region(self.active_index, region)

    def with_profile_name(self, name: str) -> DialogViewModel:
        return replace(self, profile_name=name)

    def destination_errors(self) -> tuple[str, ...]:
        """Validation errors for the profile destination (Criterion 13)."""
        name = self.profile_name.strip()
        if not name:
            return ("Enter a profile name to save this match to.",)
        if name == "default":
            return ("Choose a named profile; the built-in default is immutable.",)
        return ()

    @property
    def is_creating_profile(self) -> bool:
        """Create-versus-update wording (Criterion 12)."""
        return self.profile_name.strip() != self.opened_profile_name

    @property
    def is_valid(self) -> bool:
        return not self.destination_errors() and all(
            region.is_valid for region in self.regions
        )

    @property
    def first_invalid_index(self) -> int | None:
        for index, region in enumerate(self.regions):
            if not region.is_valid:
                return index
        return None

    @property
    def initial_focus_target(self) -> Literal["profile_name", "region_identity"]:
        """Criterion 14: focus the new-profile control only when the run's
        profile was the immutable default (so the destination starts
        blank); otherwise focus the first region's identity control."""
        return "profile_name" if not self.opened_profile_name else "region_identity"

    def has_profile_conflict(
        self, *, current_hash: str | None, current_exists: bool
    ) -> bool:
        """True when the destination profile changed since the dialog opened."""
        if self.profile_name.strip() != self.opened_profile_name:
            return False
        if current_exists != self.opened_source_existed:
            return True
        return current_exists and current_hash != self.opened_hash


def view_model_from_action(
    action: dict[str, object],
    *,
    source_profile: str,
    opened_hash: str | None = None,
    opened_source_existed: bool = False,
) -> DialogViewModel | None:
    """Build the dialog's initial state from a stored blocked-action dict.

    Returns ``None`` when the action carries no usable ranked-table region
    (the caller should fall back to the generic blocked-run notification --
    this is always true for ``comparison_prerequisite_mismatch``).
    """
    raw_items = action.get("items")
    items = (
        [item for item in raw_items if isinstance(item, dict)]
        if isinstance(raw_items, list)
        else []
    )
    regions = tuple(
        draft for item in items if (draft := region_draft_from_item(item)) is not None
    )
    if not regions:
        return None
    opened_profile_name = "" if source_profile == "default" else source_profile
    return DialogViewModel(
        regions=regions,
        profile_name=opened_profile_name,
        opened_profile_name=opened_profile_name,
        opened_hash=opened_hash,
        opened_source_existed=opened_source_existed,
    )


def region_to_row_identity_rule(region: RegionDraft) -> RowIdentityRule:
    return RowIdentityRule(
        anchor_cell=region.anchor_cell,
        identity_columns=list(region.identity_columns),
        ordinal_columns=list(region.ordinal_columns),
        duplicate_policy=region.duplicate_policy,
    )


def with_row_identity_rule(
    profile: DeliverableProfile,
    *,
    member_id: str,
    workbook_count: int,
    sheet: str,
    rule: RowIdentityRule,
) -> DeliverableProfile:
    """Return a copy with one anchor-keyed rule in the correct member scope."""
    updated = profile.model_copy(deep=True)

    def upsert(sheets: dict[str, SheetProfile]) -> None:
        sheet_profile = sheets.get(sheet, SheetProfile()).model_copy(deep=True)
        sheet_profile.row_identity_rules = [
            existing
            for existing in sheet_profile.row_identity_rules
            if existing.anchor_cell != rule.anchor_cell
        ]
        sheet_profile.row_identity_rules.append(rule)
        sheets[sheet] = sheet_profile

    if workbook_count == 1 and member_id == "primary":
        upsert(updated.excel.sheets)
        return updated
    member = updated.excel.members.get(member_id, ExcelMemberProfile()).model_copy(
        deep=True
    )
    upsert(member.sheets)
    updated.excel.members[member_id] = member
    return updated


def apply_view_model(
    profile: DeliverableProfile,
    view_model: DialogViewModel,
    *,
    workbook_count: int,
) -> DeliverableProfile:
    """Return a copy of ``profile`` with every region's confirmed rule upserted."""
    updated = profile
    for region in view_model.regions:
        updated = with_row_identity_rule(
            updated,
            member_id=region.member_id,
            workbook_count=workbook_count,
            sheet=region.sheet,
            rule=region_to_row_identity_rule(region),
        )
    return updated
