"""Pure view model, diff, and resolved-configuration builder for the
mode-aware configuration workspace (plan-20260913, Step 7).

No NiceGUI import: every function here is directly unit-testable without a
browser, mirroring the established ``qc_tool.ui.ranked_table_dialog``
convention. ``qc_tool.ui.config_workspace`` renders this view model and
calls its mutation/validation/builder helpers from event handlers.

Scope boundary (disclosed): this module builds the SHELL's resolution --
member/sheet pairing plus per-region mode/anchor/identity-column decisions
sufficient to drive execution through the bindings Step 3 already wired
(confirmed renames, keyed alignment). It does not yet implement every
region/column semantic (blank-key threshold policy, period bands, ordinal
suppression nuance, selector prerequisites) -- Step 8 ("Complete cycle
configuration semantics end to end") extends this builder and the engine
bindings together.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass, field, replace
from enum import StrEnum
from typing import Literal

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

from qc_tool.config.input_contract import (
    HeaderIntent,
    LogicalColumnContract,
    LogicalMemberContract,
    LogicalRegionContract,
    LogicalSheetContract,
    RegionMode,
    SelectorPrerequisiteContract,
    StructuralExclusionContract,
    WorkbookInputContract,
)
from qc_tool.config.profile import DeliverableProfile
from qc_tool.config.resolved_input import (
    ResolvedColumn,
    ResolvedInputConfigurationV1,
    ResolvedMember,
    ResolvedRegion,
    ResolvedSelector,
    ResolvedSheet,
)
from qc_tool.coverage import QCRunMode
from qc_tool.setup.models import MemberSetupProfile, SetupAnalysisResult, SheetSetupProfile

AlignmentRole = Literal["none", "identity", "ordinal"]

_SLUG_RE = re.compile(r"[^a-z0-9]+")


def slugify(text: str, *, prefix: str = "s") -> str:
    """A stable ``[a-z][a-z0-9_]{0,63}`` id from arbitrary display text."""
    lowered = _SLUG_RE.sub("_", text.lower()).strip("_")
    if not lowered or not lowered[0].isalpha():
        lowered = f"{prefix}_{lowered}" if lowered else prefix
    return lowered[:64]


def _unique_id(base: str, taken: set[str]) -> str:
    candidate = base
    suffix = 2
    while candidate in taken:
        candidate = f"{base[:60]}_{suffix}"
        suffix += 1
    taken.add(candidate)
    return candidate


@dataclass(frozen=True, slots=True)
class RegionDecision:
    """One region's analyst-facing configuration decision.

    Defaults mirror the architecture's own "Workbook/member and sheet
    defaults are inherited automatic" rule -- a freshly seeded decision is
    always ``mode="automatic"``, never pre-guessed as keyed/excluded.
    """

    region_id: str
    sheet: str
    current_range: str
    anchor_cell: str
    mode: RegionMode = "automatic"
    header_intent: HeaderIntent = "automatic"
    #: Only meaningful when ``header_intent == "first_data_row"``.
    first_data_row: int | None = None
    #: Explicit preamble/footer row COUNTS -- an alternative, additive way to
    #: express the same boundary as ``first_data_row`` (Step 8's "bottom-
    #: aligned preamble / top-aligned footer" criterion). Either or both may
    #: be set; they compose (the engine adapter treats ``first_data_row`` as
    #: authoritative over ``preamble_rows`` when both are present).
    preamble_rows: int = 0
    footer_rows: int = 0
    #: An analyst-entered override for the region's BASELINE-side range,
    #: independent of ``current_range`` (Step 8's "dynamic side-specific
    #: outer/data ranges" criterion). Empty means "use the auto-matched
    #: baseline region, if any."
    baseline_range: str = ""
    identity_columns: tuple[str, ...] = ()
    ordinal_columns: tuple[str, ...] = ()
    #: Columns excluded from the value-diff engine (requires a reason +
    #: expiry, like region-level exclusion). Scope boundary (disclosed,
    #: plan-20260913 Step 8): suppresses cached-value/number-format
    #: findings via the same ``ignore_ranges`` mechanism a legacy profile
    #: already uses; does NOT yet suppress a formula-text-changed finding
    #: for the same cell (the formula-diff engine has no comparable
    #: region-scoped ignore hook yet) -- an ignored formula cell may still
    #: surface a formula finding today.
    ignore_columns: tuple[str, ...] = ()
    #: Columns whose cached-value changes stay visible Expected rather than
    #: hidden (Step 8's "expected_refresh remains visible Expected" rule).
    expected_refresh_columns: tuple[str, ...] = ()
    #: Outer-whitespace trim for identity-column equality (Step 8's "exact
    #: typed equality plus optional outer-whitespace trim only" rule).
    #: Applies to every identity column in this region -- a deliberate
    #: region-level simplification of the contract's per-column field,
    #: disclosed since real-world identity columns in one region are
    #: almost always formatted consistently.
    trim_identity_whitespace: bool = False
    blank_key_policy: Literal["system_default", "tolerate", "block"] = "system_default"
    available_columns: tuple[str, ...] = ()
    duplicate_key_policy: Literal["skip", "occurrence", "position"] = "skip"
    exclusion_reason: str = ""
    exclusion_expires_on: str = ""  # ISO date, "" = not set
    ignore_columns_reason: str = ""
    ignore_columns_expires_on: str = ""
    #: A detected ranked-table candidate existed for this region but has not
    #: yet been reviewed -- surfaced as a warning until the analyst
    #: explicitly picks a mode (never auto-applied).
    ranked_candidate_pending: bool = False
    confirmed: bool = False

    @property
    def requires_exclusion_detail(self) -> bool:
        return self.mode == "excluded"

    @property
    def is_valid(self) -> bool:
        if self.mode == "excluded" and (not self.exclusion_reason or not self.exclusion_expires_on):
            return False
        if self.mode == "keyed" and not self.identity_columns:
            return False
        if self.header_intent == "first_data_row" and self.first_data_row is None:
            return False
        return not (
            self.ignore_columns
            and (not self.ignore_columns_reason or not self.ignore_columns_expires_on)
        )


@dataclass(frozen=True, slots=True)
class SelectorDecision:
    """One analyst-declared selector/scenario prerequisite cell (Step 8's
    "explicit selector prerequisites for dropdown/filter/scenario/parameter
    cells" criterion). Never carries a value -- only a label and a
    current-side cell location; the run-time engine check compares the two
    files' actual saved values without ever persisting either one.
    """

    selector_id: str
    label: str
    cell: str  # current-side A1 cell


@dataclass(frozen=True, slots=True)
class SheetReview:
    """One sheet's detected structure plus its region decisions."""

    sheet_name: str
    hidden: bool
    very_hidden: bool
    side: Literal["baseline", "current"]
    regions: tuple[RegionDecision, ...] = ()
    failure_detail: str = ""
    selectors: tuple[SelectorDecision, ...] = ()


@dataclass(frozen=True, slots=True)
class MemberReview:
    member_id: str
    current_sheets: tuple[SheetReview, ...] = ()
    baseline_sheets: tuple[SheetReview, ...] = ()
    failure_detail: str = ""
    #: current_sheet_name -> baseline_sheet_name for an analyst-declared
    #: rename (Step 8), promoting an added+removed pair into one logical
    #: sheet. Never guessed -- only ever set by an explicit UI choice.
    sheet_renames: dict[str, str] = field(default_factory=dict)


def _region_columns(min_col: int, max_col: int) -> tuple[str, ...]:
    return tuple(get_column_letter(col) for col in range(min_col, max_col + 1))


def region_decisions_from_sheet(
    sheet_profile: SheetSetupProfile, *, taken_ids: set[str]
) -> tuple[RegionDecision, ...]:
    """Seed one sheet's region decisions from its detected regions, every
    one defaulting to ``mode="automatic"``.
    """
    decisions: list[RegionDecision] = []
    for detected in sheet_profile.regions:
        region = detected.region
        region_id = _unique_id(
            slugify(f"{sheet_profile.sheet_name}_{region.cell_range}", prefix="region"),
            taken_ids,
        )
        decisions.append(
            RegionDecision(
                region_id=region_id,
                sheet=sheet_profile.sheet_name,
                current_range=region.cell_range,
                anchor_cell=f"{get_column_letter(region.min_col)}{region.min_row}",
                available_columns=_region_columns(region.min_col, region.max_col),
                ranked_candidate_pending=detected.ranked_candidate is not None,
            )
        )
    return tuple(decisions)


def member_review_from_scan(
    member_id: str, profile: MemberSetupProfile
) -> MemberReview:
    """Build one member's full review state from its raw scan payload,
    seeding every region decision at its ``automatic`` default.
    """
    taken_ids: set[str] = set()
    current_sheets = tuple(
        SheetReview(
            sheet_name=sheet.sheet_name,
            hidden=sheet.hidden,
            very_hidden=sheet.very_hidden,
            side="current",
            regions=region_decisions_from_sheet(sheet, taken_ids=taken_ids),
            failure_detail=sheet.failure_detail,
        )
        for sheet in profile.current_sheets
    )
    baseline_sheets = tuple(
        SheetReview(
            sheet_name=sheet.sheet_name,
            hidden=sheet.hidden,
            very_hidden=sheet.very_hidden,
            side="baseline",
            regions=region_decisions_from_sheet(sheet, taken_ids=taken_ids),
            failure_detail=sheet.failure_detail,
        )
        for sheet in profile.baseline_sheets
    )
    return MemberReview(
        member_id=member_id,
        current_sheets=current_sheets,
        baseline_sheets=baseline_sheets,
        failure_detail=profile.failure_detail,
    )


def sheet_pairing(
    member: MemberReview,
) -> tuple[tuple[tuple[str, str], ...], tuple[str, ...], tuple[str, ...]]:
    """``(same_name_pairs, added_current_only, removed_baseline_only)``.

    Same-name sheets pre-pair automatically (Criterion: "same-name sheets
    auto-pair"); anything else is surfaced as an explicit acknowledgement,
    never a guessed split/merge.
    """
    current_names = {sheet.sheet_name for sheet in member.current_sheets}
    baseline_names = {sheet.sheet_name for sheet in member.baseline_sheets}
    paired = tuple(sorted(current_names & baseline_names))
    added = tuple(sorted(current_names - baseline_names))
    removed = tuple(sorted(baseline_names - current_names))
    return tuple((name, name) for name in paired), added, removed


def set_sheet_rename(
    member: MemberReview, current_name: str, baseline_name: str | None
) -> MemberReview:
    """Declare (``baseline_name`` not None) or clear (``None``) that
    ``current_name`` is a rename of ``baseline_name`` (Step 8's "renamed
    sheets require explicit mapping" criterion). One-to-one by construction:
    claiming a baseline name here silently releases whichever OTHER current
    sheet previously claimed it.
    """
    renames = dict(member.sheet_renames)
    renames.pop(current_name, None)
    if baseline_name is not None:
        for other_current, other_baseline in list(renames.items()):
            if other_baseline == baseline_name:
                del renames[other_current]
        renames[current_name] = baseline_name
    return replace(member, sheet_renames=renames)


def effective_sheet_pairing(
    member: MemberReview,
) -> tuple[tuple[tuple[str, str], ...], tuple[str, ...], tuple[str, ...]]:
    """``sheet_pairing`` plus every analyst-declared rename promoted from
    added/removed into a real pair -- what the resolved configuration and
    the workspace's own acknowledgement warnings both use.
    """
    pairs, added, removed = sheet_pairing(member)
    renamed_pairs = tuple(
        (baseline_name, current_name)
        for current_name, baseline_name in sorted(member.sheet_renames.items())
        if current_name in added and baseline_name in removed
    )
    if not renamed_pairs:
        return pairs, added, removed
    renamed_currents = {current for _, current in renamed_pairs}
    renamed_baselines = {baseline for baseline, _ in renamed_pairs}
    return (
        tuple(sorted((*pairs, *renamed_pairs))),
        tuple(name for name in added if name not in renamed_currents),
        tuple(name for name in removed if name not in renamed_baselines),
    )


def add_selector(
    member: MemberReview, sheet_name: str, *, label: str, cell: str
) -> MemberReview:
    """Declare a new selector prerequisite on one current sheet (Step 8).
    Raises ``ValueError`` for an unparseable cell or a blank label --
    surfaced to the analyst, never silently ignored.
    """
    if not label.strip():
        raise ValueError("enter a label for this selector")
    parse_a1_cell(cell)  # validates; raises ValueError with a plain message
    taken_ids = {
        selector.selector_id
        for sheet in member.current_sheets
        for selector in sheet.selectors
    }
    selector_id = _unique_id(slugify(label, prefix="selector"), taken_ids)
    new_sheets = []
    for sheet in member.current_sheets:
        if sheet.sheet_name != sheet_name:
            new_sheets.append(sheet)
            continue
        new_sheets.append(
            replace(
                sheet,
                selectors=(
                    *sheet.selectors,
                    SelectorDecision(
                        selector_id=selector_id, label=label.strip(), cell=cell.upper()
                    ),
                ),
            )
        )
    return replace(member, current_sheets=tuple(new_sheets))


def remove_selector(member: MemberReview, sheet_name: str, selector_id: str) -> MemberReview:
    new_sheets = []
    for sheet in member.current_sheets:
        if sheet.sheet_name != sheet_name:
            new_sheets.append(sheet)
            continue
        new_sheets.append(
            replace(
                sheet,
                selectors=tuple(s for s in sheet.selectors if s.selector_id != selector_id),
            )
        )
    return replace(member, current_sheets=tuple(new_sheets))


_ColumnRole = Literal["identity", "ordinal", "ignore", "expected_refresh"]


def _set_column_role(
    region: RegionDecision, role: _ColumnRole, columns: tuple[str, ...]
) -> RegionDecision:
    """Assign the full set of columns holding one role, evicting every one
    of them from the other three roles (disjoint by construction -- Step
    8's "identity and ordinal columns are disjoint" criterion, generalized
    to all four column roles).
    """
    roles: dict[_ColumnRole, tuple[str, ...]] = {
        "identity": region.identity_columns,
        "ordinal": region.ordinal_columns,
        "ignore": region.ignore_columns,
        "expected_refresh": region.expected_refresh_columns,
    }
    roles[role] = columns
    for other_role in roles:
        if other_role != role:
            roles[other_role] = tuple(c for c in roles[other_role] if c not in columns)
    return replace(
        region,
        identity_columns=roles["identity"],
        ordinal_columns=roles["ordinal"],
        ignore_columns=roles["ignore"],
        expected_refresh_columns=roles["expected_refresh"],
    )


def set_identity_columns(region: RegionDecision, columns: tuple[str, ...]) -> RegionDecision:
    return _set_column_role(region, "identity", columns)


def set_ordinal_columns(region: RegionDecision, columns: tuple[str, ...]) -> RegionDecision:
    return _set_column_role(region, "ordinal", columns)


def set_ignore_columns(region: RegionDecision, columns: tuple[str, ...]) -> RegionDecision:
    return _set_column_role(region, "ignore", columns)


def set_expected_refresh_columns(
    region: RegionDecision, columns: tuple[str, ...]
) -> RegionDecision:
    return _set_column_role(region, "expected_refresh", columns)


def regions_overlap(a: RegionDecision, b: RegionDecision) -> bool:
    """Whether two regions' current-side ranges occupy any shared cell
    (Step 8's "overlapping regions are rejected" criterion).
    """
    try:
        a_min_row, a_min_col, a_max_row, a_max_col = parse_a1_range(a.current_range)
        b_min_row, b_min_col, b_max_row, b_max_col = parse_a1_range(b.current_range)
    except ValueError:
        return False
    return (
        a_min_col <= b_max_col
        and b_min_col <= a_max_col
        and a_min_row <= b_max_row
        and b_min_row <= a_max_row
    )


class WizardStep(StrEnum):
    ROLES = "roles"
    SCANNING = "scanning"
    PROFILE = "profile"
    REGIONS = "regions"
    REVIEW = "review"


#: Mode-aware navigation (Step 7 criterion): every mode shares one shell,
#: but final-package/preflight skip the sheet-pairing concerns that only
#: apply when there are two sides to reconcile.
def steps_for_mode(mode: QCRunMode) -> tuple[WizardStep, ...]:
    return (
        WizardStep.ROLES,
        WizardStep.SCANNING,
        WizardStep.PROFILE,
        WizardStep.REGIONS,
        WizardStep.REVIEW,
    )


@dataclass(frozen=True, slots=True)
class WarningItem:
    code: str
    message: str
    severity: Literal["caution", "block"] = "caution"


def compute_warnings(result: SetupAnalysisResult) -> tuple[WarningItem, ...]:
    """Bounded, content-free warnings computed from an already-completed
    scan -- never a raw cell value or a filename.
    """
    warnings: list[WarningItem] = []
    if result.failure_detail:
        warnings.append(
            WarningItem(
                code="setup_scan_failed",
                message=f"Setup analysis could not complete: {result.failure_detail}",
                severity="block",
            )
        )
    for member_id, member in result.members.items():
        if member.failure_detail:
            warnings.append(
                WarningItem(
                    code=f"member_scan_failed:{member_id}",
                    message=(
                        f"Member {member_id!r} could not be fully scanned: "
                        f"{member.failure_detail}"
                    ),
                )
            )
        for side, sheets in (
            ("baseline", member.baseline_sheets),
            ("current", member.current_sheets),
        ):
            for sheet in sheets:
                if sheet.failure_detail:
                    warnings.append(
                        WarningItem(
                            code=f"sheet_scan_failed:{member_id}:{side}:{sheet.sheet_name}",
                            message=(
                                f"{side} sheet {sheet.sheet_name!r} could not be "
                                f"fully analyzed: {sheet.failure_detail}"
                            ),
                        )
                    )
                for detected in sheet.regions:
                    if detected.ranked_candidate is not None:
                        warnings.append(
                            WarningItem(
                                code=(
                                    f"ranked_candidate:{member_id}:{sheet.sheet_name}:"
                                    f"{detected.region.cell_range}"
                                ),
                                message=(
                                    f"{sheet.sheet_name!r} {detected.region.cell_range} "
                                    "looks like a ranked/sorted table -- review its row "
                                    "matching before relying on cell-position comparison."
                                ),
                            )
                        )
        risk = member.current_xlsb_risk
        if risk is not None and not risk.safe_for_external_engine:
            warnings.append(
                WarningItem(
                    code=f"xlsb_risk:{member_id}",
                    message=(
                        "The current XLSB source has structural features "
                        "(external links or unrecognized content) that keep "
                        "formula-text enrichment degraded to presence-only checks."
                    ),
                )
            )
    return tuple(warnings)


def compute_sheet_pairing_warnings(state: ConfigWorkspaceState) -> tuple[WarningItem, ...]:
    """Added/removed-sheet acknowledgements (Step 8's "added and removed
    sheets surface as acknowledgements before run" criterion). A sheet the
    analyst has explicitly paired via a rename (``set_sheet_rename``) is
    excluded here -- it is no longer "added" or "removed", it is renamed.
    """
    warnings: list[WarningItem] = []
    for member in state.member_reviews:
        _pairs, added, removed = effective_sheet_pairing(member)
        for name in added:
            warnings.append(
                WarningItem(
                    code=f"sheet_added:{member.member_id}:{name}",
                    message=(
                        f"{name!r} is a new sheet in the current file, not present "
                        "in the baseline file."
                    ),
                    severity="block",
                )
            )
        for name in removed:
            warnings.append(
                WarningItem(
                    code=f"sheet_removed:{member.member_id}:{name}",
                    message=(
                        f"{name!r} from the baseline file is missing from the "
                        "current file."
                    ),
                    severity="block",
                )
            )
    return tuple(warnings)


@dataclass(frozen=True, slots=True)
class DiffEntry:
    scope: str
    kind: Literal["matches", "new_in_scan", "missing_from_profile", "conflict"]
    description: str


def diff_profile_against_scan(
    profile: DeliverableProfile,
    scan: SetupAnalysisResult,
    *,
    mode: QCRunMode,
) -> tuple[DiffEntry, ...]:
    """A compact field-level diff between a selected profile and the fresh
    scan (Step 7's "profile-vs-scan field diff" criterion). Never inspects
    cell content -- only sheet/member/region structural facts.
    """
    entries: list[DiffEntry] = []
    contract = profile.input_contract
    saved_sheet_names: set[str] = set()
    if contract is not None:
        for logical_member in contract.members:
            for logical_sheet in logical_member.sheets:
                if logical_sheet.preferred_sheet_name:
                    saved_sheet_names.add(logical_sheet.preferred_sheet_name)
    for member_id, member in scan.members.items():
        scanned_sheets = {sheet.sheet_name for sheet in member.current_sheets}
        legacy_ignored = set(profile.excel.ignore_sheets)
        if member_id != "primary":
            member_profile = profile.excel.members.get(member_id)
            legacy_ignored = set(member_profile.ignore_sheets) if member_profile else set()
        for sheet_name in sorted(scanned_sheets):
            if sheet_name in legacy_ignored:
                entries.append(
                    DiffEntry(
                        scope=f"{member_id}/{sheet_name}",
                        kind="conflict",
                        description=(
                            f"{sheet_name!r} is currently ignored by the saved "
                            "profile but is present in the fresh scan"
                        ),
                    )
                )
            elif contract is not None and sheet_name in saved_sheet_names:
                entries.append(
                    DiffEntry(
                        scope=f"{member_id}/{sheet_name}",
                        kind="matches",
                        description=f"{sheet_name!r} matches the saved configuration",
                    )
                )
            else:
                entries.append(
                    DiffEntry(
                        scope=f"{member_id}/{sheet_name}",
                        kind="new_in_scan",
                        description=f"{sheet_name!r} is new -- not yet in the saved profile",
                    )
                )
        if contract is not None:
            missing = sorted(saved_sheet_names - scanned_sheets)
            for sheet_name in missing:
                entries.append(
                    DiffEntry(
                        scope=f"{member_id}/{sheet_name}",
                        kind="missing_from_profile",
                        description=(
                            f"the saved profile expects {sheet_name!r}, which the "
                            "current scan did not find"
                        ),
                    )
                )
    return tuple(entries)


def is_clean_profile_diff(diff: tuple[DiffEntry, ...]) -> bool:
    """Whether every scope matched (Step 7's "compact clean-profile
    summary" criterion: a saved profile with no drift needs one line, not
    per-sheet confirmation).
    """
    return all(entry.kind == "matches" for entry in diff)


@dataclass(frozen=True, slots=True)
class ConfigWorkspaceState:
    """The workspace's own in-memory decisions for one browser session --
    never persisted verbatim (``qc_tool.history.config_session`` persists
    only its own bounded projection).
    """

    mode: QCRunMode
    profile_name: str = "default"
    member_reviews: tuple[MemberReview, ...] = ()
    warnings_acknowledged: frozenset[str] = frozenset()
    preview_side: Literal["baseline", "current"] = "current"

    def member_review(self, member_id: str) -> MemberReview | None:
        for member in self.member_reviews:
            if member.member_id == member_id:
                return member
        return None

    def with_member_review(self, updated: MemberReview) -> ConfigWorkspaceState:
        others = tuple(m for m in self.member_reviews if m.member_id != updated.member_id)
        return replace(self, member_reviews=(*others, updated))


def update_region_decision(
    member: MemberReview, sheet_name: str, region_id: str, **changes: object
) -> MemberReview:
    """Return a new ``MemberReview`` with one region decision replaced."""
    new_sheets: list[SheetReview] = []
    for sheet in member.current_sheets:
        if sheet.sheet_name != sheet_name:
            new_sheets.append(sheet)
            continue
        new_regions = tuple(
            replace(region, **changes) if region.region_id == region_id else region
            for region in sheet.regions
        )
        new_sheets.append(replace(sheet, regions=new_regions))
    return replace(member, current_sheets=tuple(new_sheets))


def parse_a1_cell(text: str) -> tuple[int, int]:
    """Parse a single-cell A1 reference (e.g. ``"B4"``) to ``(row, col)``,
    both 1-based. Raises ``ValueError`` with a plain, analyst-facing message
    on a range, garbage text, or a non-positive coordinate.
    """
    cleaned = text.strip().upper().replace("$", "")
    if not cleaned or ":" in cleaned:
        raise ValueError("enter a single cell reference, e.g. B4")
    try:
        row, col = coordinate_to_tuple(cleaned)
    except (ValueError, TypeError) as exc:
        raise ValueError("not a valid cell reference") from exc
    if row < 1 or col < 1:
        raise ValueError("not a valid cell reference")
    return row, col


def parse_a1_range(text: str) -> tuple[int, int, int, int]:
    """Parse an A1 range (``"B4:F20"``) or a single cell (treated as a 1x1
    range) to ``(min_row, min_col, max_row, max_col)``. Raises
    ``ValueError`` on anything unparseable or with reversed corners.
    """
    cleaned = text.strip().upper().replace("$", "")
    if not cleaned:
        raise ValueError("enter a range, e.g. B4:F20")
    if ":" not in cleaned:
        row, col = parse_a1_cell(cleaned)
        return row, col, row, col
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cleaned)
    except (ValueError, TypeError) as exc:
        raise ValueError("not a valid range") from exc
    if min_col is None or min_row is None or max_col is None or max_row is None:
        raise ValueError("not a valid range")
    if min_row > max_row or min_col > max_col:
        raise ValueError("range corners are reversed")
    return min_row, min_col, max_row, max_col


def format_a1_range(min_row: int, min_col: int, max_row: int, max_col: int) -> str:
    """Inverse of `parse_a1_range` -- ``"B4"`` for a 1x1 range, else
    ``"B4:F20"``.
    """
    top_left = f"{get_column_letter(min_col)}{min_row}"
    if min_row == max_row and min_col == max_col:
        return top_left
    return f"{top_left}:{get_column_letter(max_col)}{max_row}"


def region_with_bounds(
    region: RegionDecision, *, min_row: int, min_col: int, max_row: int, max_col: int
) -> RegionDecision:
    """Rebuild a region's anchor/range/available-columns from explicit
    bounds -- the shared tail end of both the "click a preview cell" and
    "type an A1 range" controls (Step 7's "click/A1 controls" criterion).
    Every column-role assignment that falls outside the new column span is
    dropped (it no longer names a real column); mode, header intent,
    exclusion detail, and confirmed all survive unchanged.
    """
    columns = _region_columns(min_col, max_col)
    return replace(
        region,
        anchor_cell=f"{get_column_letter(min_col)}{min_row}",
        current_range=format_a1_range(min_row, min_col, max_row, max_col),
        available_columns=columns,
        identity_columns=tuple(c for c in region.identity_columns if c in columns),
        ordinal_columns=tuple(c for c in region.ordinal_columns if c in columns),
        ignore_columns=tuple(c for c in region.ignore_columns if c in columns),
        expected_refresh_columns=tuple(
            c for c in region.expected_refresh_columns if c in columns
        ),
    )


def apply_anchor_click(region: RegionDecision, row: int, col: int) -> RegionDecision:
    """"Click controls": re-anchor a region at a clicked preview cell,
    preserving its current width/height (a shift, never a resize).
    """
    old_min_row, old_min_col, old_max_row, old_max_col = parse_a1_range(region.current_range)
    height = old_max_row - old_min_row
    width = old_max_col - old_min_col
    return region_with_bounds(
        region, min_row=row, min_col=col, max_row=row + height, max_col=col + width
    )


def apply_manual_range(region: RegionDecision, range_text: str) -> RegionDecision:
    """"A1 controls": move/resize a region to an explicitly typed range.
    Raises ``ValueError`` (surfaced to the analyst, never silently ignored)
    when the text does not parse.
    """
    min_row, min_col, max_row, max_col = parse_a1_range(range_text)
    return region_with_bounds(
        region, min_row=min_row, min_col=min_col, max_row=max_row, max_col=max_col
    )


def apply_region_transform(
    member: MemberReview,
    sheet_name: str,
    region_id: str,
    transform: Callable[[RegionDecision], RegionDecision],
) -> MemberReview:
    """Like `update_region_decision` but for edits that must recompute
    several fields together -- `apply_anchor_click`/`apply_manual_range` are
    the intended `transform` callables. Propagates a `ValueError` raised by
    `transform` (e.g. an unparseable typed range) to the caller unchanged,
    and raises one itself when the transformed region would overlap a
    sibling region on the same sheet (Step 8's "overlapping regions are
    rejected" criterion).
    """
    new_sheets = []
    for sheet in member.current_sheets:
        if sheet.sheet_name != sheet_name:
            new_sheets.append(sheet)
            continue
        new_regions: list[RegionDecision] = []
        transformed: RegionDecision | None = None
        for region in sheet.regions:
            if region.region_id == region_id:
                transformed = transform(region)
                new_regions.append(transformed)
            else:
                new_regions.append(region)
        if transformed is not None:
            for other in new_regions:
                if other.region_id != region_id and regions_overlap(transformed, other):
                    raise ValueError(
                        f"this range overlaps region {other.region_id!r} "
                        f"({other.current_range})"
                    )
        new_sheets.append(replace(sheet, regions=tuple(new_regions)))
    return replace(member, current_sheets=tuple(new_sheets))


def confirm_all_regions(member: MemberReview) -> MemberReview:
    """Bulk-confirm (Step 7's "bulk confirmation" criterion): mark every
    currently-valid region across every sheet as analyst-confirmed in one
    action instead of touching each region individually. An invalid region
    (e.g. keyed mode still missing identity columns) is left exactly as it
    is -- bulk confirmation never papers over an incomplete decision.
    """
    new_sheets = []
    for sheet in member.current_sheets:
        new_regions = tuple(
            replace(region, confirmed=True) if region.is_valid else region
            for region in sheet.regions
        )
        new_sheets.append(replace(sheet, regions=new_regions))
    return replace(member, current_sheets=tuple(new_sheets))


def unresolved_blockers(
    state: ConfigWorkspaceState, warnings: tuple[WarningItem, ...]
) -> tuple[str, ...]:
    """Everything that must be resolved before a final action may proceed."""
    blockers: list[str] = []
    for warning in warnings:
        if warning.severity == "block" and warning.code not in state.warnings_acknowledged:
            blockers.append(warning.message)
    for member in state.member_reviews:
        for sheet in member.current_sheets:
            for region in sheet.regions:
                if region.mode != "automatic" and not region.is_valid:
                    blockers.append(
                        f"{sheet.sheet_name!r} {region.current_range}: "
                        "finish configuring this region before continuing"
                    )
    return tuple(blockers)


def _resolved_column(region: RegionDecision, letter: str) -> ResolvedColumn:
    role: AlignmentRole = "none"
    policy: Literal["normal", "ignore", "expected_refresh"] = "normal"
    if letter in region.identity_columns:
        role = "identity"
    elif letter in region.ordinal_columns:
        role = "ordinal"
    elif letter in region.ignore_columns:
        policy = "ignore"
    elif letter in region.expected_refresh_columns:
        policy = "expected_refresh"
    return ResolvedColumn(
        column_id=slugify(f"{region.region_id}_{letter}", prefix="col"),
        baseline_letter=letter,
        current_letter=letter,
        alignment_role=role,
        comparison_policy=policy,
        trim_outer_whitespace=region.trim_identity_whitespace and role == "identity",
        coverage="confirmed" if (role != "none" or policy != "normal") else "automatic_confirmed",
    )


def _matching_baseline_region(
    member: MemberReview, current_sheet: SheetReview, region_index: int
) -> RegionDecision | None:
    """The baseline region auto-matched to one current region -- by ordinal
    position within the paired baseline sheet's own detected regions, the
    same conservative heuristic ``qc_tool.setup.analysis.analyze_member``
    already uses for its own ranked-candidate pairing. ``None`` when no
    baseline sheet is paired, or its region count does not line up simply.
    """
    pairs, _added, _removed = effective_sheet_pairing(member)
    baseline_name = next(
        (base for base, curr in pairs if curr == current_sheet.sheet_name), None
    )
    if baseline_name is None:
        return None
    baseline_sheet = next(
        (sheet for sheet in member.baseline_sheets if sheet.sheet_name == baseline_name), None
    )
    if baseline_sheet is None or region_index >= len(baseline_sheet.regions):
        return None
    return baseline_sheet.regions[region_index]


def _resolved_region(
    region: RegionDecision, baseline_region: RegionDecision | None
) -> ResolvedRegion:
    columns = tuple(
        _resolved_column(region, letter)
        for letter in (
            *region.identity_columns,
            *region.ordinal_columns,
            *region.ignore_columns,
            *region.expected_refresh_columns,
        )
    )
    coverage = "confirmed" if region.confirmed else "automatic_confirmed"
    if region.mode == "excluded":
        coverage = "degraded_acknowledged" if region.confirmed else "automatic_confirmed"
    elif region.mode == "positional":
        coverage = "positional"
    baseline_outer_range = region.baseline_range or (
        baseline_region.current_range if baseline_region is not None else None
    )
    return ResolvedRegion(
        region_id=region.region_id,
        mode=region.mode,
        header_intent=region.header_intent,
        baseline_outer_range=baseline_outer_range,
        current_outer_range=region.current_range,
        baseline_data_range=baseline_outer_range,
        current_data_range=region.current_range,
        baseline_first_data_row=(
            baseline_region.first_data_row if baseline_region is not None else None
        ),
        current_first_data_row=region.first_data_row,
        baseline_preamble_rows=(
            baseline_region.preamble_rows if baseline_region is not None else region.preamble_rows
        ),
        current_preamble_rows=region.preamble_rows,
        baseline_footer_rows=(
            baseline_region.footer_rows if baseline_region is not None else region.footer_rows
        ),
        current_footer_rows=region.footer_rows,
        columns=columns,
        duplicate_key_policy=region.duplicate_key_policy,
        blank_key_policy=region.blank_key_policy,
        coverage=coverage,
        degraded_reason=region.exclusion_reason,
    )


def build_resolved_configuration(
    state: ConfigWorkspaceState,
    *,
    profile: DeliverableProfile,
    profile_sha256: str,
    warnings_acknowledged: tuple[str, ...] = (),
) -> ResolvedInputConfigurationV1:
    """Build a real ``ResolvedInputConfigurationV1`` from the workspace's
    current decisions. Every region defaulting to ``automatic`` (untouched
    by the analyst) still resolves -- it simply carries no identity/ordinal
    columns and its physical range as-scanned, so an untouched sheet keeps
    running exactly as automatic detection would today.
    """
    members: list[ResolvedMember] = []
    for member in state.member_reviews:
        sheets: list[ResolvedSheet] = []
        pairs, _added, _removed = effective_sheet_pairing(member)
        baseline_by_current = {curr: base for base, curr in pairs}
        for sheet in member.current_sheets:
            baseline_name = baseline_by_current.get(sheet.sheet_name)
            regions = tuple(
                _resolved_region(region, _matching_baseline_region(member, sheet, index))
                for index, region in enumerate(sheet.regions)
            )
            selectors = tuple(
                ResolvedSelector(
                    selector_id=selector.selector_id,
                    baseline_cell=selector.cell if baseline_name else None,
                    current_cell=selector.cell,
                )
                for selector in sheet.selectors
            )
            sheets.append(
                ResolvedSheet(
                    sheet_id=sheet_id_for(member.member_id, sheet.sheet_name),
                    baseline_sheet_name=baseline_name,
                    current_sheet_name=sheet.sheet_name,
                    regions=regions,
                    selectors=selectors,
                    coverage="automatic_confirmed" if regions else "confirmed",
                )
            )
        members.append(
            ResolvedMember(
                member_id=member.member_id,
                sheets=tuple(sheets),
            )
        )
    return ResolvedInputConfigurationV1(
        source="input_contract",
        mode=state.mode,
        profile_name=profile.name,
        profile_sha256=profile_sha256,
        members=tuple(members),
        warnings_acknowledged=warnings_acknowledged,
    )


def _column_contract(region: RegionDecision, letter: str) -> LogicalColumnContract:
    role: AlignmentRole = "none"
    policy: Literal["normal", "ignore", "expected_refresh"] = "normal"
    exclusion: StructuralExclusionContract | None = None
    if letter in region.identity_columns:
        role = "identity"
    elif letter in region.ordinal_columns:
        role = "ordinal"
    elif letter in region.ignore_columns:
        policy = "ignore"
        exclusion = StructuralExclusionContract(
            reason=region.ignore_columns_reason or "ignored via the configuration workspace",
            expires_on=_parse_iso_date(region.ignore_columns_expires_on),
        )
    elif letter in region.expected_refresh_columns:
        policy = "expected_refresh"
    return LogicalColumnContract(
        column_id=slugify(f"{region.region_id}_{letter}", prefix="col"),
        alignment_role=role,
        comparison_policy=policy,
        trim_outer_whitespace=region.trim_identity_whitespace and role == "identity",
        exclusion=exclusion,
    )


def _region_contract(region: RegionDecision) -> LogicalRegionContract:
    exclusion: StructuralExclusionContract | None = None
    if region.mode == "excluded":
        exclusion = StructuralExclusionContract(
            reason=region.exclusion_reason or "excluded via the configuration workspace",
            expires_on=_parse_iso_date(region.exclusion_expires_on),
        )
    columns = tuple(
        _column_contract(region, letter)
        for letter in (
            *region.identity_columns,
            *region.ordinal_columns,
            *region.ignore_columns,
            *region.expected_refresh_columns,
        )
    )
    return LogicalRegionContract(
        region_id=region.region_id,
        mode=region.mode,
        header_intent=region.header_intent,
        anchor_cell=region.anchor_cell,
        preferred_current_range=region.current_range,
        preferred_first_data_row=region.first_data_row,
        columns=columns,
        blank_key_policy=region.blank_key_policy,
        duplicate_key_policy=region.duplicate_key_policy,
        exclusion=exclusion,
    )


def _parse_iso_date(value: str):
    import datetime as dt

    if not value:
        return dt.date.today() + dt.timedelta(days=90)
    return dt.date.fromisoformat(value)


def sheet_id_for(member_id: str, sheet_name: str) -> str:
    """The one stable slug shared by both the resolved configuration and any
    saved logical contract built from the same workspace decisions.
    """
    return slugify(f"{member_id}_{sheet_name}", prefix="sheet")


def _region_is_touched(region: RegionDecision) -> bool:
    """Whether a region carries ANY analyst decision worth saving -- the
    "only regions the analyst actually touched" test `build_input_contract`
    applies before recording a logical region.
    """
    return bool(
        region.mode != "automatic"
        or region.header_intent != "automatic"
        or region.identity_columns
        or region.ordinal_columns
        or region.ignore_columns
        or region.expected_refresh_columns
        or region.preamble_rows
        or region.footer_rows
        or region.blank_key_policy != "system_default"
        or region.baseline_range
    )


def build_input_contract(state: ConfigWorkspaceState) -> WorkbookInputContract:
    """Project the workspace's own decisions into a durable saved
    ``WorkbookInputContract`` for ``Save profile``/``Update profile``.

    Only regions the analyst actually touched (see ``_region_is_touched``)
    become saved logical regions -- an untouched, still-automatic sheet is
    recorded as a bare logical sheet hint (no regions), matching the
    architecture's own "a bare identity/preference hint... never conflicts
    with legacy fields" rule.
    """
    members: list[LogicalMemberContract] = []
    for member in state.member_reviews:
        sheets: list[LogicalSheetContract] = []
        for sheet in member.current_sheets:
            regions = tuple(
                _region_contract(region) for region in sheet.regions if _region_is_touched(region)
            )
            sheet_id = sheet_id_for(member.member_id, sheet.sheet_name)
            selectors = tuple(
                SelectorPrerequisiteContract(
                    selector_id=selector.selector_id,
                    label=selector.label,
                    owner_sheet_id=sheet_id,
                    preferred_current_cell=selector.cell,
                )
                for selector in sheet.selectors
            )
            sheets.append(
                LogicalSheetContract(
                    sheet_id=sheet_id,
                    preferred_sheet_name=sheet.sheet_name,
                    regions=regions,
                    selectors=selectors,
                )
            )
        members.append(LogicalMemberContract(member_id=member.member_id, sheets=tuple(sheets)))
    return WorkbookInputContract(members=tuple(members))

