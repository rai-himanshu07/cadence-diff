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

import datetime as dt
import re
from collections.abc import Callable, Mapping, Sequence
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


def _is_expired(iso_date: str) -> bool:
    """True when ``iso_date`` (an already-set ``exclusion_expires_on`` /
    ``ignore_columns_expires_on`` value) names a day strictly before today.

    An unset (empty) value is never "expired" here -- ``RegionDecision.
    is_valid`` already separately requires the field to be non-empty before
    reaching this check at all. An unparseable value is treated as expired
    (fail closed): a corrupted date is not a reason to silently keep
    excluding content.
    """
    if not iso_date:
        return False
    try:
        parsed = dt.date.fromisoformat(iso_date)
    except ValueError:
        return True
    return parsed < dt.date.today()


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
    #: Baseline-side letter override per column (Step 12 fix): pairs of
    #: (current_letter, baseline_letter) for any identity/ordinal/ignore/
    #: expected_refresh column whose position differs between baseline and
    #: current -- absent means "same letter both sides" (today's default,
    #: unchanged for every column with no entry here).
    column_baseline_letters: tuple[tuple[str, str], ...] = ()
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
    ranked_candidate_detected: bool = False
    ranked_candidate_pending: bool = False
    confirmed: bool = False

    @property
    def requires_exclusion_detail(self) -> bool:
        return self.mode == "excluded"

    @property
    def needs_ranked_resolution(self) -> bool:
        return (
            self.ranked_candidate_detected and self.mode == "automatic"
        ) or (self.ranked_candidate_pending and not self.confirmed)

    def baseline_letter_for(self, letter: str) -> str:
        """The baseline-side letter for a current-side column ``letter`` --
        the analyst-entered override if one exists, else the same letter.
        """
        for current, baseline in self.column_baseline_letters:
            if current == letter:
                return baseline
        return letter

    @property
    def is_valid(self) -> bool:
        if self.mode == "excluded":
            if not self.exclusion_reason or not self.exclusion_expires_on:
                return False
            if _is_expired(self.exclusion_expires_on):
                return False
        if self.mode == "keyed" and not self.identity_columns:
            return False
        if self.header_intent == "first_data_row" and self.first_data_row is None:
            return False
        if self.ignore_columns:
            if not self.ignore_columns_reason or not self.ignore_columns_expires_on:
                return False
            if _is_expired(self.ignore_columns_expires_on):
                return False
        return True


@dataclass(frozen=True, slots=True)
class SelectorDecision:
    """One analyst-declared selector/scenario prerequisite cell (Step 8's
    "explicit selector prerequisites for dropdown/filter/scenario/parameter
    cells" criterion). Never carries a value -- only a label and cell
    location(s); the run-time engine check compares the two files' actual
    saved values without ever persisting either one.
    """

    selector_id: str
    label: str
    cell: str  # current-side A1 cell
    #: Baseline-side cell override (Step 12 fix) -- "" means "same cell as
    #: current" (today's default, unchanged for every existing selector).
    baseline_cell: str = ""


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
                ranked_candidate_detected=detected.ranked_candidate is not None,
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


def set_column_baseline_letter(
    region: RegionDecision, letter: str, baseline_letter: str
) -> RegionDecision:
    """Set (or clear, when ``baseline_letter`` is blank) an analyst-entered
    baseline-side letter override for one current-side column ``letter``
    (Step 12 fix -- "logical columns resolve separately on each side").
    """
    remaining = tuple(
        (current, baseline)
        for current, baseline in region.column_baseline_letters
        if current != letter
    )
    baseline_letter = baseline_letter.strip().upper()
    if not baseline_letter or baseline_letter == letter:
        return replace(region, column_baseline_letters=remaining)
    return replace(
        region, column_baseline_letters=(*remaining, (letter, baseline_letter))
    )


def set_selector_baseline_cell(
    member: MemberReview, sheet_name: str, selector_id: str, baseline_cell: str
) -> MemberReview:
    """Set (or clear, when ``baseline_cell`` is blank) an analyst-entered
    baseline-side cell override for one selector (Step 12 fix). Raises
    ``ValueError`` for an unparseable non-blank cell.
    """
    cleaned = baseline_cell.strip().upper()
    if cleaned:
        parse_a1_cell(cleaned)  # validates; raises ValueError with a plain message
    new_sheets = []
    for sheet in member.current_sheets:
        if sheet.sheet_name != sheet_name:
            new_sheets.append(sheet)
            continue
        new_sheets.append(
            replace(
                sheet,
                selectors=tuple(
                    replace(selector, baseline_cell=cleaned)
                    if selector.selector_id == selector_id
                    else selector
                    for selector in sheet.selectors
                ),
            )
        )
    return replace(member, current_sheets=tuple(new_sheets))


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


def compute_warnings(
    result: SetupAnalysisResult,
    *,
    state: ConfigWorkspaceState | None = None,
) -> tuple[WarningItem, ...]:
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
                    ranked_is_pending = True
                    if state is not None:
                        review = state.member_review(member_id)
                        reviewed_sheet = next(
                            (
                                item
                                for item in review.current_sheets
                                if item.sheet_name == sheet.sheet_name
                            ),
                            None,
                        ) if review is not None else None
                        reviewed_region = next(
                            (
                                item
                                for item in reviewed_sheet.regions
                                if item.current_range == detected.region.cell_range
                            ),
                            None,
                        ) if reviewed_sheet is not None else None
                        ranked_is_pending = bool(
                            reviewed_region is not None
                            and reviewed_region.needs_ranked_resolution
                        )
                    if detected.ranked_candidate is not None and ranked_is_pending:
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


#: Same numeric bar ``qc_tool.excel.ranked_identity.MIN_KEY_OVERLAP`` uses
#: for the auto-detector -- kept as an independent constant (not imported)
#: since this module never depends on that detection layer, but the two
#: thresholds are DELIBERATELY kept in sync so "low overlap" means the same
#: real-world degree of mismatch whether it comes from automatic detection
#: or a confirmed manual selection.
LOW_KEY_OVERLAP_THRESHOLD = 0.90


def low_key_overlap_warning_code(member_id: str, sheet_name: str, region_id: str) -> str:
    """The stable warning code for one region's low-confirmed-identity-
    overlap acknowledgement (Step 12 Fix 5) -- shared between the UI
    warnings list and ``build_resolved_configuration``'s coverage-state
    lookup so the two always agree on which region an acknowledgement
    covers.
    """
    return f"low_key_overlap:{member_id}:{sheet_name}:{region_id}"


def compute_key_overlap_warnings(
    state: ConfigWorkspaceState, key_overlap_ratios: Mapping[str, float]
) -> tuple[WarningItem, ...]:
    """One blocking warning per keyed region whose last-computed, analyst-
    CONFIRMED identity-column overlap ratio measured below
    ``LOW_KEY_OVERLAP_THRESHOLD`` (Step 12 Fix 5).

    ``severity="block"`` (matching ``compute_warnings``/``compute_sheet_
    pairing_warnings``/``compute_slide_pairing_warnings`` above): a low
    overlap on the analyst's own confirmed columns is a real, disclosed
    degradation that must be explicitly ACKNOWLEDGED (checked in the
    workspace's Warnings section) before any final action may proceed --
    the analyst may still accept it and continue (mirrors
    ``allow_large_workbooks``'s "acknowledge and continue" precedent), but
    silently ignoring it is not an option. The acknowledgement itself
    stays run-only: ``state.warnings_acknowledged`` lives only on the
    session, never persisted into a saved ``DeliverableProfile``.
    ``key_overlap_ratios`` is keyed by ``region_id`` -- a region absent
    from it was never queried and gets no warning (never a false "clean"
    claim from silence).
    """
    warnings: list[WarningItem] = []
    for member in state.member_reviews:
        for sheet in member.current_sheets:
            for region in sheet.regions:
                if region.mode != "keyed":
                    continue
                ratio = key_overlap_ratios.get(region.region_id)
                if ratio is None or ratio >= LOW_KEY_OVERLAP_THRESHOLD:
                    continue
                warnings.append(
                    WarningItem(
                        code=low_key_overlap_warning_code(
                            member.member_id, sheet.sheet_name, region.region_id
                        ),
                        message=(
                            f"{sheet.sheet_name!r} {region.current_range}: the "
                            "confirmed identity columns only overlap "
                            f"{ratio:.0%} between baseline and current -- row "
                            "matching may be unreliable for a large share of "
                            "this region."
                        ),
                        severity="block",
                    )
                )
    return tuple(warnings)


def region_key_overlap_query_bounds(
    region: RegionDecision, baseline_region: RegionDecision | None
) -> tuple[int, int, int, int] | None:
    """``(baseline_first_row, baseline_last_row, current_first_row,
    current_last_row)`` for a bounded key-overlap query over this region's
    resolved ranges, or ``None`` when either side's range cannot be parsed.

    Deliberately simpler than ``_resolved_region``'s own preamble/footer
    composition: it queries the WHOLE resolved range on each side,
    including any header/preamble row. This is a disclosed, bounded
    approximation for a diagnostic ratio (never authoritative evidence,
    matching this module's own "setup analysis never becomes evidence"
    boundary) -- a header row counted on both sides shifts the measured
    ratio by at most one row out of the whole region, immaterial to a 0-1
    overlap fraction over any real table.
    """
    baseline_range = region.baseline_range or (
        baseline_region.current_range if baseline_region is not None else None
    )
    if not baseline_range or not region.current_range:
        return None
    try:
        base_min_row, _, base_max_row, _ = parse_a1_range(baseline_range)
        curr_min_row, _, curr_max_row, _ = parse_a1_range(region.current_range)
    except ValueError:
        return None
    return (base_min_row, base_max_row, curr_min_row, curr_max_row)


@dataclass(frozen=True, slots=True)
class SlideDecision:
    """One slide's analyst-facing configuration decision (Step 9) --
    PowerPoint's counterpart to a keyed row/column decision. ``included``
    is only meaningful on the CURRENT side (Step 9's "cycle slide
    inclusion" criterion); a baseline-side entry always stays included,
    since it exists purely to compute added/removed acknowledgements.
    """

    slide_index: int  # 1-based
    title: str
    included: bool = True


@dataclass(frozen=True, slots=True)
class DeckReview:
    """One PowerPoint deck's slide inventory and decisions -- a single
    global review (mirrors ``DeliverableProfile.ppt`` being one profile-
    wide section, unlike Excel's per-member ``MemberReview``). Absent
    (``None`` at the ``ConfigWorkspaceState`` level) whenever the current
    mode/role selection has no PPT file at all.
    """

    baseline_slides: tuple[SlideDecision, ...] = ()
    current_slides: tuple[SlideDecision, ...] = ()
    #: current title -> baseline title, mirrors ``MemberReview.sheet_renames``.
    slide_renames: dict[str, str] = field(default_factory=dict)
    failure_detail: str = ""


def deck_review_from_titles(
    baseline_titles: Sequence[tuple[int, str]],
    current_titles: Sequence[tuple[int, str]],
) -> DeckReview:
    """Seed a fresh ``DeckReview`` from ``qc_tool.io.peek.peek_slide_titles``
    -shaped ``(1-based index, title)`` pairs -- every slide defaults to
    included, mirroring every region's ``automatic`` default.
    """
    return DeckReview(
        baseline_slides=tuple(
            SlideDecision(slide_index=index, title=title) for index, title in baseline_titles
        ),
        current_slides=tuple(
            SlideDecision(slide_index=index, title=title) for index, title in current_titles
        ),
    )


def set_slide_included(deck: DeckReview, slide_index: int, included: bool) -> DeckReview:
    """Toggle one CURRENT-side slide's inclusion (Step 9's "cycle slide
    inclusion" criterion)."""
    new_slides = tuple(
        replace(slide, included=included) if slide.slide_index == slide_index else slide
        for slide in deck.current_slides
    )
    return replace(deck, current_slides=new_slides)


def slide_pairing(
    deck: DeckReview,
) -> tuple[tuple[tuple[str, str], ...], tuple[str, ...], tuple[str, ...]]:
    """``(same_title_pairs, added_current_only, removed_baseline_only)`` --
    PowerPoint's counterpart to ``sheet_pairing``: same-title slides
    pre-pair automatically, anything else needs an explicit rename or
    acknowledgement, never a guessed split/merge.
    """
    current_titles = {slide.title for slide in deck.current_slides}
    baseline_titles = {slide.title for slide in deck.baseline_slides}
    paired = tuple(sorted(current_titles & baseline_titles))
    added = tuple(sorted(current_titles - baseline_titles))
    removed = tuple(sorted(baseline_titles - current_titles))
    return tuple((name, name) for name in paired), added, removed


def set_slide_rename(
    deck: DeckReview, current_title: str, baseline_title: str | None
) -> DeckReview:
    """Declare (``baseline_title`` not None) or clear (``None``) that
    ``current_title`` is a rename of ``baseline_title`` (Step 9's "explicit
    renamed-slide pins" criterion). One-to-one by construction, mirrors
    ``set_sheet_rename`` exactly.
    """
    renames = dict(deck.slide_renames)
    renames.pop(current_title, None)
    if baseline_title is not None:
        for other_current, other_baseline in list(renames.items()):
            if other_baseline == baseline_title:
                del renames[other_current]
        renames[current_title] = baseline_title
    return replace(deck, slide_renames=renames)


def effective_slide_pairing(
    deck: DeckReview,
) -> tuple[tuple[tuple[str, str], ...], tuple[str, ...], tuple[str, ...]]:
    """``slide_pairing`` plus every analyst-declared rename promoted from
    added/removed into a real pair -- mirrors ``effective_sheet_pairing``.
    """
    pairs, added, removed = slide_pairing(deck)
    renamed_pairs = tuple(
        (baseline_title, current_title)
        for current_title, baseline_title in sorted(deck.slide_renames.items())
        if current_title in added and baseline_title in removed
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


def compute_slide_pairing_warnings(deck: DeckReview | None) -> tuple[WarningItem, ...]:
    """Added/removed-slide acknowledgements for cycle mode (Step 9's "cycle
    slide inclusion and explicit renamed-slide pins" criterion) -- mirrors
    ``compute_sheet_pairing_warnings`` exactly. Never fires for a single-
    sided deck review (preflight/final-package have no baseline deck).
    """
    if deck is None or not deck.baseline_slides:
        return ()
    warnings: list[WarningItem] = []
    _pairs, added, removed = effective_slide_pairing(deck)
    for title in added:
        warnings.append(
            WarningItem(
                code=f"slide_added:{title}",
                message=(
                    f"{title!r} is a new slide in the current deck, not present "
                    "in the baseline deck."
                ),
                severity="block",
            )
        )
    for title in removed:
        warnings.append(
            WarningItem(
                code=f"slide_removed:{title}",
                message=(
                    f"{title!r} from the baseline deck is missing from the "
                    "current deck."
                ),
                severity="block",
            )
        )
    return tuple(warnings)


def compute_required_slide_warnings(
    deck: DeckReview | None, required_slides: Sequence[str]
) -> tuple[WarningItem, ...]:
    """Required-slide consistency (Step 9): every profile-required slide
    title must exist among the CURRENT deck's slides. A caution, never a
    hard block here -- the engine's own required-slide check remains
    authoritative for what actually fails a real run.
    """
    if deck is None or not required_slides:
        return ()
    current_titles = {slide.title for slide in deck.current_slides}
    return tuple(
        WarningItem(
            code=f"required_slide_missing:{title}",
            message=f"Required slide {title!r} was not found in the current deck.",
        )
        for title in required_slides
        if title not in current_titles
    )


def resolve_slide_anchors(
    deck: DeckReview | None, anchor_titles: Sequence[str]
) -> tuple[WarningItem, ...]:
    """Final-package logical slide-anchor resolution (Step 9): for each
    DISTINCT saved anchor title (a ``CrosscheckMapping.slide`` value),
    resolve it against the current deck's slide titles from title alone
    (no ordinal is stored on a saved mapping today). Zero matches is
    "missing"; two or more matches is "ambiguous" -- both are warning-only
    disclosures, NEVER a blocker and NEVER a guess. Normal QC's own
    mapping-verification findings remain authoritative for what actually
    reconciles at run time; this is purely a pre-run heads-up.
    """
    if deck is None:
        return ()
    counts: dict[str, int] = {}
    for slide in deck.current_slides:
        counts[slide.title] = counts.get(slide.title, 0) + 1
    warnings: list[WarningItem] = []
    seen: set[str] = set()
    for title in anchor_titles:
        if not title or title in seen:
            continue
        seen.add(title)
        count = counts.get(title, 0)
        if count == 0:
            warnings.append(
                WarningItem(
                    code=f"slide_anchor_missing:{title}",
                    message=(
                        f"Saved reference to slide {title!r} was not found in "
                        "the current deck -- this mapping is stale."
                    ),
                )
            )
        elif count > 1:
            warnings.append(
                WarningItem(
                    code=f"slide_anchor_ambiguous:{title}",
                    message=(
                        f"{count} slides in the current deck share the title "
                        f"{title!r} -- the saved reference is ambiguous and "
                        "will not be guessed."
                    ),
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


def _resolved_column_signature(
    columns: tuple[ResolvedColumn, ...],
) -> dict[str, tuple[str, str]]:
    return {c.column_id: (c.alignment_role, c.comparison_policy) for c in columns}


def _diff_resolved_columns(
    prev_columns: tuple[ResolvedColumn, ...], curr_columns: tuple[ResolvedColumn, ...]
) -> str:
    """One concise clause naming how many columns' role/policy changed, or
    "" when nothing did. Never names a column id (content-free, matching
    ``LogicalFindingAddress``'s own bounded-identifier discipline).
    """
    prev_roles = _resolved_column_signature(prev_columns)
    curr_roles = _resolved_column_signature(curr_columns)
    changed = [
        column_id
        for column_id in sorted(set(prev_roles) | set(curr_roles))
        if prev_roles.get(column_id) != curr_roles.get(column_id)
    ]
    if not changed:
        return ""
    return f"column role/policy changed for {len(changed)} column(s)"


def _diff_resolved_selectors(
    scope: str,
    prev_selectors: tuple[ResolvedSelector, ...],
    curr_selectors: tuple[ResolvedSelector, ...],
) -> list[DiffEntry]:
    entries: list[DiffEntry] = []
    prev_ids = {s.selector_id for s in prev_selectors}
    curr_ids = {s.selector_id for s in curr_selectors}
    for selector_id in sorted(curr_ids - prev_ids):
        entries.append(
            DiffEntry(
                scope=f"{scope}/{selector_id}",
                kind="new_in_scan",
                description="a selector prerequisite is new since the predecessor run",
            )
        )
    for selector_id in sorted(prev_ids - curr_ids):
        entries.append(
            DiffEntry(
                scope=f"{scope}/{selector_id}",
                kind="missing_from_profile",
                description="a selector prerequisite is no longer present",
            )
        )
    return entries


def _diff_resolved_regions(
    scope: str,
    prev_regions: tuple[ResolvedRegion, ...],
    curr_regions: tuple[ResolvedRegion, ...],
) -> list[DiffEntry]:
    entries: list[DiffEntry] = []
    prev_by_id = {r.region_id: r for r in prev_regions}
    curr_by_id = {r.region_id: r for r in curr_regions}
    for region_id in sorted(set(prev_by_id) | set(curr_by_id)):
        prev_region = prev_by_id.get(region_id)
        curr_region = curr_by_id.get(region_id)
        region_scope = f"{scope}/{region_id}"
        if prev_region is None:
            entries.append(
                DiffEntry(
                    scope=region_scope,
                    kind="new_in_scan",
                    description="region is new since the predecessor run",
                )
            )
            continue
        if curr_region is None:
            entries.append(
                DiffEntry(
                    scope=region_scope,
                    kind="missing_from_profile",
                    description="region is no longer present",
                )
            )
            continue
        clauses: list[str] = []
        if prev_region.mode != curr_region.mode:
            clauses.append(f"mode {prev_region.mode!r} -> {curr_region.mode!r}")
        if prev_region.header_intent != curr_region.header_intent:
            clauses.append(
                f"header intent {prev_region.header_intent!r} -> "
                f"{curr_region.header_intent!r}"
            )
        if prev_region.duplicate_key_policy != curr_region.duplicate_key_policy:
            clauses.append(
                "duplicate-key policy "
                f"{prev_region.duplicate_key_policy!r} -> "
                f"{curr_region.duplicate_key_policy!r}"
            )
        if prev_region.blank_key_policy != curr_region.blank_key_policy:
            clauses.append(
                f"blank-key policy {prev_region.blank_key_policy!r} -> "
                f"{curr_region.blank_key_policy!r}"
            )
        column_clause = _diff_resolved_columns(prev_region.columns, curr_region.columns)
        if column_clause:
            clauses.append(column_clause)
        if clauses:
            entries.append(
                DiffEntry(
                    scope=region_scope,
                    kind="conflict",
                    description="; ".join(clauses),
                )
            )
    return entries


def _diff_resolved_member(
    member_id: str, prev_member: ResolvedMember, curr_member: ResolvedMember
) -> list[DiffEntry]:
    entries: list[DiffEntry] = []
    prev_sheets = {s.sheet_id: s for s in prev_member.sheets}
    curr_sheets = {s.sheet_id: s for s in curr_member.sheets}
    for sheet_id in sorted(set(prev_sheets) | set(curr_sheets)):
        prev_sheet = prev_sheets.get(sheet_id)
        curr_sheet = curr_sheets.get(sheet_id)
        scope = f"{member_id}/{sheet_id}"
        if prev_sheet is None:
            entries.append(
                DiffEntry(
                    scope=scope,
                    kind="new_in_scan",
                    description="sheet is new since the predecessor run",
                )
            )
            continue
        if curr_sheet is None:
            entries.append(
                DiffEntry(
                    scope=scope,
                    kind="missing_from_profile",
                    description="sheet is no longer present",
                )
            )
            continue
        if (
            prev_sheet.baseline_sheet_name != curr_sheet.baseline_sheet_name
            or prev_sheet.current_sheet_name != curr_sheet.current_sheet_name
        ):
            entries.append(
                DiffEntry(
                    scope=scope,
                    kind="conflict",
                    description="sheet pairing changed since the predecessor run",
                )
            )
        entries.extend(
            _diff_resolved_regions(scope, prev_sheet.regions, curr_sheet.regions)
        )
        entries.extend(
            _diff_resolved_selectors(scope, prev_sheet.selectors, curr_sheet.selectors)
        )
    return entries


def diff_resolved_configurations(
    previous: ResolvedInputConfigurationV1 | None,
    current: ResolvedInputConfigurationV1 | None,
) -> tuple[DiffEntry, ...]:
    """Concise, run-specific diff between a run's resolved configuration and
    its predecessor's (plan-20260913 Step 10's "History surfaces config diff
    versus predecessor using persisted resolved snapshots, not heuristics"
    criterion).

    Compares the exact per-run RESOLUTION -- member/sheet pairing, region
    mode/header-intent/key-policy, column role/policy composition, and
    selector prerequisite declarations -- never just ``profile_sha256``, so
    a policy-only edit that changed a run's actual resolution is disclosed
    even when someone only glances at the profile hash. Deliberately
    excludes pure numeric range/first-data-row drift, which is the ROUTINE,
    expected difference between almost any two runs over growing data and
    would make this "concise" summary noisy rather than useful. Never
    touches cell content or selector values -- only stable logical
    ids/enums already present on ``ResolvedInputConfigurationV1``.

    ``None`` on one side only (a legacy run recorded before the input
    contract existed, compared against one that has it) is disclosed as a
    single explanatory entry rather than silently producing an empty diff
    that would misleadingly read as "nothing changed". ``None`` on both
    sides (neither run used the input contract) yields an empty diff.
    """
    if previous is None and current is None:
        return ()
    if previous is None or current is None:
        return (
            DiffEntry(
                scope="configuration",
                kind="conflict",
                description=(
                    "one of these two runs has no saved resolved configuration "
                    "to compare against (a legacy run recorded before the input "
                    "contract existed) -- no run-specific configuration diff is "
                    "available"
                ),
            ),
        )
    entries: list[DiffEntry] = []
    previous_members = {m.member_id: m for m in previous.members}
    current_members = {m.member_id: m for m in current.members}
    for member_id in sorted(set(previous_members) | set(current_members)):
        prev_member = previous_members.get(member_id)
        curr_member = current_members.get(member_id)
        if prev_member is None:
            entries.append(
                DiffEntry(
                    scope=member_id,
                    kind="new_in_scan",
                    description="member is new since the predecessor run",
                )
            )
            continue
        if curr_member is None:
            entries.append(
                DiffEntry(
                    scope=member_id,
                    kind="missing_from_profile",
                    description="member is no longer present",
                )
            )
            continue
        entries.extend(_diff_resolved_member(member_id, prev_member, curr_member))
    return tuple(entries)


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
    #: Absent whenever the current mode/role selection has no PPT file.
    deck_review: DeckReview | None = None
    #: Run-only workload/dependency safety overrides (Step 9's "run-only
    #: workload/dependency overrides" criterion) -- editable live in the
    #: workspace rather than fixed at session-creation time.
    allow_large_workbooks: bool = False
    allow_dependency_indexing: bool = False
    #: The saved profile file's content hash at the moment `profile_name`
    #: was last set in this session (`None` when no file exists under that
    #: name yet, e.g. the built-in unsaved "default"). Step 11's
    #: "optimistic conflict protection" criterion: a save compares this
    #: against the file's CURRENT hash immediately before writing, so a
    #: concurrent edit from another tab/session is never silently
    #: clobbered. Never the file's bytes themselves -- only a digest.
    profile_opened_hash: str | None = None

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


def set_region_data_start(
    region: RegionDecision, first_data_row: int | None
) -> RegionDecision:
    """Set the one canonical current-side data-start boundary.

    ``None`` restores automatic detection. The outer range's first row means
    there is no header; a later row means every preceding row is preamble.
    Legacy ``preamble_rows`` is cleared so two controls can never disagree.
    """
    min_row, _min_col, max_row, _max_col = parse_a1_range(region.current_range)
    if first_data_row is None:
        return replace(
            region,
            header_intent="automatic",
            first_data_row=None,
            preamble_rows=0,
        )
    last_data_row = max_row - region.footer_rows
    if first_data_row < min_row or first_data_row > last_data_row:
        raise ValueError(
            f"data start must be between rows {min_row} and {last_data_row}"
        )
    return replace(
        region,
        header_intent=("no_header" if first_data_row == min_row else "first_data_row"),
        first_data_row=(None if first_data_row == min_row else first_data_row),
        preamble_rows=0,
    )


def set_region_footer_rows(region: RegionDecision, footer_rows: int) -> RegionDecision:
    """Set a footer count without allowing it to consume the data band."""
    min_row, _min_col, max_row, _max_col = parse_a1_range(region.current_range)
    if footer_rows < 0:
        raise ValueError("footer rows cannot be negative")
    data_start = (
        region.first_data_row
        if region.header_intent == "first_data_row"
        and region.first_data_row is not None
        else min_row
        if region.header_intent == "no_header"
        else min_row + region.preamble_rows
    )
    if footer_rows > max_row - data_start:
        raise ValueError("footer rows must leave at least one data row")
    return replace(region, footer_rows=footer_rows)


def effective_region_data_range(
    region: RegionDecision, *, outer_range: str | None = None
) -> str | None:
    """Project the region's relative data boundaries onto one side's range."""
    current_min_row, _current_min_col, _current_max_row, _current_max_col = (
        parse_a1_range(region.current_range)
    )
    min_row, min_col, max_row, max_col = parse_a1_range(
        outer_range or region.current_range
    )
    if region.header_intent == "first_data_row" and region.first_data_row is not None:
        data_min_row = min_row + (region.first_data_row - current_min_row)
    elif region.header_intent == "no_header":
        data_min_row = min_row
    elif region.preamble_rows > 0:
        data_min_row = min_row + region.preamble_rows
    else:
        return None
    data_max_row = max_row - region.footer_rows
    if data_min_row > data_max_row:
        return None
    return format_a1_range(data_min_row, min_col, data_max_row, max_col)


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
    old_min_row, _old_min_col, _old_max_row, _old_max_col = parse_a1_range(
        region.current_range
    )
    first_data_row = (
        region.first_data_row + (min_row - old_min_row)
        if region.header_intent == "first_data_row"
        and region.first_data_row is not None
        else None
    )
    preamble_rows = (
        region.preamble_rows
        if region.header_intent not in {"no_header", "first_data_row"}
        else 0
    )
    data_start = first_data_row or min_row + preamble_rows
    if data_start > max_row - region.footer_rows:
        raise ValueError(
            "the resized range must leave at least one data row after "
            "header and footer boundaries"
        )
    columns = _region_columns(min_col, max_col)
    return replace(
        region,
        anchor_cell=f"{get_column_letter(min_col)}{min_row}",
        current_range=format_a1_range(min_row, min_col, max_row, max_col),
        first_data_row=first_data_row,
        preamble_rows=preamble_rows,
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
            replace(
                region,
                confirmed=True,
                ranked_candidate_pending=False,
            )
            if region.is_valid
            and not (
                region.ranked_candidate_detected
                and region.mode == "automatic"
            )
            else region
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
                if (
                    region.ranked_candidate_detected
                    and region.mode == "automatic"
                ):
                    blockers.append(
                        f"{sheet.sheet_name!r} {region.current_range}: choose "
                        "Match rows by key, Compare by position, or Exclude "
                        "before running"
                    )
                elif region.needs_ranked_resolution:
                    blockers.append(
                        f"{sheet.sheet_name!r} {region.current_range}: confirm "
                        "the row setup before running"
                    )
                elif region.mode == "excluded" and region.exclusion_reason and _is_expired(
                    region.exclusion_expires_on
                ):
                    blockers.append(
                        f"{sheet.sheet_name!r} {region.current_range}: this "
                        "exclusion expired on "
                        f"{region.exclusion_expires_on}; renew or remove it "
                        "before continuing"
                    )
                elif region.ignore_columns and region.ignore_columns_reason and _is_expired(
                    region.ignore_columns_expires_on
                ):
                    blockers.append(
                        f"{sheet.sheet_name!r} {region.current_range}: the "
                        "ignored-column exclusion expired on "
                        f"{region.ignore_columns_expires_on}; renew or remove "
                        "it before continuing"
                    )
                elif region.mode != "automatic" and not region.is_valid:
                    blockers.append(
                        f"{sheet.sheet_name!r} {region.current_range}: "
                        "finish configuring this region before continuing"
                    )
    return tuple(blockers)


@dataclass(frozen=True, slots=True)
class WorkspaceReadiness:
    """One immutable decision used by both final-action rendering and handlers."""

    ready: bool
    blockers: tuple[str, ...]


def compute_workspace_readiness(
    state: ConfigWorkspaceState,
    warnings: tuple[WarningItem, ...],
    *,
    setup_complete: bool,
    expected_excel_member_ids: frozenset[str],
    ppt_required: bool,
) -> WorkspaceReadiness:
    """Require terminal setup plus an exact, fully hydrated artifact review."""
    blockers = list(unresolved_blockers(state, warnings))
    if not setup_complete:
        blockers.append("Setup analysis is not complete")
    actual_member_ids = {member.member_id for member in state.member_reviews}
    for member_id in sorted(expected_excel_member_ids - actual_member_ids):
        blockers.append(f"missing required Excel member review: {member_id}")
    for member_id in sorted(actual_member_ids - expected_excel_member_ids):
        blockers.append(f"stale Excel member review: {member_id}")
    if ppt_required and state.deck_review is None:
        blockers.append("PowerPoint review is not ready")
    return WorkspaceReadiness(ready=not blockers, blockers=tuple(blockers))


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
        baseline_letter=region.baseline_letter_for(letter),
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
    region: RegionDecision,
    baseline_region: RegionDecision | None,
    *,
    key_overlap_ratio: float | None = None,
    low_overlap_acknowledged: bool = False,
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
        # A deliberate content exclusion is its OWN distinct coverage state
        # -- never "degraded_acknowledged", which is reserved for a
        # DIFFERENT concept (a forced capability limitation the analyst
        # accepted, e.g. low-overlap key alignment), not an intentional
        # scope choice.
        coverage = "excluded"
    elif region.mode == "positional":
        coverage = "positional"
    elif (
        region.mode == "keyed"
        and key_overlap_ratio is not None
        and key_overlap_ratio < LOW_KEY_OVERLAP_THRESHOLD
        and low_overlap_acknowledged
    ):
        # The canonical "forced capability limitation the analyst accepted"
        # case (Step 12 Fix 5): the analyst's own confirmed identity columns
        # measure a low overlap between the two files, and the analyst
        # explicitly acknowledged running with that degradation anyway.
        coverage = "degraded_acknowledged"
    baseline_outer_range = region.baseline_range or (
        baseline_region.current_range if baseline_region is not None else None
    )
    current_first_data_row = (
        region.first_data_row if region.header_intent == "first_data_row" else None
    )
    current_preamble_rows = (
        region.preamble_rows
        if region.header_intent not in {"no_header", "first_data_row"}
        else 0
    )
    current_outer_top = parse_a1_range(region.current_range)[0]
    baseline_outer_top = (
        parse_a1_range(baseline_outer_range)[0]
        if baseline_outer_range is not None
        else None
    )
    baseline_first_data_row = (
        baseline_outer_top + (current_first_data_row - current_outer_top)
        if baseline_outer_top is not None
        and current_first_data_row is not None
        else None
    )
    return ResolvedRegion(
        region_id=region.region_id,
        mode=region.mode,
        header_intent=region.header_intent,
        baseline_outer_range=baseline_outer_range,
        current_outer_range=region.current_range,
        baseline_data_range=baseline_outer_range,
        current_data_range=region.current_range,
        baseline_first_data_row=baseline_first_data_row,
        current_first_data_row=current_first_data_row,
        baseline_preamble_rows=current_preamble_rows,
        current_preamble_rows=current_preamble_rows,
        baseline_footer_rows=region.footer_rows,
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
    file_hashes: dict[str, str] | None = None,
    key_overlap_ratios: Mapping[str, float] | None = None,
) -> ResolvedInputConfigurationV1:
    """Build a real ``ResolvedInputConfigurationV1`` from the workspace's
    current decisions. Every region defaulting to ``automatic`` (untouched
    by the analyst) still resolves -- it simply carries no identity/ordinal
    columns and its physical range as-scanned, so an untouched sheet keeps
    running exactly as automatic detection would today.

    ``file_hashes`` is the run's own role-keyed source hashes (the exact
    shape ``qc_tool.run_preflight.hash_run_files`` produces, e.g.
    ``"baseline_excel"``/``"current_excel"`` for the primary member,
    ``"baseline_excel:member_id"`` otherwise -- ``PackageManifest.role_key``'s
    own convention). Without it, every member's ``baseline_source_sha256``/
    ``current_source_sha256`` stay ``None`` and ``validate_freshness()``
    unconditionally rejects the resulting configuration as stale the
    moment a run actually tries to use it -- this parameter is required
    for any resolved configuration that will reach ``perform_run()``.

    ``key_overlap_ratios`` (Step 12 Fix 5) is the workspace's last-computed
    confirmed-identity overlap ratio per region, keyed by ``region_id`` --
    an ephemeral, on-demand query result never persisted as a draft choice
    (mirrors ``file_hashes``'s own "extra fact supplied at build time"
    shape). Absent for any region never queried; only downgrades coverage
    to ``degraded_acknowledged`` when the matching low-overlap warning code
    is also present in ``warnings_acknowledged``.
    """
    hashes = file_hashes or {}
    ratios = key_overlap_ratios or {}
    members: list[ResolvedMember] = []
    for member in state.member_reviews:
        sheets: list[ResolvedSheet] = []
        pairs, _added, _removed = effective_sheet_pairing(member)
        baseline_by_current = {curr: base for base, curr in pairs}
        for sheet in member.current_sheets:
            baseline_name = baseline_by_current.get(sheet.sheet_name)
            regions = tuple(
                _resolved_region(
                    region,
                    _matching_baseline_region(member, sheet, index),
                    key_overlap_ratio=ratios.get(region.region_id),
                    low_overlap_acknowledged=(
                        low_key_overlap_warning_code(
                            member.member_id, sheet.sheet_name, region.region_id
                        )
                        in warnings_acknowledged
                    ),
                )
                for index, region in enumerate(sheet.regions)
            )
            selectors = tuple(
                ResolvedSelector(
                    selector_id=selector.selector_id,
                    baseline_cell=(selector.baseline_cell or selector.cell)
                    if baseline_name
                    else None,
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
        suffix = "" if member.member_id == "primary" else f":{member.member_id}"
        members.append(
            ResolvedMember(
                member_id=member.member_id,
                baseline_source_sha256=hashes.get(f"baseline_excel{suffix}"),
                current_source_sha256=hashes.get(f"current_excel{suffix}"),
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


def _resolved_region_to_contract(region: ResolvedRegion) -> LogicalRegionContract | None:
    """One resolved region reconstructed as a durable logical region, or
    ``None`` when it cannot be reconstructed faithfully.

    ``mode == "excluded"`` is never reconstructed: ``ResolvedRegion`` only
    ever carries a bare disclosure string (``degraded_reason``), never the
    exclusion's required expiry date, so inventing one here would
    fabricate data this function has never actually seen -- the analyst
    can always re-apply an exclusion with its own reason/expiry in the
    workspace before saving. The same reasoning excludes any column whose
    ``comparison_policy == "ignore"``.
    """
    if region.mode == "excluded":
        return None
    outer_range = region.current_outer_range or region.current_data_range
    if outer_range is None:
        return None
    min_row, min_col, _max_row, _max_col = parse_a1_range(outer_range)
    anchor_cell = f"{get_column_letter(min_col)}{min_row}"
    columns = tuple(
        LogicalColumnContract(
            column_id=column.column_id,
            alignment_role=column.alignment_role,
            comparison_policy=column.comparison_policy,
            trim_outer_whitespace=column.trim_outer_whitespace,
        )
        for column in region.columns
        if column.comparison_policy != "ignore"
    )
    return LogicalRegionContract(
        region_id=region.region_id,
        mode=region.mode,
        header_intent=region.header_intent,
        anchor_cell=anchor_cell,
        preferred_current_range=outer_range,
        preferred_first_data_row=(
            region.current_first_data_row
            if region.header_intent == "first_data_row"
            else None
        ),
        columns=columns,
        blank_key_policy=region.blank_key_policy,
        duplicate_key_policy=region.duplicate_key_policy,
    )


def input_contract_from_resolved_configuration(
    resolved: ResolvedInputConfigurationV1,
) -> WorkbookInputContract:
    """Reconstruct a durable ``WorkbookInputContract`` from a completed
    run's own resolved configuration -- the inverse of
    ``build_resolved_configuration``, used for POST-RUN profile saving
    (Step 11's "a successful temporary run can later create or update a
    named profile" criterion) rather than the live workspace session.
    """
    members: list[LogicalMemberContract] = []
    for member in resolved.members:
        sheets: list[LogicalSheetContract] = []
        for sheet in member.sheets:
            regions = tuple(
                contract
                for region in sheet.regions
                if (contract := _resolved_region_to_contract(region)) is not None
            )
            selectors = tuple(
                SelectorPrerequisiteContract(
                    selector_id=selector.selector_id,
                    label=selector.selector_id,
                    owner_sheet_id=sheet.sheet_id,
                    preferred_current_cell=selector.current_cell,
                )
                for selector in sheet.selectors
                if selector.current_cell is not None
            )
            sheets.append(
                LogicalSheetContract(
                    sheet_id=sheet.sheet_id,
                    preferred_sheet_name=sheet.current_sheet_name,
                    regions=regions,
                    selectors=selectors,
                )
            )
        members.append(
            LogicalMemberContract(member_id=member.member_id, sheets=tuple(sheets))
        )
    return WorkbookInputContract(members=tuple(members))


def summarize_contract_promotion(
    existing: WorkbookInputContract | None,
    new: WorkbookInputContract,
) -> tuple[str, ...]:
    """Plain-English lines describing what saving ``new`` over ``existing``
    would change (Step 11's "only after diff confirmation" criterion).
    Coarse -- sheet presence/count, not a full field diff -- proportionate
    to a one-time confirmation prompt, not a forensic report. Never a
    filesystem path, sheet NAME, or cell value; only counts and stable
    logical ids already present on the contracts being compared.
    """
    if existing is None:
        sheet_count = sum(len(member.sheets) for member in new.members)
        return (f"Creates a new profile with {sheet_count} configured sheet(s).",)
    existing_sheets = {
        sheet.sheet_id: sheet for member in existing.members for sheet in member.sheets
    }
    new_sheets = {sheet.sheet_id: sheet for member in new.members for sheet in member.sheets}
    added = sorted(set(new_sheets) - set(existing_sheets))
    removed = sorted(set(existing_sheets) - set(new_sheets))
    changed = sorted(
        sheet_id
        for sheet_id in set(existing_sheets) & set(new_sheets)
        if existing_sheets[sheet_id] != new_sheets[sheet_id]
    )
    lines: list[str] = []
    if added:
        lines.append(f"{len(added)} sheet(s) gain saved configuration.")
    if removed:
        lines.append(f"{len(removed)} sheet(s) lose their saved configuration.")
    if changed:
        lines.append(f"{len(changed)} sheet(s) have different saved configuration.")
    if not lines:
        lines.append("No configuration changes -- saving would be a no-op.")
    return tuple(lines)

