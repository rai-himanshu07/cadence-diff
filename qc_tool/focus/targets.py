"""Explicit finding-class matrix that produces private focus target seeds.

Targets are generated once per run, after triage has assigned finding ids and
after role hashes exist. Every finding class has an explicit rule; an unmapped
class or a producer that contradicts its rule raises, which leaves the run
successful with an empty sidecar rather than a guessed navigation target.
"""

from __future__ import annotations

import json
from collections.abc import Iterable, Mapping, Sequence
from enum import Enum, auto

from qc_tool.coverage import QCRunMode
from qc_tool.findings import Finding, FindingClass
from qc_tool.focus.locator import parse_excel_location, parse_qualified_location
from qc_tool.focus.model import (
    EMPTY_SIDECAR_JSON,
    FOCUS_SIDECAR_VERSION,
    FocusArtifact,
    FocusRole,
    FocusTargetSeed,
)


class FocusTargetContractError(RuntimeError):
    """A producer emitted evidence its focus rule cannot represent."""


class TargetRule(Enum):
    """How one finding class may become role-specific focus targets."""

    #: No MVP target: run-level, prose-only, or no structured locator.
    NONE = auto()
    #: Structured A1 locators, one seed per side the producer actually set.
    EXCEL_LOCATION = auto()
    #: Current workbook only; a same-workbook baseline marker is never a role.
    EXCEL_CURRENT_LOCATION = auto()
    EXCEL_SHEET_CURRENT = auto()
    EXCEL_SHEET_BASELINE = auto()
    #: Matched slide pair: each side seeded from its own typed index.
    PPT_MATCHED_SLIDE = auto()
    #: Single-deck evidence with a typed current slide index.
    PPT_CURRENT_SLIDE = auto()
    PPT_SLIDE_ADDED = auto()
    PPT_SLIDE_REMOVED = auto()
    #: Current Excel source cell plus the current deck slide showing the figure.
    CROSSCHECK_CURRENT = auto()


_MODE_ROLES: dict[QCRunMode, frozenset[FocusRole]] = {
    QCRunMode.CYCLE_COMPARISON: frozenset(FocusRole),
    QCRunMode.CURRENT_FILE_PREFLIGHT: frozenset(
        {FocusRole.CURRENT_EXCEL, FocusRole.CURRENT_PPT}
    ),
    QCRunMode.FINAL_PACKAGE: frozenset(
        {FocusRole.CURRENT_EXCEL, FocusRole.CURRENT_PPT}
    ),
}


FINDING_CLASS_RULES: dict[FindingClass, TargetRule] = {
    # --- Excel cell-level evidence -------------------------------------------
    FindingClass.VALUE_CHANGED: TargetRule.EXCEL_LOCATION,
    FindingClass.FORMULA_ERROR: TargetRule.EXCEL_LOCATION,
    FindingClass.FORMULA_HARDCODED: TargetRule.EXCEL_LOCATION,
    FindingClass.FORMULA_REMOVED: TargetRule.EXCEL_LOCATION,
    FindingClass.FORMULA_MISSING: TargetRule.EXCEL_LOCATION,
    FindingClass.FORMULA_CACHE_MISSING: TargetRule.EXCEL_LOCATION,
    FindingClass.FORMULA_NOT_EXTENDED: TargetRule.EXCEL_LOCATION,
    FindingClass.FORMULA_LOGIC_CHANGED: TargetRule.EXCEL_LOCATION,
    FindingClass.FORMULA_INCONSISTENT: TargetRule.EXCEL_LOCATION,
    FindingClass.CIRCULAR_REFERENCE: TargetRule.EXCEL_CURRENT_LOCATION,
    FindingClass.NUMBER_FORMAT_CHANGED: TargetRule.EXCEL_LOCATION,
    FindingClass.STYLE_CHANGED: TargetRule.EXCEL_LOCATION,
    FindingClass.CELL_COMMENT_CHANGED: TargetRule.EXCEL_LOCATION,
    # --- Excel axis and region evidence --------------------------------------
    FindingClass.ROW_DELETED: TargetRule.EXCEL_LOCATION,
    FindingClass.COLUMN_DELETED: TargetRule.EXCEL_LOCATION,
    FindingClass.ROW_INSERTED: TargetRule.EXCEL_LOCATION,
    FindingClass.COLUMN_INSERTED: TargetRule.EXCEL_LOCATION,
    FindingClass.ROW_KEY_CHANGED: TargetRule.EXCEL_LOCATION,
    FindingClass.COLUMN_KEY_CHANGED: TargetRule.EXCEL_LOCATION,
    FindingClass.ROW_GROWTH: TargetRule.EXCEL_LOCATION,
    FindingClass.COLUMN_GROWTH: TargetRule.EXCEL_LOCATION,
    FindingClass.REGION_UNPAIRED: TargetRule.EXCEL_LOCATION,
    # --- Excel sheet-level evidence ------------------------------------------
    FindingClass.SHEET_ADDED: TargetRule.EXCEL_SHEET_CURRENT,
    FindingClass.SHEET_REMOVED: TargetRule.EXCEL_SHEET_BASELINE,
    FindingClass.WORKBOOK_ADDED: TargetRule.NONE,
    FindingClass.WORKBOOK_REMOVED: TargetRule.NONE,
    # Focus never unhides a worksheet, so a hidden-state change has no target.
    FindingClass.HIDDEN_CHANGED: TargetRule.NONE,
    # --- Excel profile controls (current workbook) ---------------------------
    FindingClass.REQUIRED_VALUE_MISSING: TargetRule.EXCEL_CURRENT_LOCATION,
    FindingClass.DUPLICATE_KEY: TargetRule.EXCEL_CURRENT_LOCATION,
    FindingClass.NUMERIC_BOUND_VIOLATION: TargetRule.EXCEL_CURRENT_LOCATION,
    FindingClass.TIE_OUT_MISMATCH: TargetRule.EXCEL_CURRENT_LOCATION,
    # --- Excel object, workbook, and prose-only evidence ---------------------
    FindingClass.NAMED_RANGE_CHANGED: TargetRule.NONE,
    FindingClass.TABLE_STRUCTURE_CHANGED: TargetRule.NONE,
    FindingClass.DATA_VALIDATION_CHANGED: TargetRule.NONE,
    FindingClass.CONDITIONAL_FORMAT_CHANGED: TargetRule.NONE,
    FindingClass.CHART_STRUCTURE_CHANGED: TargetRule.NONE,
    FindingClass.CHART_PLOT_CHANGED: TargetRule.NONE,
    FindingClass.CHART_SERIES_CHANGED: TargetRule.NONE,
    FindingClass.CHART_AXIS_CHANGED: TargetRule.NONE,
    FindingClass.CHART_LEGEND_CHANGED: TargetRule.NONE,
    FindingClass.CHART_LABELS_CHANGED: TargetRule.NONE,
    FindingClass.CHART_GEOMETRY_CHANGED: TargetRule.NONE,
    FindingClass.PIVOT_SOURCE_CHANGED: TargetRule.NONE,
    FindingClass.ALIGNMENT_LOW_CONFIDENCE: TargetRule.NONE,
    FindingClass.FINDINGS_CAPPED: TargetRule.NONE,
    FindingClass.PERIOD_DUPLICATE: TargetRule.NONE,
    FindingClass.PERIOD_OUT_OF_ORDER: TargetRule.NONE,
    FindingClass.PERIOD_GAP: TargetRule.NONE,
    FindingClass.CALCULATION_MODE: TargetRule.NONE,
    FindingClass.EXTERNAL_LINK: TargetRule.NONE,
    FindingClass.ACTIVE_CONTENT: TargetRule.NONE,
    #: A VBA module has no cell or slide the analyst could be taken to.
    FindingClass.VBA_MODULE_CHANGED: TargetRule.NONE,
    #: Queries and connections are workbook-level, not addressable locations.
    FindingClass.POWER_QUERY_CHANGED: TargetRule.NONE,
    FindingClass.CONNECTION_CHANGED: TargetRule.NONE,
    FindingClass.EXTERNAL_CONNECTION: TargetRule.NONE,
    FindingClass.NAMED_RANGE_INVALID: TargetRule.NONE,
    FindingClass.CHART_REFERENCE_INVALID: TargetRule.NONE,
    FindingClass.CHART_LENGTH_MISMATCH: TargetRule.NONE,
    FindingClass.PIVOT_SOURCE_INVALID: TargetRule.NONE,
    FindingClass.HIDDEN_CONTENT: TargetRule.NONE,
    FindingClass.CONTROL_INVALID: TargetRule.NONE,
    FindingClass.WAIVER_EXPIRED: TargetRule.NONE,
    # --- PowerPoint ----------------------------------------------------------
    FindingClass.SLIDE_ADDED: TargetRule.PPT_SLIDE_ADDED,
    FindingClass.SLIDE_REMOVED: TargetRule.PPT_SLIDE_REMOVED,
    FindingClass.SLIDE_REORDERED: TargetRule.PPT_MATCHED_SLIDE,
    FindingClass.SLIDE_TEXT_CHANGED: TargetRule.PPT_MATCHED_SLIDE,
    FindingClass.TABLE_VALUE_CHANGED: TargetRule.PPT_MATCHED_SLIDE,
    FindingClass.CHART_VALUE_CHANGED: TargetRule.PPT_MATCHED_SLIDE,
    FindingClass.PPT_TABLE_STRUCTURE_CHANGED: TargetRule.PPT_MATCHED_SLIDE,
    FindingClass.PPT_CHART_STRUCTURE_CHANGED: TargetRule.PPT_MATCHED_SLIDE,
    FindingClass.PPT_CHART_PLOT_CHANGED: TargetRule.PPT_MATCHED_SLIDE,
    FindingClass.PPT_CHART_SERIES_CHANGED: TargetRule.PPT_MATCHED_SLIDE,
    FindingClass.PPT_CHART_AXIS_CHANGED: TargetRule.PPT_MATCHED_SLIDE,
    FindingClass.PPT_CHART_LEGEND_CHANGED: TargetRule.PPT_MATCHED_SLIDE,
    FindingClass.PPT_CHART_LABELS_CHANGED: TargetRule.PPT_MATCHED_SLIDE,
    FindingClass.PPT_SHAPE_GEOMETRY_CHANGED: TargetRule.PPT_MATCHED_SLIDE,
    FindingClass.PPT_MEDIA_CHANGED: TargetRule.PPT_MATCHED_SLIDE,
    FindingClass.PPT_DRAFT_TOKEN: TargetRule.PPT_CURRENT_SLIDE,
    FindingClass.PPT_EMPTY_SLIDE: TargetRule.PPT_CURRENT_SLIDE,
    FindingClass.PPT_TABLE_BLANK: TargetRule.PPT_CURRENT_SLIDE,
    FindingClass.PPT_CHART_LENGTH_MISMATCH: TargetRule.PPT_CURRENT_SLIDE,
    FindingClass.PPT_CHART_VALUE_MISSING: TargetRule.PPT_CURRENT_SLIDE,
    # Deck-wide evidence: no single slide is the subject.
    FindingClass.PPT_DUPLICATE_TITLE: TargetRule.NONE,
    FindingClass.PPT_REQUIRED_SLIDE_MISSING: TargetRule.NONE,
    FindingClass.PPT_PERIOD_INCONSISTENT: TargetRule.NONE,
    FindingClass.PPT_REPEATED_CLAIM_MISMATCH: TargetRule.NONE,
    # --- Crosscheck ----------------------------------------------------------
    FindingClass.CROSSCHECK_MISMATCH: TargetRule.CROSSCHECK_CURRENT,
    FindingClass.CROSSCHECK_UNRESOLVED: TargetRule.CROSSCHECK_CURRENT,
    FindingClass.PACKAGE_PERIOD_MISMATCH: TargetRule.NONE,
}


def _excel_seed(
    role: FocusRole,
    sheet: str | None,
    address: str | None,
    member_id: str = "primary",
) -> FocusTargetSeed | None:
    if not sheet:
        return None
    return FocusTargetSeed(
        artifact=FocusArtifact.EXCEL,
        role=role,
        member_id=member_id,
        sheet=sheet,
        address=address,
    )


def _ppt_seed(
    role: FocusRole,
    slide_index: int | None,
    shape_id: int | None = None,
) -> FocusTargetSeed | None:
    if slide_index is None:
        return None
    return FocusTargetSeed(
        artifact=FocusArtifact.PPT,
        role=role,
        slide_index=slide_index,
        shape_id=shape_id,
    )


def _seeds_for(finding: Finding, rule: TargetRule) -> list[FocusTargetSeed]:
    seeds: list[FocusTargetSeed | None] = []
    match rule:
        case TargetRule.NONE:
            return []
        case TargetRule.EXCEL_LOCATION:
            current = parse_excel_location(finding.location)
            baseline = parse_excel_location(finding.baseline_location)
            if current is not None:
                seeds.append(
                    _excel_seed(
                        FocusRole.CURRENT_EXCEL,
                        finding.sheet,
                        current,
                        finding.artifact_member,
                    )
                )
            if baseline is not None:
                seeds.append(
                    _excel_seed(
                        FocusRole.BASELINE_EXCEL,
                        finding.sheet,
                        baseline,
                        finding.artifact_member,
                    )
                )
        case TargetRule.EXCEL_CURRENT_LOCATION:
            current = parse_excel_location(finding.location)
            if current is not None:
                seeds.append(
                    _excel_seed(
                        FocusRole.CURRENT_EXCEL,
                        finding.sheet,
                        current,
                        finding.artifact_member,
                    )
                )
        case TargetRule.EXCEL_SHEET_CURRENT:
            seeds.append(
                _excel_seed(
                    FocusRole.CURRENT_EXCEL,
                    finding.sheet,
                    None,
                    finding.artifact_member,
                )
            )
        case TargetRule.EXCEL_SHEET_BASELINE:
            seeds.append(
                _excel_seed(
                    FocusRole.BASELINE_EXCEL,
                    finding.sheet,
                    None,
                    finding.artifact_member,
                )
            )
        case TargetRule.PPT_MATCHED_SLIDE:
            seeds.append(
                _ppt_seed(
                    FocusRole.CURRENT_PPT,
                    finding.slide_index,
                    finding.focus_shape_id,
                )
            )
            seeds.append(
                _ppt_seed(
                    FocusRole.BASELINE_PPT,
                    finding.baseline_slide_index,
                    finding.baseline_focus_shape_id,
                )
            )
        case TargetRule.PPT_CURRENT_SLIDE:
            seeds.append(
                _ppt_seed(
                    FocusRole.CURRENT_PPT,
                    finding.slide_index,
                    finding.focus_shape_id,
                )
            )
        case TargetRule.PPT_SLIDE_ADDED:
            if finding.baseline_slide_index is not None:
                raise FocusTargetContractError(
                    "an added slide cannot carry a baseline slide index"
                )
            seeds.append(_ppt_seed(FocusRole.CURRENT_PPT, finding.slide_index))
        case TargetRule.PPT_SLIDE_REMOVED:
            if finding.slide_index is not None:
                raise FocusTargetContractError(
                    "a removed slide cannot carry a current slide index"
                )
            seeds.append(_ppt_seed(FocusRole.BASELINE_PPT, finding.baseline_slide_index))
        case TargetRule.CROSSCHECK_CURRENT:
            qualified = parse_qualified_location(finding.location)
            if qualified is not None:
                sheet, address = qualified
                seeds.append(
                    _excel_seed(
                        FocusRole.CURRENT_EXCEL,
                        sheet,
                        address,
                        finding.artifact_member,
                    )
                )
            seeds.append(
                _ppt_seed(
                    FocusRole.CURRENT_PPT,
                    finding.slide_index,
                    finding.focus_shape_id,
                )
            )
    return [seed for seed in seeds if seed is not None]


def build_focus_targets(
    findings: Sequence[Finding],
    *,
    mode: QCRunMode,
    file_hashes: Mapping[str, str],
) -> dict[str, tuple[FocusTargetSeed, ...]]:
    """Generate the private sidecar payload for one recorded run.

    Raises ``FocusTargetContractError`` when a finding has no assigned id, its
    class has no rule, or its producer contradicts that rule.
    """
    targets: dict[str, tuple[FocusTargetSeed, ...]] = {}
    for finding_id, seeds in _iter_focus_seeds(
        findings, mode=mode, file_hashes=file_hashes
    ):
        targets[finding_id] = seeds
    return targets


def encode_focus_targets_streaming(
    findings: Iterable[Finding],
    *,
    mode: QCRunMode,
    file_hashes: Mapping[str, str],
) -> str:
    """One-pass sidecar encoding that never holds every seed at once.

    Byte-identical to ``encode_focus_targets(build_focus_targets(...))`` while
    finding ids keep their four-digit zero padding (every existing fixture):
    there, production order equals the old ``sorted()`` order. Past 9,999
    findings the entries stay in production order — the decoded mapping is
    order-insensitive, and no stored row predates this encoder at that scale.
    """
    fragments: list[str] = []
    for finding_id, seeds in _iter_focus_seeds(
        findings, mode=mode, file_hashes=file_hashes
    ):
        encoded_seeds = json.dumps([seed.model_dump(mode="json") for seed in seeds])
        fragments.append(f"{json.dumps(finding_id)}: {encoded_seeds}")
    if not fragments:
        return EMPTY_SIDECAR_JSON
    body = ", ".join(fragments)
    return f'{{"version": {FOCUS_SIDECAR_VERSION}, "targets": {{{body}}}}}'


def _rule_for(finding: Finding) -> TargetRule:
    """A population has no single-cell focus target regardless of class."""
    if finding.population is not None:
        return TargetRule.NONE
    rule = FINDING_CLASS_RULES.get(finding.finding_class)
    if rule is None:
        raise FocusTargetContractError("finding class has no focus target rule")
    return rule


def focus_seeds_for_finding(
    finding: Finding,
    *,
    mode: QCRunMode,
    file_hashes: Mapping[str, str],
) -> tuple[FocusTargetSeed, ...]:
    """Role seeds for one triaged finding; the per-block storage unit.

    Raises ``FocusTargetContractError`` exactly like ``build_focus_targets``.
    """
    if not finding.finding_id:
        raise FocusTargetContractError("focus targets need triaged finding ids")
    rule = _rule_for(finding)
    return tuple(
        seed
        for seed in _seeds_for(finding, rule)
        if seed.role in _MODE_ROLES[mode] and file_hashes.get(seed.role_key)
    )


def _iter_focus_seeds(
    findings: Iterable[Finding],
    *,
    mode: QCRunMode,
    file_hashes: Mapping[str, str],
) -> Iterable[tuple[str, tuple[FocusTargetSeed, ...]]]:
    for finding in findings:
        if not finding.finding_id:
            raise FocusTargetContractError("focus targets need triaged finding ids")
        rule = _rule_for(finding)
        seeds = tuple(
            seed
            for seed in _seeds_for(finding, rule)
            if seed.role in _MODE_ROLES[mode] and file_hashes.get(seed.role_key)
        )
        if seeds:
            yield finding.finding_id, seeds
