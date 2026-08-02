"""Excel data-validation and conditional-format matching, diff, and coverage."""

from __future__ import annotations

import re
from collections.abc import Callable
from typing import TypeVar

from openpyxl.utils.cell import range_boundaries

from qc_tool.coverage import CoverageState
from qc_tool.findings import Finding, FindingClass, FindingSubtype
from qc_tool.io.model import (
    ConditionalFormatDescriptor,
    DataValidationDescriptor,
    WorkbookSnapshot,
)

_T = TypeVar("_T")
_FORMULA_NUMBER_RE = re.compile(r"(?<![A-Za-z_$])\d+(?:\.\d+)?")


def _consume_unique_matches(
    baseline: list[_T],
    current: list[_T],
    key: Callable[[_T], object | None],
) -> list[tuple[_T, _T]]:
    baseline_keys: dict[object, list[_T]] = {}
    current_keys: dict[object, list[_T]] = {}
    for item in baseline:
        value = key(item)
        if value is not None:
            baseline_keys.setdefault(value, []).append(item)
    for item in current:
        value = key(item)
        if value is not None:
            current_keys.setdefault(value, []).append(item)
    pairs: list[tuple[_T, _T]] = []
    for value in sorted(baseline_keys.keys() & current_keys.keys(), key=repr):
        baseline_items = baseline_keys[value]
        current_items = current_keys[value]
        if len(baseline_items) != 1 or len(current_items) != 1:
            continue
        baseline_item = baseline_items[0]
        current_item = current_items[0]
        pairs.append((baseline_item, current_item))
        baseline.remove(baseline_item)
        current.remove(current_item)
    return pairs


def _consume_mutual_best_matches(
    baseline: list[_T],
    current: list[_T],
    score: Callable[[_T, _T], int | None],
) -> list[tuple[_T, _T]]:
    pairs: list[tuple[_T, _T]] = []
    while baseline and current:
        candidates = [
            (candidate_score, baseline_item, current_item)
            for baseline_item in baseline
            for current_item in current
            if (candidate_score := score(baseline_item, current_item)) is not None
        ]
        eligible: list[tuple[int, _T, _T]] = []
        for candidate_score, baseline_item, current_item in candidates:
            baseline_scores = [
                value
                for value, item, _ in candidates
                if item is baseline_item
            ]
            current_scores = [
                value
                for value, _, item in candidates
                if item is current_item
            ]
            if (
                candidate_score == max(baseline_scores)
                and baseline_scores.count(candidate_score) == 1
                and candidate_score == max(current_scores)
                and current_scores.count(candidate_score) == 1
            ):
                eligible.append((candidate_score, baseline_item, current_item))
        if not eligible:
            break
        used_baseline: set[int] = set()
        used_current: set[int] = set()
        matched = False
        for _, baseline_item, current_item in sorted(
            eligible,
            key=lambda item: (-item[0], repr(item[1]), repr(item[2])),
        ):
            if id(baseline_item) in used_baseline or id(current_item) in used_current:
                continue
            pairs.append((baseline_item, current_item))
            baseline.remove(baseline_item)
            current.remove(current_item)
            used_baseline.add(id(baseline_item))
            used_current.add(id(current_item))
            matched = True
        if not matched:
            break
    return pairs


def _range_boxes(ranges: tuple[str, ...]) -> list[tuple[int, int, int, int]]:
    boxes: list[tuple[int, int, int, int]] = []
    for cell_range in ranges:
        try:
            min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        except ValueError:
            continue
        if min_col is None or min_row is None or max_col is None or max_row is None:
            continue
        boxes.append((min_col, min_row, max_col, max_row))
    return boxes


def _target_affinity(
    baseline: tuple[str, ...],
    current: tuple[str, ...],
) -> int:
    affinity = 0
    for base_min_col, base_min_row, base_max_col, base_max_row in _range_boxes(
        baseline
    ):
        for curr_min_col, curr_min_row, curr_max_col, curr_max_row in _range_boxes(
            current
        ):
            overlap_columns = max(
                0,
                min(base_max_col, curr_max_col) - max(base_min_col, curr_min_col) + 1,
            )
            overlap_rows = max(
                0,
                min(base_max_row, curr_max_row) - max(base_min_row, curr_min_row) + 1,
            )
            if overlap_columns and overlap_rows:
                affinity = max(affinity, 100)
    return affinity


def _formula_shape(formula: str | None) -> str | None:
    if formula is None:
        return None
    return _FORMULA_NUMBER_RE.sub("#", formula.replace(" ", "").casefold())


def _validation_similarity(
    baseline: DataValidationDescriptor,
    current: DataValidationDescriptor,
) -> int | None:
    if (
        baseline.sheet != current.sheet
        or baseline.validation_type != current.validation_type
    ):
        return None
    target_affinity = _target_affinity(
        baseline.target_ranges,
        current.target_ranges,
    )
    if target_affinity == 0:
        return None
    return (
        target_affinity * 100
        + 20 * (baseline.operator == current.operator)
        + 15 * (_formula_shape(baseline.formula1) == _formula_shape(current.formula1))
        + 15 * (_formula_shape(baseline.formula2) == _formula_shape(current.formula2))
    )


def _conditional_similarity(
    baseline: ConditionalFormatDescriptor,
    current: ConditionalFormatDescriptor,
) -> int | None:
    if baseline.sheet != current.sheet or baseline.rule_type != current.rule_type:
        return None
    target_affinity = _target_affinity(
        baseline.target_ranges,
        current.target_ranges,
    )
    if target_affinity == 0:
        return None
    baseline_shapes = tuple(_formula_shape(formula) for formula in baseline.formulas)
    current_shapes = tuple(_formula_shape(formula) for formula in current.formulas)
    return (
        target_affinity * 100
        + 20 * (baseline.operator == current.operator)
        + 30 * (baseline_shapes == current_shapes)
    )


def _validation_condition(
    rule: DataValidationDescriptor,
) -> tuple[object, ...]:
    return (
        rule.sheet,
        rule.validation_type,
        rule.operator,
        rule.formula1,
        rule.formula2,
        bool(rule.allow_blank),
    )


def _validation_display(rule: DataValidationDescriptor) -> tuple[object, ...]:
    return (
        bool(rule.dropdown_suppressed),
        bool(rule.show_error_message),
        bool(rule.show_input_message),
        rule.error_style,
        rule.prompt_title,
        rule.prompt,
        rule.error_title,
        rule.error,
    )


def _validation_target(rule: DataValidationDescriptor) -> tuple[object, ...]:
    return rule.sheet, rule.target_ranges


def _match_validations(
    baseline: list[DataValidationDescriptor],
    current: list[DataValidationDescriptor],
) -> tuple[
    list[tuple[DataValidationDescriptor, DataValidationDescriptor]],
    list[DataValidationDescriptor],
    list[DataValidationDescriptor],
]:
    unmatched_baseline = list(baseline)
    unmatched_current = list(current)
    pairs = _consume_unique_matches(
        unmatched_baseline,
        unmatched_current,
        _validation_condition,
    )
    pairs.extend(
        _consume_unique_matches(
            unmatched_baseline,
            unmatched_current,
            _validation_target,
        )
    )
    pairs.extend(
        _consume_mutual_best_matches(
            unmatched_baseline,
            unmatched_current,
            _validation_similarity,
        )
    )
    return pairs, unmatched_baseline, unmatched_current


def _conditional_condition(
    rule: ConditionalFormatDescriptor,
) -> tuple[object, ...]:
    return (
        rule.sheet,
        rule.rule_type,
        rule.operator,
        rule.formulas,
        rule.text,
        rule.time_period,
        rule.rank,
        bool(rule.percent),
        bool(rule.bottom),
        bool(rule.above_average),
        bool(rule.equal_average),
        rule.standard_deviations,
    )


def _conditional_target(
    rule: ConditionalFormatDescriptor,
) -> tuple[object, ...]:
    return rule.sheet, rule.target_ranges


def _match_conditional_formats(
    baseline: list[ConditionalFormatDescriptor],
    current: list[ConditionalFormatDescriptor],
) -> tuple[
    list[tuple[ConditionalFormatDescriptor, ConditionalFormatDescriptor]],
    list[ConditionalFormatDescriptor],
    list[ConditionalFormatDescriptor],
]:
    unmatched_baseline = list(baseline)
    unmatched_current = list(current)
    pairs = _consume_unique_matches(
        unmatched_baseline,
        unmatched_current,
        _conditional_condition,
    )
    pairs.extend(
        _consume_unique_matches(
            unmatched_baseline,
            unmatched_current,
            _conditional_target,
        )
    )
    pairs.extend(
        _consume_mutual_best_matches(
            unmatched_baseline,
            unmatched_current,
            _conditional_similarity,
        )
    )
    return pairs, unmatched_baseline, unmatched_current


def _ranges_text(ranges: tuple[str, ...]) -> str:
    return " ".join(ranges)


def _validation_label(rule: DataValidationDescriptor) -> str:
    return (
        f"{rule.validation_type or 'validation'} rule "
        f"{rule.source_index + 1}"
    )


def _conditional_label(rule: ConditionalFormatDescriptor) -> str:
    return f"{rule.rule_type} rule {rule.source_index + 1}"


def _validation_event_key(rule: DataValidationDescriptor) -> str:
    return f"excel:{rule.sheet}:data-validation:{rule.source_id or rule.source_index}"


def _conditional_event_key(rule: ConditionalFormatDescriptor) -> str:
    return f"excel:{rule.sheet}:conditional-format:{rule.source_id or rule.source_index}"


def _priority_order_key(
    rule: ConditionalFormatDescriptor,
) -> tuple[int, int, int]:
    return (
        1 if rule.priority is None else 0,
        rule.priority if rule.priority is not None else 0,
        rule.source_index,
    )


def _reordered_pairs(
    pairs: list[tuple[ConditionalFormatDescriptor, ConditionalFormatDescriptor]],
) -> set[int]:
    """Paired rules whose relative evaluation order actually inverted.

    Renumbering caused only by additions or removals leaves every paired
    ordering intact and must not fan out into independent findings.
    """
    inverted: set[int] = set()
    for sheet in {current.sheet for _, current in pairs}:
        sheet_pairs = [pair for pair in pairs if pair[1].sheet == sheet]
        baseline_rank = {
            id(pair[1]): rank
            for rank, pair in enumerate(
                sorted(sheet_pairs, key=lambda pair: _priority_order_key(pair[0]))
            )
        }
        current_rank = {
            id(pair[1]): rank
            for rank, pair in enumerate(
                sorted(sheet_pairs, key=lambda pair: _priority_order_key(pair[1]))
            )
        }
        for index, (_, left) in enumerate(sheet_pairs):
            for _, right in sheet_pairs[index + 1 :]:
                before = baseline_rank[id(left)] < baseline_rank[id(right)]
                after = current_rank[id(left)] < current_rank[id(right)]
                if before != after:
                    inverted.update((id(left), id(right)))
    return inverted


def _validation_findings(
    baseline: DataValidationDescriptor,
    current: DataValidationDescriptor,
) -> list[Finding]:
    findings: list[Finding] = []
    label = _validation_label(current)
    event_key = _validation_event_key(current)
    if baseline.target_ranges != current.target_ranges:
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.DATA_VALIDATION_CHANGED,
                subtype=FindingSubtype.OBJECT_TARGET_CHANGED,
                event_key=event_key,
                sheet=current.sheet,
                element=label,
                baseline_value=_ranges_text(baseline.target_ranges),
                current_value=_ranges_text(current.target_ranges),
                message=f"data-validation {label} target coverage changed",
            )
        )
    if _validation_condition(baseline) != _validation_condition(current):
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.DATA_VALIDATION_CHANGED,
                subtype=FindingSubtype.OBJECT_CONDITION_CHANGED,
                event_key=event_key,
                sheet=current.sheet,
                element=label,
                baseline_value=(
                    f"{baseline.validation_type}|{baseline.operator}|"
                    f"{baseline.formula1}|{baseline.formula2}"
                ),
                current_value=(
                    f"{current.validation_type}|{current.operator}|"
                    f"{current.formula1}|{current.formula2}"
                ),
                message=f"data-validation {label} criteria changed",
            )
        )
    if _validation_display(baseline) != _validation_display(current):
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.DATA_VALIDATION_CHANGED,
                subtype=FindingSubtype.OBJECT_DISPLAY_CHANGED,
                event_key=event_key,
                sheet=current.sheet,
                element=label,
                message=f"data-validation {label} display text/settings changed",
            )
        )
    return findings


def _conditional_findings(
    baseline: ConditionalFormatDescriptor,
    current: ConditionalFormatDescriptor,
    *,
    order_changed: bool,
) -> list[Finding]:
    findings: list[Finding] = []
    label = _conditional_label(current)
    event_key = _conditional_event_key(current)
    if baseline.target_ranges != current.target_ranges:
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.CONDITIONAL_FORMAT_CHANGED,
                subtype=FindingSubtype.OBJECT_TARGET_CHANGED,
                event_key=event_key,
                sheet=current.sheet,
                element=label,
                baseline_value=_ranges_text(baseline.target_ranges),
                current_value=_ranges_text(current.target_ranges),
                message=f"conditional-format {label} target coverage changed",
            )
        )
    semantic_supported = baseline.semantic_supported and current.semantic_supported
    if (
        semantic_supported
        and _conditional_condition(baseline) != _conditional_condition(current)
    ):
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.CONDITIONAL_FORMAT_CHANGED,
                subtype=FindingSubtype.OBJECT_CONDITION_CHANGED,
                event_key=event_key,
                sheet=current.sheet,
                element=label,
                baseline_value=" | ".join(baseline.formulas),
                current_value=" | ".join(current.formulas),
                message=f"conditional-format {label} condition changed",
            )
        )
    if semantic_supported and order_changed:
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.CONDITIONAL_FORMAT_CHANGED,
                subtype=FindingSubtype.OBJECT_ORDER_CHANGED,
                event_key=event_key,
                sheet=current.sheet,
                element=label,
                baseline_value=(
                    None if baseline.priority is None else str(baseline.priority)
                ),
                current_value=(
                    None if current.priority is None else str(current.priority)
                ),
                message=(
                    f"conditional-format {label} evaluation order changed relative "
                    "to the other retained rules"
                ),
            )
        )
    if (
        semantic_supported
        and bool(baseline.stop_if_true) != bool(current.stop_if_true)
    ):
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.CONDITIONAL_FORMAT_CHANGED,
                subtype=FindingSubtype.OBJECT_SETTINGS_CHANGED,
                event_key=event_key,
                sheet=current.sheet,
                element=label,
                baseline_value=str(bool(baseline.stop_if_true)),
                current_value=str(bool(current.stop_if_true)),
                message=f"conditional-format {label} stop-if-true changed",
            )
        )
    if (
        baseline.style_supported
        and current.style_supported
        and baseline.style_key != current.style_key
    ):
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.CONDITIONAL_FORMAT_CHANGED,
                subtype=FindingSubtype.OBJECT_STYLE_CHANGED,
                event_key=event_key,
                sheet=current.sheet,
                element=label,
                baseline_value=baseline.style_key,
                current_value=current.style_key,
                message=f"conditional-format {label} differential style changed",
            )
        )
    return findings


def diff_interaction_rules(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
) -> list[Finding]:
    """Compare every captured data-validation and conditional-format rule."""
    if not baseline.interaction_rules_available or not current.interaction_rules_available:
        return []
    validation_pairs, removed_validations, added_validations = _match_validations(
        baseline.data_validations,
        current.data_validations,
    )
    findings = [
        Finding(
            artifact="excel",
            finding_class=FindingClass.DATA_VALIDATION_CHANGED,
            subtype=FindingSubtype.OBJECT_REMOVED,
            event_key=_validation_event_key(rule),
            sheet=rule.sheet,
            element=_validation_label(rule),
            baseline_value=_ranges_text(rule.target_ranges),
            message=f"data-validation {_validation_label(rule)} removed",
        )
        for rule in removed_validations
    ]
    findings.extend(
        Finding(
            artifact="excel",
            finding_class=FindingClass.DATA_VALIDATION_CHANGED,
            subtype=FindingSubtype.OBJECT_ADDED,
            event_key=_validation_event_key(rule),
            sheet=rule.sheet,
            element=_validation_label(rule),
            current_value=_ranges_text(rule.target_ranges),
            message=f"data-validation {_validation_label(rule)} added",
        )
        for rule in added_validations
    )
    for baseline_rule, current_rule in validation_pairs:
        findings.extend(_validation_findings(baseline_rule, current_rule))

    conditional_pairs, removed_conditionals, added_conditionals = (
        _match_conditional_formats(
            baseline.conditional_formats,
            current.conditional_formats,
        )
    )
    findings.extend(
        Finding(
            artifact="excel",
            finding_class=FindingClass.CONDITIONAL_FORMAT_CHANGED,
            subtype=FindingSubtype.OBJECT_REMOVED,
            event_key=_conditional_event_key(rule),
            sheet=rule.sheet,
            element=_conditional_label(rule),
            baseline_value=_ranges_text(rule.target_ranges),
            message=f"conditional-format {_conditional_label(rule)} removed",
        )
        for rule in removed_conditionals
    )
    findings.extend(
        Finding(
            artifact="excel",
            finding_class=FindingClass.CONDITIONAL_FORMAT_CHANGED,
            subtype=FindingSubtype.OBJECT_ADDED,
            event_key=_conditional_event_key(rule),
            sheet=rule.sheet,
            element=_conditional_label(rule),
            current_value=_ranges_text(rule.target_ranges),
            message=f"conditional-format {_conditional_label(rule)} added",
        )
        for rule in added_conditionals
    )
    reordered = _reordered_pairs(conditional_pairs)
    for baseline_rule, current_rule in conditional_pairs:
        findings.extend(
            _conditional_findings(
                baseline_rule,
                current_rule,
                order_changed=id(current_rule) in reordered,
            )
        )
    return findings


def interaction_rule_coverage(
    workbook: WorkbookSnapshot,
) -> tuple[CoverageState, str]:
    """Semantic coverage for data validation and conditional-format conditions."""
    if not workbook.interaction_rules_available:
        return CoverageState.UNAVAILABLE, workbook.interaction_rule_detail
    if not workbook.interaction_rules_supported:
        return CoverageState.DEGRADED, workbook.interaction_rule_detail
    return CoverageState.CHECKED, "Interaction-rule semantics captured"


def conditional_style_coverage(
    workbook: WorkbookSnapshot,
) -> tuple[CoverageState, str]:
    """Coverage for stable differential-format comparison."""
    if not workbook.interaction_rules_available:
        return CoverageState.UNAVAILABLE, workbook.conditional_format_style_detail
    if not workbook.conditional_format_styles_supported:
        return CoverageState.DEGRADED, workbook.conditional_format_style_detail
    return CoverageState.CHECKED, "Stable differential styles captured"
