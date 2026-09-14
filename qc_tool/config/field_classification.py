"""Closed field-classification registry for Re-QC scope-comparability.

Every leaf field of ``DeliverableProfile`` (including the new saved
``WorkbookInputContract``) and every field of ``qc_tool.runqueue.RunRequest``
must be classified as exactly one of:

- ``scope_semantics``: can add, remove, suppress, reclassify, reprioritize,
  or otherwise change retained evidence for a logical scope. A change here
  makes affected scopes not comparable for Re-QC delta/carry-forward
  (``qc_tool.history.config_compatibility``, Step 4).
- ``output_representation``: affects only atomic-vs-population rendering.
  Gated by the existing, separate output-mode compatibility contract
  (``qc_tool.config.profile.resolve_output_policy``); classified here only
  so the drift guard has a place to put it.
- ``presentation_only``: affects only display (aliases, labels, names) and
  never evidence identity or comparability.
- ``operational_only``: pure execution/plumbing controls that cannot alter
  retained evidence for a completed run (identifiers, paths, workload
  permission gates that only affect whether a run proceeds at all).

``discover_profile_fields``/``discover_run_request_fields`` walk the live
model trees so a newly added field cannot silently escape classification;
a drift test in ``tests/`` fails the moment a discovered path is missing
from the registry (or a registry entry no longer corresponds to any real
field).
"""

from __future__ import annotations

import dataclasses
from typing import Literal, get_args, get_origin

from pydantic import BaseModel

FieldCategory = Literal[
    "scope_semantics", "output_representation", "presentation_only", "operational_only"
]

#: Safety bound against any unforeseen self-referential model shape; every
#: real model tree here is well under this depth.
_MAX_WALK_DEPTH = 12


def unwrap_optional(annotation: object) -> object:
    if get_origin(annotation) is not None and type(None) in get_args(annotation):
        remaining = [arg for arg in get_args(annotation) if arg is not type(None)]
        if len(remaining) == 1:
            return remaining[0]
    return annotation


def model_item_type(annotation: object) -> tuple[type[BaseModel], str] | None:
    """If ``annotation`` is a container of exactly one BaseModel subtype
    (``list[M]``, ``tuple[M, ...]``, or ``dict[K, M]``), return
    ``(M, suffix)`` where ``suffix`` is ``"[]"`` for a list/tuple and
    ``"{}"`` for a dict (so a discovered path names its own shape).
    """
    annotation = unwrap_optional(annotation)
    origin = get_origin(annotation)
    if origin is None:
        return None
    args = [arg for arg in get_args(annotation) if arg is not Ellipsis]
    model_args = [
        arg for arg in args if isinstance(arg, type) and issubclass(arg, BaseModel)
    ]
    if len(model_args) != 1:
        return None
    suffix = "{}" if origin is dict else "[]"
    return model_args[0], suffix


def _walk_model(
    model_cls: type[BaseModel], *, prefix: str, depth: int, out: set[str]
) -> None:
    if depth > _MAX_WALK_DEPTH:
        raise RecursionError(f"field walk exceeded max depth at {prefix!r}")
    for field_name, field_info in model_cls.model_fields.items():
        annotation = field_info.annotation
        path = f"{prefix}.{field_name}" if prefix else field_name
        direct = unwrap_optional(annotation)
        if isinstance(direct, type) and issubclass(direct, BaseModel):
            _walk_model(direct, prefix=path, depth=depth + 1, out=out)
            continue
        nested = model_item_type(annotation)
        if nested is not None:
            nested_cls, suffix = nested
            _walk_model(nested_cls, prefix=f"{path}{suffix}", depth=depth + 1, out=out)
            continue
        out.add(path)


def discover_profile_fields() -> set[str]:
    """Every leaf field path reachable from ``DeliverableProfile``."""
    from qc_tool.config.profile import DeliverableProfile

    out: set[str] = set()
    _walk_model(DeliverableProfile, prefix="", depth=0, out=out)
    return out


def discover_run_request_fields() -> set[str]:
    """Every field of ``RunRequest``.

    ``RunRequest.profile`` is deliberately a raw ``dict[str, Any]`` snapshot
    of a ``DeliverableProfile`` (queue payloads are primitive-only); its
    content is governed by ``discover_profile_fields`` and is not
    re-enumerated here. ``package_manifest`` is likewise an untyped
    primitive payload classified as one leaf.
    """
    from qc_tool.runqueue import RunRequest

    out: set[str] = set()
    for field in dataclasses.fields(RunRequest):
        out.add(field.name)
    return out


#: dotted-path (from the DeliverableProfile root) -> category. Governs only
#: saved *policy* fields (this registry). Per-run *resolved bindings*
#: (baseline/current sheet names, ranges, column letters) are compared
#: directly as `ResolvedInputConfigurationV1` equality by the Step 4
#: compatibility service, not through this registry -- a region's
#: `anchor_cell`/`preferred_current_range` here are only hints for finding
#: the region again, classified `presentation_only`/pointer-like for that
#: reason, not because a rebinding can never matter (a rebinding shows up as
#: a changed *resolved* sheet/region, which the Step 4 service compares
#: separately).
PROFILE_FIELD_CLASSIFICATION: dict[str, FieldCategory] = {
    "name": "presentation_only",
    "contract_id": "operational_only",
    "description": "presentation_only",
    "tolerance.absolute": "scope_semantics",
    "tolerance.relative": "scope_semantics",
    "restatement_windows.week": "scope_semantics",
    "restatement_windows.month": "scope_semantics",
    "restatement_windows.quarter": "scope_semantics",
    "severity": "scope_semantics",
    "materiality_severity": "scope_semantics",
    # -- excel: legacy physical-name fields --------------------------------
    "excel.ignore_sheets": "scope_semantics",
    "excel.formula_engine": "operational_only",
    "excel.sheets{}.ignore": "scope_semantics",
    "excel.sheets{}.ignore_ranges": "scope_semantics",
    "excel.sheets{}.refresh_ranges": "scope_semantics",
    "excel.sheets{}.acceptance_bands[].cell_range": "scope_semantics",
    "excel.sheets{}.acceptance_bands[].absolute": "scope_semantics",
    "excel.sheets{}.acceptance_bands[].relative": "scope_semantics",
    "excel.sheets{}.regions[].cell_range": "scope_semantics",
    "excel.sheets{}.regions[].orientation": "scope_semantics",
    "excel.sheets{}.regions[].header_row": "scope_semantics",
    "excel.sheets{}.regions[].key_column": "scope_semantics",
    "excel.sheets{}.cadence_bands[].cell_range": "scope_semantics",
    "excel.sheets{}.cadence_bands[].kind": "scope_semantics",
    "excel.sheets{}.availability_rules[].name": "presentation_only",
    "excel.sheets{}.availability_rules[].cell_range": "scope_semantics",
    "excel.sheets{}.availability_rules[].period_range": "scope_semantics",
    "excel.sheets{}.availability_rules[].required_through": "scope_semantics",
    "excel.sheets{}.availability_rules[].allow_blank_after": "scope_semantics",
    "excel.sheets{}.availability_rules[].role": "scope_semantics",
    "excel.sheets{}.chart_windows": "scope_semantics",
    "excel.sheets{}.row_identity_rules[].anchor_cell": "scope_semantics",
    "excel.sheets{}.row_identity_rules[].header_row": "scope_semantics",
    "excel.sheets{}.row_identity_rules[].identity_columns": "scope_semantics",
    "excel.sheets{}.row_identity_rules[].ordinal_columns": "scope_semantics",
    "excel.sheets{}.row_identity_rules[].duplicate_policy": "scope_semantics",
    "excel.sheets{}.row_identity_rules[].footer_row": "scope_semantics",
    "excel.sheets{}.row_identity_rules[].baseline_header_row": "scope_semantics",
    "excel.sheets{}.row_identity_rules[].baseline_footer_row": "scope_semantics",
    "excel.sheets{}.row_identity_rules[].exact_typed_equality": "scope_semantics",
    "excel.sheets{}.row_identity_rules[].trim_identity_whitespace": "scope_semantics",
    "excel.controls.required_ranges[].name": "presentation_only",
    "excel.controls.required_ranges[].sheet": "scope_semantics",
    "excel.controls.required_ranges[].cell_range": "scope_semantics",
    "excel.controls.unique_ranges[].name": "presentation_only",
    "excel.controls.unique_ranges[].sheet": "scope_semantics",
    "excel.controls.unique_ranges[].cell_range": "scope_semantics",
    "excel.controls.unique_ranges[].skip_header": "scope_semantics",
    "excel.controls.numeric_bounds[].name": "presentation_only",
    "excel.controls.numeric_bounds[].sheet": "scope_semantics",
    "excel.controls.numeric_bounds[].cell_range": "scope_semantics",
    "excel.controls.numeric_bounds[].minimum": "scope_semantics",
    "excel.controls.numeric_bounds[].maximum": "scope_semantics",
    "excel.controls.tie_outs[].name": "presentation_only",
    "excel.controls.tie_outs[].target": "scope_semantics",
    "excel.controls.tie_outs[].components": "scope_semantics",
    "excel.controls.tie_outs[].terms[].reference": "scope_semantics",
    "excel.controls.tie_outs[].terms[].operation": "scope_semantics",
    "excel.controls.tie_outs[].absolute_tolerance": "scope_semantics",
    "excel.controls.tie_outs[].relative_tolerance": "scope_semantics",
    "excel.comparison_prerequisites[].name": "presentation_only",
    "excel.comparison_prerequisites[].sheet": "scope_semantics",
    "excel.comparison_prerequisites[].cell": "scope_semantics",
    "excel.members{}.ignore_sheets": "scope_semantics",
    "excel.members{}.formula_engine": "operational_only",
    "excel.members{}.sheets{}.ignore": "scope_semantics",
    "excel.members{}.sheets{}.ignore_ranges": "scope_semantics",
    "excel.members{}.sheets{}.refresh_ranges": "scope_semantics",
    "excel.members{}.sheets{}.acceptance_bands[].cell_range": "scope_semantics",
    "excel.members{}.sheets{}.acceptance_bands[].absolute": "scope_semantics",
    "excel.members{}.sheets{}.acceptance_bands[].relative": "scope_semantics",
    "excel.members{}.sheets{}.regions[].cell_range": "scope_semantics",
    "excel.members{}.sheets{}.regions[].orientation": "scope_semantics",
    "excel.members{}.sheets{}.regions[].header_row": "scope_semantics",
    "excel.members{}.sheets{}.regions[].key_column": "scope_semantics",
    "excel.members{}.sheets{}.cadence_bands[].cell_range": "scope_semantics",
    "excel.members{}.sheets{}.cadence_bands[].kind": "scope_semantics",
    "excel.members{}.sheets{}.availability_rules[].name": "presentation_only",
    "excel.members{}.sheets{}.availability_rules[].cell_range": "scope_semantics",
    "excel.members{}.sheets{}.availability_rules[].period_range": "scope_semantics",
    "excel.members{}.sheets{}.availability_rules[].required_through": "scope_semantics",
    "excel.members{}.sheets{}.availability_rules[].allow_blank_after": "scope_semantics",
    "excel.members{}.sheets{}.availability_rules[].role": "scope_semantics",
    "excel.members{}.sheets{}.chart_windows": "scope_semantics",
    "excel.members{}.sheets{}.row_identity_rules[].anchor_cell": "scope_semantics",
    "excel.members{}.sheets{}.row_identity_rules[].header_row": "scope_semantics",
    "excel.members{}.sheets{}.row_identity_rules[].identity_columns": "scope_semantics",
    "excel.members{}.sheets{}.row_identity_rules[].ordinal_columns": "scope_semantics",
    "excel.members{}.sheets{}.row_identity_rules[].duplicate_policy": "scope_semantics",
    "excel.members{}.sheets{}.row_identity_rules[].footer_row": "scope_semantics",
    "excel.members{}.sheets{}.row_identity_rules[].baseline_header_row": "scope_semantics",
    "excel.members{}.sheets{}.row_identity_rules[].baseline_footer_row": "scope_semantics",
    "excel.members{}.sheets{}.row_identity_rules[].exact_typed_equality": "scope_semantics",
    "excel.members{}.sheets{}.row_identity_rules[].trim_identity_whitespace": "scope_semantics",
    "excel.members{}.controls.required_ranges[].name": "presentation_only",
    "excel.members{}.controls.required_ranges[].sheet": "scope_semantics",
    "excel.members{}.controls.required_ranges[].cell_range": "scope_semantics",
    "excel.members{}.controls.unique_ranges[].name": "presentation_only",
    "excel.members{}.controls.unique_ranges[].sheet": "scope_semantics",
    "excel.members{}.controls.unique_ranges[].cell_range": "scope_semantics",
    "excel.members{}.controls.unique_ranges[].skip_header": "scope_semantics",
    "excel.members{}.controls.numeric_bounds[].name": "presentation_only",
    "excel.members{}.controls.numeric_bounds[].sheet": "scope_semantics",
    "excel.members{}.controls.numeric_bounds[].cell_range": "scope_semantics",
    "excel.members{}.controls.numeric_bounds[].minimum": "scope_semantics",
    "excel.members{}.controls.numeric_bounds[].maximum": "scope_semantics",
    "excel.members{}.controls.tie_outs[].name": "presentation_only",
    "excel.members{}.controls.tie_outs[].target": "scope_semantics",
    "excel.members{}.controls.tie_outs[].components": "scope_semantics",
    "excel.members{}.controls.tie_outs[].terms[].reference": "scope_semantics",
    "excel.members{}.controls.tie_outs[].terms[].operation": "scope_semantics",
    "excel.members{}.controls.tie_outs[].absolute_tolerance": "scope_semantics",
    "excel.members{}.controls.tie_outs[].relative_tolerance": "scope_semantics",
    "excel.members{}.comparison_prerequisites[].name": "presentation_only",
    "excel.members{}.comparison_prerequisites[].sheet": "scope_semantics",
    "excel.members{}.comparison_prerequisites[].cell": "scope_semantics",
    # -- ppt ----------------------------------------------------------------
    "ppt.slide_pins": "scope_semantics",
    "ppt.match_threshold": "scope_semantics",
    "ppt.chart_windows": "scope_semantics",
    "ppt.availability_rules[].name": "presentation_only",
    "ppt.availability_rules[].slide": "scope_semantics",
    "ppt.availability_rules[].scope": "scope_semantics",
    "ppt.availability_rules[].element": "scope_semantics",
    "ppt.availability_rules[].series": "scope_semantics",
    "ppt.availability_rules[].required_through": "scope_semantics",
    "ppt.availability_rules[].allow_blank_after": "scope_semantics",
    "ppt.availability_rules[].role": "scope_semantics",
    "ppt.required_slides": "scope_semantics",
    "ppt.draft_tokens": "scope_semantics",
    # -- crosscheck -----------------------------------------------------------
    "crosscheck.mappings[].slide": "scope_semantics",
    "crosscheck.mappings[].line_skeleton": "scope_semantics",
    "crosscheck.mappings[].figure_index": "scope_semantics",
    "crosscheck.mappings[].label": "presentation_only",
    "crosscheck.mappings[].source_sheet": "scope_semantics",
    "crosscheck.mappings[].source_cell": "scope_semantics",
    "crosscheck.mappings[].source_member": "scope_semantics",
    "crosscheck.max_candidates": "operational_only",
    # -- waivers ---------------------------------------------------------
    "waivers[].finding_class": "scope_semantics",
    "waivers[].reason": "presentation_only",
    "waivers[].expires": "scope_semantics",
    "waivers[].sheet": "scope_semantics",
    "waivers[].slide": "scope_semantics",
    "waivers[].location": "scope_semantics",
    "waivers[].element": "scope_semantics",
    "waivers[].member": "scope_semantics",
    # -- review / output-representation policy ---------------------------
    "review_policy.version": "operational_only",
    "review_policy.populations.enabled": "output_representation",
    "review_policy.populations.threshold": "output_representation",
    "review_policy.populations.classes": "output_representation",
    "review_policy.populations.max_rectangles": "output_representation",
    "review_policy.populations.max_explicit_pairs": "output_representation",
    # -- saved logical input contract (plan-20260913) ------------------------
    "input_contract.version": "operational_only",
    "input_contract.members[].member_id": "operational_only",
    "input_contract.members[].alias": "presentation_only",
    "input_contract.members[].sheets[].sheet_id": "operational_only",
    "input_contract.members[].sheets[].alias": "presentation_only",
    "input_contract.members[].sheets[].preferred_sheet_name": "presentation_only",
    "input_contract.members[].sheets[].default_mode": "scope_semantics",
    "input_contract.members[].sheets[].exclusion.reason": "presentation_only",
    "input_contract.members[].sheets[].exclusion.expires_on": "scope_semantics",
    "input_contract.members[].sheets[].regions[].region_id": "operational_only",
    "input_contract.members[].sheets[].regions[].alias": "presentation_only",
    "input_contract.members[].sheets[].regions[].mode": "scope_semantics",
    "input_contract.members[].sheets[].regions[].header_intent": "scope_semantics",
    "input_contract.members[].sheets[].regions[].anchor_cell": "presentation_only",
    "input_contract.members[].sheets[].regions[].preferred_current_range": "presentation_only",
    "input_contract.members[].sheets[].regions[].preferred_first_data_row": "presentation_only",
    "input_contract.members[].sheets[].regions[].dynamic_expansion": "scope_semantics",
    "input_contract.members[].sheets[].regions[].blank_key_policy": "scope_semantics",
    "input_contract.members[].sheets[].regions[].duplicate_key_policy": "scope_semantics",
    "input_contract.members[].sheets[].regions[].exclusion.reason": "presentation_only",
    "input_contract.members[].sheets[].regions[].exclusion.expires_on": "scope_semantics",
    "input_contract.members[].sheets[].regions[].columns[].column_id": "operational_only",
    "input_contract.members[].sheets[].regions[].columns[].alias": "presentation_only",
    "input_contract.members[].sheets[].regions[].columns[].alignment_role": (
        "scope_semantics"
    ),
    "input_contract.members[].sheets[].regions[].columns[].comparison_policy": (
        "scope_semantics"
    ),
    "input_contract.members[].sheets[].regions[].columns[].trim_outer_whitespace": (
        "scope_semantics"
    ),
    "input_contract.members[].sheets[].regions[].columns[]"
    ".formula_backed_identity_acknowledged": "scope_semantics",
    "input_contract.members[].sheets[].regions[].columns[].exclusion.reason": (
        "presentation_only"
    ),
    "input_contract.members[].sheets[].regions[].columns[].exclusion.expires_on": (
        "scope_semantics"
    ),
    "input_contract.members[].sheets[].regions[].period_bands[].band_id": "operational_only",
    "input_contract.members[].sheets[].regions[].period_bands[].axis": (
        "scope_semantics"
    ),
    "input_contract.members[].sheets[].regions[].period_bands[].cadence_kind": (
        "scope_semantics"
    ),
    "input_contract.members[].sheets[].regions[].period_bands[]"
    ".preferred_current_range": "presentation_only",
    "input_contract.members[].sheets[].regions[].period_bands[].dynamic_expansion": (
        "scope_semantics"
    ),
    "input_contract.members[].sheets[].period_bands[].band_id": "operational_only",
    "input_contract.members[].sheets[].period_bands[].axis": "scope_semantics",
    "input_contract.members[].sheets[].period_bands[].cadence_kind": "scope_semantics",
    "input_contract.members[].sheets[].period_bands[].preferred_current_range": "presentation_only",
    "input_contract.members[].sheets[].period_bands[].dynamic_expansion": "scope_semantics",
    "input_contract.members[].sheets[].selectors[].selector_id": "operational_only",
    "input_contract.members[].sheets[].selectors[].label": "presentation_only",
    "input_contract.members[].sheets[].selectors[].owner_sheet_id": "operational_only",
    "input_contract.members[].sheets[].selectors[].preferred_current_cell": "presentation_only",
}

#: dotted-path -> category for ``RunRequest`` fields only. ``profile``'s
#: content is governed by ``PROFILE_FIELD_CLASSIFICATION`` and is not
#: re-enumerated here.
RUN_REQUEST_FIELD_CLASSIFICATION: dict[str, FieldCategory] = {
    "request_id": "operational_only",
    "work_dir": "operational_only",
    "mode": "scope_semantics",
    "profile_name": "presentation_only",
    "profile": "operational_only",
    "files": "operational_only",
    "display_files": "presentation_only",
    "requested_output_mode": "output_representation",
    "allow_large_workbooks": "operational_only",
    "allow_dependency_indexing": "scope_semantics",
    "acceptance_absolute": "scope_semantics",
    "acceptance_relative": "scope_semantics",
    "compare_sheets": "scope_semantics",
    "compare_slides": "scope_semantics",
    "rerun_of": "operational_only",
    "package_manifest": "scope_semantics",
    "compare_member_sheets": "scope_semantics",
    #: Content governed by ResolvedInputConfigurationV1 itself, not this
    #: registry; the digest is a pure derived value.
    "resolved_input_configuration": "operational_only",
    "resolved_input_digest": "operational_only",
    #: Job routing only (plan-20260913, Step 5); always "qc_run" for a real
    #: QC submission.
    "job_kind": "operational_only",
}


def unclassified_profile_fields() -> set[str]:
    """Fields the drift guard has not yet seen classified."""
    return discover_profile_fields() - set(PROFILE_FIELD_CLASSIFICATION)


def stale_profile_classifications() -> set[str]:
    """Registry entries that no longer correspond to a real field."""
    return set(PROFILE_FIELD_CLASSIFICATION) - discover_profile_fields()


def unclassified_run_request_fields() -> set[str]:
    return discover_run_request_fields() - set(RUN_REQUEST_FIELD_CLASSIFICATION)


def stale_run_request_classifications() -> set[str]:
    return set(RUN_REQUEST_FIELD_CLASSIFICATION) - discover_run_request_fields()


def classification_for(path: str) -> FieldCategory:
    if path in PROFILE_FIELD_CLASSIFICATION:
        return PROFILE_FIELD_CLASSIFICATION[path]
    if path in RUN_REQUEST_FIELD_CLASSIFICATION:
        return RUN_REQUEST_FIELD_CLASSIFICATION[path]
    raise KeyError(f"unclassified field: {path!r}")
