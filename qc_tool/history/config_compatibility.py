"""Scope-semantics compatibility service for Re-QC delta and carry-forward
(plan-20260913-mode-aware-configuration-wizard.md, Step 4).

Extends -- never replaces -- the existing atomic/population output-mode
compatibility contract (``qc_tool.config.profile.resolve_output_policy`` /
``qc_tool.engine.output_representations_compatible``) with a SECOND,
independent gate: whether the underlying *scope semantics* -- which
evidence could exist at all for a logical scope -- match between two runs'
profile snapshots. A change here means findings for the affected scope(s)
are not comparable (never silently "resolved"); it says nothing about
output representation, which the existing contract still governs.

This is a genuine, general fix to today's Re-QC delta/carry-forward logic,
not a feature gated behind the new saved input contract: a profile-only
change today (loosen ``tolerance``, let a waiver expire, edit an
``availability_rule``) can already make a finding disappear or appear with
nothing in the underlying workbook changing, and neither existing delta
call site currently checks for that. It is active for every profile pair,
with or without a saved ``input_contract``.

Granularity, most specific first:
- Every finding is gated by a GLOBAL digest (``tolerance``,
  ``restatement_windows``, ``severity``, ``materiality_severity``,
  ``waivers``) that applies universally regardless of artifact.
- An "excel" finding is additionally gated by:
  - a member-wide excel digest (``ignore_sheets``, ``controls``,
    ``comparison_prerequisites`` -- fields that cannot be attributed to one
    sheet) for its ``artifact_member``.
  - a sheet-scoped digest. When the finding carries a ``logical_address``
    (Step 3), the scope key is its stable ``sheet_id`` and the digest comes
    from the SAVED ``input_contract`` entry for that id -- deliberately
    ignoring the legacy per-physical-name profile entry, so a confirmed
    rename alone (Step 3's whole point) never trips this gate. Otherwise
    the scope key is the finding's own physical sheet name and the digest
    comes from the legacy ``excel.sheets{}`` / ``excel.members{}.sheets{}``
    profile entry.
- A "ppt" finding is gated by one whole-deck ppt digest.
- A "crosscheck" finding is gated by one whole-profile crosscheck digest.

Every digest is built by pruning the real profile sub-model down to just
its ``scope_semantics``-classified leaves (via
``qc_tool.config.field_classification``), so presentation-only edits
(aliases, names, waiver reasons) never affect comparability -- only a
change that could add, remove, suppress, reclassify, or reprioritize
retained evidence does.
"""

from __future__ import annotations

import hashlib
import json
from collections.abc import Sequence
from dataclasses import dataclass, field

from pydantic import BaseModel

from qc_tool.config.field_classification import (
    PROFILE_FIELD_CLASSIFICATION,
    model_item_type,
    unwrap_optional,
)
from qc_tool.config.input_contract import LogicalMemberContract, LogicalSheetContract
from qc_tool.config.profile import DeliverableProfile, ExcelMemberProfile
from qc_tool.config.resolved_input import (
    ResolvedInputConfigurationV1,
    ResolvedMember,
    ResolvedRegion,
    ResolvedSheet,
)
from qc_tool.engine import FindingsDelta, compare_findings
from qc_tool.findings import Finding

#: Mirrors field_classification's own safety bound against an unforeseen
#: self-referential model shape.
_MAX_WALK_DEPTH = 12


def _prune_scope_semantics(
    instance: BaseModel, *, prefix: str, depth: int = 0
) -> dict[str, object] | None:
    """Nested structure containing only ``scope_semantics``-classified
    leaves of ``instance``, mirroring its own shape (dict keys = field
    names; list/tuple fields become a list of per-item pruned dicts; dict
    fields become a dict of per-item pruned values). ``None`` when nothing
    under this node is scope_semantics-classified -- an "absent" node and a
    "present but fully default" node are deliberately indistinguishable.
    """
    if depth > _MAX_WALK_DEPTH:
        raise RecursionError(f"scope-semantics prune exceeded max depth at {prefix!r}")
    result: dict[str, object] = {}
    for field_name, field_info in type(instance).model_fields.items():
        path = f"{prefix}.{field_name}" if prefix else field_name
        value = getattr(instance, field_name)
        annotation = field_info.annotation
        direct = unwrap_optional(annotation)
        if isinstance(direct, type) and issubclass(direct, BaseModel):
            if value is None:
                continue
            pruned = _prune_scope_semantics(value, prefix=path, depth=depth + 1)
            if pruned:
                result[field_name] = pruned
            continue
        nested = model_item_type(annotation)
        if nested is not None:
            _, suffix = nested
            container_path = f"{path}{suffix}"
            if suffix == "[]":
                pruned_items = [
                    pruned_item
                    for item in (value or ())
                    if isinstance(item, BaseModel)
                    and (
                        pruned_item := _prune_scope_semantics(
                            item, prefix=container_path, depth=depth + 1
                        )
                    )
                ]
                if pruned_items:
                    result[field_name] = pruned_items
            else:
                pruned_map = {
                    key: pruned_item
                    for key, item in (value or {}).items()
                    if isinstance(item, BaseModel)
                    and (
                        pruned_item := _prune_scope_semantics(
                            item, prefix=container_path, depth=depth + 1
                        )
                    )
                    is not None
                }
                if pruned_map:
                    result[field_name] = pruned_map
            continue
        if PROFILE_FIELD_CLASSIFICATION.get(path) == "scope_semantics":
            result[field_name] = value
    return result or None


def _digest(payload: object) -> str:
    canonical = json.dumps(payload, sort_keys=True, separators=(",", ":"), default=str)
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def _global_scope_digest(profile: DeliverableProfile) -> str:
    payload = {
        "tolerance": _prune_scope_semantics(profile.tolerance, prefix="tolerance"),
        "restatement_windows": _prune_scope_semantics(
            profile.restatement_windows, prefix="restatement_windows"
        ),
        "severity": (
            profile.severity
            if PROFILE_FIELD_CLASSIFICATION.get("severity") == "scope_semantics"
            else None
        ),
        "materiality_severity": (
            profile.materiality_severity
            if PROFILE_FIELD_CLASSIFICATION.get("materiality_severity") == "scope_semantics"
            else None
        ),
        "waivers": [
            pruned
            for waiver in profile.waivers
            if (pruned := _prune_scope_semantics(waiver, prefix="waivers[]"))
        ],
    }
    return _digest(payload)


def _member_profile(profile: DeliverableProfile, member_id: str) -> ExcelMemberProfile | None:
    if member_id == "primary":
        return profile.excel
    return profile.excel.members.get(member_id)


def _member_prefix(member_id: str) -> str:
    return "excel" if member_id == "primary" else "excel.members{}"


def _pruned_member(profile: DeliverableProfile, member_id: str) -> dict[str, object]:
    member = _member_profile(profile, member_id)
    if member is None:
        return {}
    return _prune_scope_semantics(member, prefix=_member_prefix(member_id)) or {}


def _member_wide_digest(profile: DeliverableProfile, member_id: str) -> str:
    pruned = _pruned_member(profile, member_id)
    payload = {k: v for k, v in pruned.items() if k not in ("sheets", "members")}
    return _digest(payload)


def _legacy_sheet_digest(profile: DeliverableProfile, member_id: str, sheet_name: str) -> str:
    pruned = _pruned_member(profile, member_id)
    sheets = pruned.get("sheets")
    entry = sheets.get(sheet_name) if isinstance(sheets, dict) else None
    return _digest(entry)


def _contract_sheet(
    profile: DeliverableProfile, member_id: str, sheet_id: str
) -> LogicalSheetContract | None:
    contract = profile.input_contract
    if contract is None:
        return None
    member: LogicalMemberContract | None = next(
        (m for m in contract.members if m.member_id == member_id), None
    )
    if member is None:
        return None
    return next((s for s in member.sheets if s.sheet_id == sheet_id), None)


def _contract_sheet_digest(profile: DeliverableProfile, member_id: str, sheet_id: str) -> str:
    sheet = _contract_sheet(profile, member_id, sheet_id)
    if sheet is None:
        return _digest(None)
    pruned = _prune_scope_semantics(
        sheet, prefix="input_contract.members[].sheets[]"
    )
    return _digest(pruned)


def _ppt_digest(profile: DeliverableProfile) -> str:
    return _digest(_prune_scope_semantics(profile.ppt, prefix="ppt"))


def _crosscheck_digest(profile: DeliverableProfile) -> str:
    return _digest(_prune_scope_semantics(profile.crosscheck, prefix="crosscheck"))


def _resolved_global_payload(
    resolved: ResolvedInputConfigurationV1,
) -> dict[str, object]:
    return {
        "version": resolved.version,
        "source": resolved.source,
        "mode": resolved.mode,
        "inspection_contract_version": resolved.inspection_contract_version,
        "warnings_acknowledged": sorted(resolved.warnings_acknowledged),
        "degradations_accepted": sorted(resolved.degradations_accepted),
        "exclusions_renewed": sorted(resolved.exclusions_renewed),
        "override_reasons": sorted(resolved.override_reasons),
    }


def _resolved_column_payload(column) -> dict[str, object]:
    return {
        "column_id": column.column_id,
        "alignment_role": column.alignment_role,
        "comparison_policy": column.comparison_policy,
        "trim_outer_whitespace": column.trim_outer_whitespace,
        "coverage": column.coverage,
    }


def _resolved_region_payload(region: ResolvedRegion) -> dict[str, object]:
    return {
        "region_id": region.region_id,
        "mode": region.mode,
        "header_intent": region.header_intent,
        "baseline_preamble_rows": region.baseline_preamble_rows,
        "current_preamble_rows": region.current_preamble_rows,
        "baseline_footer_rows": region.baseline_footer_rows,
        "current_footer_rows": region.current_footer_rows,
        "columns": sorted(
            (_resolved_column_payload(column) for column in region.columns),
            key=lambda item: str(item["column_id"]),
        ),
        "duplicate_key_policy": region.duplicate_key_policy,
        "blank_key_policy": region.blank_key_policy,
        "coverage": region.coverage,
    }


def _resolved_sheet_payload(sheet: ResolvedSheet) -> dict[str, object]:
    return {
        "sheet_id": sheet.sheet_id,
        "coverage": sheet.coverage,
        "regions": sorted(
            (_resolved_region_payload(region) for region in sheet.regions),
            key=lambda item: str(item["region_id"]),
        ),
        "selectors": sorted(
            (
                {
                    "selector_id": selector.selector_id,
                    "equal": selector.equal,
                    "baseline_formula_backed": selector.baseline_formula_backed,
                    "current_formula_backed": selector.current_formula_backed,
                    "coverage": selector.coverage,
                }
                for selector in sheet.selectors
            ),
            key=lambda item: str(item["selector_id"]),
        ),
    }


def _resolved_member(
    resolved: ResolvedInputConfigurationV1, member_id: str
) -> ResolvedMember | None:
    return next(
        (member for member in resolved.members if member.member_id == member_id),
        None,
    )


def _resolved_sheet_by_id(
    member: ResolvedMember, sheet_id: str
) -> ResolvedSheet | None:
    return next((sheet for sheet in member.sheets if sheet.sheet_id == sheet_id), None)


def _resolved_sheet_for_finding(
    member: ResolvedMember, finding: Finding
) -> ResolvedSheet | None:
    if finding.logical_address is not None:
        return _resolved_sheet_by_id(member, finding.logical_address.sheet_id)
    if finding.sheet is None:
        return None
    matches = [
        sheet
        for sheet in member.sheets
        if finding.sheet in {sheet.baseline_sheet_name, sheet.current_sheet_name}
    ]
    return matches[0] if len(matches) == 1 else None


@dataclass
class ConfigurationCompatibility:
    """Reusable scope-comparability checker for one pair of profile
    snapshots. Cheap to construct; call ``comparable(finding)`` once per
    finding on either side of a Re-QC delta or carry-forward candidate.
    """

    _previous: DeliverableProfile
    _current: DeliverableProfile
    _global_match: bool
    _previous_resolved: ResolvedInputConfigurationV1 | None = None
    _current_resolved: ResolvedInputConfigurationV1 | None = None
    _resolved_global_match: bool = True
    _member_wide_cache: dict[str, bool] = field(default_factory=dict)
    _sheet_cache: dict[tuple[str, str, str], bool] = field(default_factory=dict)
    _ppt_match_value: bool | None = None
    _crosscheck_match_value: bool | None = None

    def _member_wide_match(self, member_id: str) -> bool:
        cached = self._member_wide_cache.get(member_id)
        if cached is None:
            cached = _member_wide_digest(self._previous, member_id) == _member_wide_digest(
                self._current, member_id
            )
            self._member_wide_cache[member_id] = cached
        return cached

    def _sheet_match(self, finding: Finding) -> bool:
        member_id = finding.artifact_member
        if finding.logical_address is not None:
            sheet_id = finding.logical_address.sheet_id
            cache_key = ("id", member_id, sheet_id)
            cached = self._sheet_cache.get(cache_key)
            if cached is None:
                cached = _contract_sheet_digest(
                    self._previous, member_id, sheet_id
                ) == _contract_sheet_digest(self._current, member_id, sheet_id)
                self._sheet_cache[cache_key] = cached
            return cached
        if finding.sheet is None:
            return True
        cache_key = ("name", member_id, finding.sheet)
        cached = self._sheet_cache.get(cache_key)
        if cached is None:
            cached = _legacy_sheet_digest(
                self._previous, member_id, finding.sheet
            ) == _legacy_sheet_digest(self._current, member_id, finding.sheet)
            self._sheet_cache[cache_key] = cached
        return cached

    def _ppt_match(self) -> bool:
        if self._ppt_match_value is None:
            self._ppt_match_value = _ppt_digest(self._previous) == _ppt_digest(self._current)
        return self._ppt_match_value

    def _crosscheck_match(self) -> bool:
        if self._crosscheck_match_value is None:
            self._crosscheck_match_value = _crosscheck_digest(
                self._previous
            ) == _crosscheck_digest(self._current)
        return self._crosscheck_match_value

    def _resolved_match(self, finding: Finding) -> bool:
        previous = self._previous_resolved
        current = self._current_resolved
        if previous is None and current is None:
            return True
        if previous is None or current is None or not self._resolved_global_match:
            return False
        if previous.source == "legacy_default" and current.source == "legacy_default":
            return True
        if previous.source != "input_contract" or current.source != "input_contract":
            return False
        if finding.artifact != "excel":
            return True
        member_id = (
            finding.logical_address.member_id
            if finding.logical_address is not None
            else finding.artifact_member
        )
        previous_member = _resolved_member(previous, member_id)
        current_member = _resolved_member(current, member_id)
        if previous_member is None or current_member is None:
            return False
        previous_sheet = _resolved_sheet_for_finding(previous_member, finding)
        current_sheet = _resolved_sheet_for_finding(current_member, finding)
        if finding.sheet is not None or finding.logical_address is not None:
            if previous_sheet is None or current_sheet is None:
                return False
            if previous_sheet.sheet_id != current_sheet.sheet_id:
                return False
            logical_address = finding.logical_address
            if logical_address is not None and logical_address.region_id is not None:
                previous_region = next(
                    (
                        region
                        for region in previous_sheet.regions
                        if region.region_id == logical_address.region_id
                    ),
                    None,
                )
                current_region = next(
                    (
                        region
                        for region in current_sheet.regions
                        if region.region_id == logical_address.region_id
                    ),
                    None,
                )
                if previous_region is None or current_region is None:
                    return False
                return _digest(_resolved_region_payload(previous_region)) == _digest(
                    _resolved_region_payload(current_region)
                )
            return _digest(_resolved_sheet_payload(previous_sheet)) == _digest(
                _resolved_sheet_payload(current_sheet)
            )
        previous_payload = sorted(
            (_resolved_sheet_payload(sheet) for sheet in previous_member.sheets),
            key=lambda item: str(item["sheet_id"]),
        )
        current_payload = sorted(
            (_resolved_sheet_payload(sheet) for sheet in current_member.sheets),
            key=lambda item: str(item["sheet_id"]),
        )
        return _digest(previous_payload) == _digest(current_payload)

    def comparable(self, finding: Finding) -> bool:
        """Whether ``finding``'s logical scope has identical scope-affecting
        policy in both profile snapshots. A caller decides what "not
        comparable" means (exclude from a delta, route to an ambiguous
        carry-forward bucket, etc.) -- this only answers the policy
        question, never finding identity/location matching.
        """
        if not self._global_match:
            return False
        if not self._resolved_match(finding):
            return False
        if finding.artifact == "excel":
            return self._member_wide_match(finding.artifact_member) and self._sheet_match(
                finding
            )
        if finding.artifact == "ppt":
            return self._ppt_match()
        if finding.artifact == "crosscheck":
            return self._crosscheck_match()
        return True

    def fully_compatible(self) -> bool:
        """True only when the GLOBAL gate matches; a fast, coarse check for
        callers that want to skip per-finding filtering entirely on the
        (common) unchanged-profile path. ``False`` does not necessarily
        mean every scope differs -- ``comparable`` still narrows per finding.
        """
        return self._global_match


def configuration_compatible(
    previous: DeliverableProfile,
    current: DeliverableProfile,
    *,
    previous_resolved: ResolvedInputConfigurationV1 | None = None,
    current_resolved: ResolvedInputConfigurationV1 | None = None,
) -> ConfigurationCompatibility:
    """Build a ``ConfigurationCompatibility`` checker for one profile pair."""
    return ConfigurationCompatibility(
        _previous=previous,
        _current=current,
        _global_match=_global_scope_digest(previous) == _global_scope_digest(current),
        _previous_resolved=previous_resolved,
        _current_resolved=current_resolved,
        _resolved_global_match=(
            (previous_resolved is None and current_resolved is None)
            or (
                previous_resolved is not None
                and current_resolved is not None
                and _resolved_global_payload(previous_resolved)
                == _resolved_global_payload(current_resolved)
            )
        ),
    )


@dataclass(frozen=True)
class ScopeExclusionSummary:
    """How many findings on each side of a Re-QC delta were excluded because
    their logical scope's comparison policy differs between the two runs'
    profile snapshots -- never counted resolved/new, never evaluated at
    all for this comparison (plan-20260913 Step 10's "comparable/not-
    evaluated summaries" criterion). Zero on both sides is the common,
    unchanged-profile case.
    """

    previous_excluded: int = 0
    current_excluded: int = 0

    @property
    def any_excluded(self) -> bool:
        return self.previous_excluded > 0 or self.current_excluded > 0


def compatible_compare_findings(
    previous_findings: Sequence[Finding],
    current_findings: Sequence[Finding],
    *,
    previous_profile: DeliverableProfile,
    current_profile: DeliverableProfile,
    previous_resolved: ResolvedInputConfigurationV1 | None = None,
    current_resolved: ResolvedInputConfigurationV1 | None = None,
) -> tuple[FindingsDelta, ScopeExclusionSummary]:
    """Drop-in, scope-aware replacement for
    ``qc_tool.engine.compare_findings``.

    Excludes any finding (from either side) whose logical scope's
    configuration differs between the two profile snapshots before
    delegating to ``compare_findings`` -- so a suppressed or rescoped
    finding is never silently counted "resolved", and a newly-visible one
    is never silently counted "new". Returns ``(delta, exclusion_summary)``;
    callers disclose ``exclusion_summary.any_excluded`` the same way they
    already disclose an output-representation change.

    Same caller contract as ``compare_findings``: both finding sequences
    must already have passed ``output_representations_compatible`` --
    this is an independent, additional gate, not a replacement for it.
    """
    compatibility = configuration_compatible(
        previous_profile,
        current_profile,
        previous_resolved=previous_resolved,
        current_resolved=current_resolved,
    )
    filtered_previous = [f for f in previous_findings if compatibility.comparable(f)]
    filtered_current = [f for f in current_findings if compatibility.comparable(f)]
    exclusion_summary = ScopeExclusionSummary(
        previous_excluded=len(previous_findings) - len(filtered_previous),
        current_excluded=len(current_findings) - len(filtered_current),
    )
    return compare_findings(filtered_previous, filtered_current), exclusion_summary
