"""Compatibility, precedence, and profile-migration layer between the
legacy physical-name profile fields and the new saved ``InputContractV1``.

Plan: docs/plans/plan-20260913-mode-aware-configuration-wizard.md, Step 2.

Three responsibilities, kept deliberately separate:

1. ``resolve_legacy_configuration`` -- the Step 4 dispatch point for a
   profile with no ``input_contract``: legacy physical-name fields and
   automatic detection drive execution unchanged, byte-for-byte.
2. ``region_authority_conflicts`` -- a profile-level self-consistency
   check wired into ``qc_tool.config.lint.lint_profile`` (universal, not
   browser-specific): a migrated logical region/sheet and a legacy
   ``RegionOverride``/``RowIdentityRule`` must never both claim the same
   physical sheet.
3. ``browser_save_requires_exclusion_upgrade`` -- a pure predicate the
   Step 7 browser wizard's save handler must call before persisting a
   profile edited through the wizard. It is NOT wired into
   ``qc_tool.config.lint.lint_profile``/``qc_tool.config.editor.save_draft``
   because CLI and the legacy generic form editor must keep honoring bare
   legacy exclusions directly, unchanged.

``describe_legacy_excel_scope`` renders legacy structural fields in the new
wizard's vocabulary for display only; it never changes what a field does
(no semantic widening) and never persists anything.
"""

from __future__ import annotations

from qc_tool.config.profile import DeliverableProfile, ExcelMemberProfile
from qc_tool.config.profile import profile_sha256 as _profile_sha256
from qc_tool.config.resolved_input import ResolvedInputConfigurationV1
from qc_tool.coverage import QCRunMode


def resolve_legacy_configuration(
    profile: DeliverableProfile, mode: QCRunMode
) -> ResolvedInputConfigurationV1:
    """The dispatch point for a profile with no saved ``input_contract``.

    Always returns the legacy placeholder: legacy physical-name fields and
    automatic detection continue to drive execution unchanged. A profile
    that DOES carry an ``input_contract`` is resolved by the real per-run
    resolver (setup analysis + engine wiring, later steps), not by this
    function.
    """
    if profile.input_contract is not None:
        raise ValueError(
            "resolve_legacy_configuration only applies to a profile with no "
            "input_contract; use the real per-run resolver instead"
        )
    return ResolvedInputConfigurationV1.legacy_default(
        profile_name=profile.name,
        profile_sha256=_profile_sha256(profile),
        mode=mode,
    )


def _member_excel_profile(
    profile: DeliverableProfile, member_id: str
) -> ExcelMemberProfile | None:
    if member_id == "primary":
        return profile.excel
    return profile.excel.members.get(member_id)


def region_authority_conflicts(profile: DeliverableProfile) -> tuple[str, ...]:
    """Detect a physical sheet claimed by BOTH a migrated logical region and
    a legacy structural field. Overlapping legacy structural fields are a
    lint/validation error rather than silently winning or merging.

    A logical sheet only counts as "migrated" (structurally authoritative)
    when it actually records a structural decision -- at least one region
    with ``mode != "automatic"``, or ``default_mode == "positional"``. A
    bare identity/preference hint (no regions, ``default_mode ==
    "automatic"``) never conflicts with legacy fields on the same sheet.
    """
    contract = profile.input_contract
    if contract is None:
        return ()
    conflicts: list[str] = []
    for logical_member in contract.members:
        member_profile = _member_excel_profile(profile, logical_member.member_id)
        if member_profile is None:
            continue
        for logical_sheet in logical_member.sheets:
            physical_name = logical_sheet.preferred_sheet_name
            if not physical_name:
                continue
            sheet_profile = member_profile.sheets.get(physical_name)
            if sheet_profile is None:
                continue
            is_migrated = logical_sheet.default_mode == "positional" or any(
                region.mode != "automatic" for region in logical_sheet.regions
            )
            if not is_migrated:
                continue
            where = f"member {logical_member.member_id!r} sheet {physical_name!r}"
            if sheet_profile.regions:
                conflicts.append(
                    f"{where}: legacy excel.sheets[].regions conflicts with a "
                    "migrated logical region; remove the legacy RegionOverride "
                    "entries"
                )
            if sheet_profile.row_identity_rules:
                conflicts.append(
                    f"{where}: legacy row_identity_rules conflict with a "
                    "migrated logical region; remove them or revert the "
                    "logical region to mode='automatic'"
                )
    return tuple(conflicts)


def browser_save_requires_exclusion_upgrade(
    profile: DeliverableProfile,
) -> tuple[str, ...]:
    """Bare legacy exclusions the BROWSER wizard must refuse to save until
    upgraded to an explicit reason/expiry.

    Scoped to the browser wizard's own save handler (Step 7); CLI saves and
    the legacy generic form editor's ``save_draft`` are unaffected and keep
    honoring bare legacy exclusions directly.
    """
    violations: list[str] = []

    def _scan(prefix: str, excel: ExcelMemberProfile) -> None:
        for name in excel.ignore_sheets:
            violations.append(f"{prefix}.ignore_sheets[{name!r}] has no reason/expiry")
        for sheet_name, sheet_profile in excel.sheets.items():
            sheet_where = f"{prefix}.sheets[{sheet_name!r}]"
            if sheet_profile.ignore:
                violations.append(f"{sheet_where}.ignore has no reason/expiry")
            for cell_range in sheet_profile.ignore_ranges:
                violations.append(
                    f"{sheet_where}.ignore_ranges[{cell_range!r}] has no "
                    "reason/expiry"
                )

    _scan("excel", profile.excel)
    for member_id, member_profile in profile.excel.members.items():
        _scan(f"excel.members[{member_id!r}]", member_profile)
    return tuple(violations)


def describe_legacy_excel_scope(profile: DeliverableProfile) -> tuple[str, ...]:
    """Human-readable projection of legacy Excel/PPT structural fields into
    the new wizard's vocabulary, for display only. Never changes behavior,
    never persists, never widens what a legacy field actually does.
    """
    lines: list[str] = []
    if profile.excel.ignore_sheets:
        joined = ", ".join(sorted(profile.excel.ignore_sheets))
        lines.append(f"excluded sheets (legacy): {joined}")
    for sheet_name, sheet_profile in sorted(profile.excel.sheets.items()):
        if sheet_profile.ignore:
            lines.append(f"sheet {sheet_name!r}: excluded (legacy ignore)")
        for region in sheet_profile.regions:
            lines.append(
                f"sheet {sheet_name!r}: positional region {region.cell_range} "
                f"({region.orientation}, legacy RegionOverride)"
            )
        for rule in sheet_profile.row_identity_rules:
            identity = ", ".join(rule.identity_columns)
            lines.append(
                f"sheet {sheet_name!r}: keyed region anchored at "
                f"{rule.anchor_cell} (legacy RowIdentityRule, identity "
                f"columns {identity})"
            )
        for band in sheet_profile.cadence_bands:
            lines.append(
                f"sheet {sheet_name!r}: {band.kind} period band "
                f"{band.cell_range} (legacy CadenceBand)"
            )
    for index, prerequisite in enumerate(profile.excel.comparison_prerequisites):
        label = prerequisite.name or str(index)
        lines.append(
            f"selector prerequisite (legacy): {label} at "
            f"{prerequisite.sheet}!{prerequisite.cell}"
        )
    for name, target in sorted(profile.ppt.slide_pins.items()):
        lines.append(f"slide pin (legacy): {name!r} -> {target!r}")
    for slide in profile.ppt.required_slides:
        lines.append(f"required slide (legacy): {slide!r}")
    return tuple(lines)
