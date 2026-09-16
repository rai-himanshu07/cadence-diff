"""Tests for the Step 4 scope-semantics compatibility service
(``qc_tool.history.config_compatibility``).

Plan: docs/plans/plan-20260913-mode-aware-configuration-wizard.md, Step 4.
"""

from __future__ import annotations

import datetime as dt
from typing import Literal

from qc_tool.config.input_contract import (
    LogicalMemberContract,
    LogicalRegionContract,
    LogicalSheetContract,
    WorkbookInputContract,
)
from qc_tool.config.profile import (
    DeliverableProfile,
    ExcelAvailabilityRule,
    ExcelProfile,
    FindingWaiver,
    NumericTolerance,
    SheetProfile,
)
from qc_tool.config.resolved_input import (
    ResolvedColumn,
    ResolvedInputConfigurationV1,
    ResolvedMember,
    ResolvedRegion,
    ResolvedSelector,
    ResolvedSheet,
)
from qc_tool.findings import Finding, FindingClass, LogicalFindingAddress, Severity
from qc_tool.history.config_compatibility import (
    compatible_compare_findings,
    configuration_compatible,
)


def _excel_finding(
    *,
    sheet: str = "Data",
    finding_class: FindingClass = FindingClass.VALUE_CHANGED,
    logical_address: LogicalFindingAddress | None = None,
    artifact_member: str = "primary",
) -> Finding:
    return Finding(
        artifact="excel",
        artifact_member=artifact_member,
        finding_class=finding_class,
        severity=Severity.CRITICAL,
        sheet=sheet,
        location="A1",
        message="changed",
        logical_address=logical_address,
    )


def _ppt_finding() -> Finding:
    return Finding(
        artifact="ppt",
        finding_class=FindingClass.VALUE_CHANGED,
        severity=Severity.CRITICAL,
        slide="Overview",
        message="changed",
    )


def _crosscheck_finding() -> Finding:
    return Finding(
        artifact="crosscheck",
        finding_class=FindingClass.CROSSCHECK_MISMATCH,
        severity=Severity.CRITICAL,
        message="changed",
    )


def _resolved(
    *,
    mode: Literal["automatic", "keyed", "positional", "excluded"] = "keyed",
    baseline_sheet: str = "Data 2025",
    current_sheet: str = "Data 2026",
    baseline_letter: str = "A",
    current_letter: str = "A",
    selector_formula_backed: bool = False,
    baseline_hash: str = "a" * 64,
    current_hash: str = "b" * 64,
) -> ResolvedInputConfigurationV1:
    return ResolvedInputConfigurationV1(
        members=(
            ResolvedMember(
                member_id="primary",
                baseline_source_sha256=baseline_hash,
                current_source_sha256=current_hash,
                sheets=(
                    ResolvedSheet(
                        sheet_id="ledger",
                        baseline_sheet_name=baseline_sheet,
                        current_sheet_name=current_sheet,
                        regions=(
                            ResolvedRegion(
                                region_id="r1",
                                mode=mode,
                                baseline_outer_range="A1:B10",
                                current_outer_range="A1:B10",
                                columns=(
                                    ResolvedColumn(
                                        column_id="account",
                                        baseline_letter=baseline_letter,
                                        current_letter=current_letter,
                                        alignment_role="identity",
                                    ),
                                ),
                            ),
                        ),
                        selectors=(
                            ResolvedSelector(
                                selector_id="scenario",
                                baseline_cell="B2",
                                current_cell="B2",
                                equal=True,
                                baseline_formula_backed=selector_formula_backed,
                                current_formula_backed=selector_formula_backed,
                            ),
                        ),
                    ),
                ),
            ),
        )
    )


def test_byte_identical_profiles_are_fully_compatible() -> None:
    profile = DeliverableProfile(name="monthly")
    compat = configuration_compatible(profile, profile.model_copy(deep=True))
    assert compat.fully_compatible()
    assert compat.comparable(_excel_finding())
    assert compat.comparable(_ppt_finding())
    assert compat.comparable(_crosscheck_finding())


def test_alias_only_change_remains_comparable() -> None:
    """``name``/``description`` are presentation_only; must never gate."""
    previous = DeliverableProfile(name="monthly")
    current = DeliverableProfile(name="monthly-renamed", description="new label")
    compat = configuration_compatible(previous, current)
    assert compat.fully_compatible()
    assert compat.comparable(_excel_finding())


def test_global_tolerance_relaxation_makes_every_scope_not_comparable() -> None:
    previous = DeliverableProfile(name="monthly")
    current = DeliverableProfile(
        name="monthly", tolerance=NumericTolerance(absolute=5.0)
    )
    compat = configuration_compatible(previous, current)
    assert not compat.fully_compatible()
    assert not compat.comparable(_excel_finding())
    assert not compat.comparable(_ppt_finding())
    assert not compat.comparable(_crosscheck_finding())


def test_waiver_reason_only_change_remains_comparable() -> None:
    """``waivers[].reason`` is presentation_only; expiry/scope fields are not."""
    previous = DeliverableProfile(
        name="monthly",
        waivers=[
            FindingWaiver(
                finding_class=FindingClass.VALUE_CHANGED,
                reason="old reason",
                expires=dt.date(2030, 1, 1),
                sheet="Data",
            )
        ],
    )
    current = DeliverableProfile(
        name="monthly",
        waivers=[
            FindingWaiver(
                finding_class=FindingClass.VALUE_CHANGED,
                reason="new reason, more detail",
                expires=dt.date(2030, 1, 1),
                sheet="Data",
            )
        ],
    )
    compat = configuration_compatible(previous, current)
    assert compat.fully_compatible()


def test_waiver_expiry_change_makes_every_scope_not_comparable() -> None:
    previous = DeliverableProfile(
        name="monthly",
        waivers=[
            FindingWaiver(
                finding_class=FindingClass.VALUE_CHANGED,
                reason="r",
                expires=dt.date(2030, 1, 1),
                sheet="Data",
            )
        ],
    )
    current = DeliverableProfile(
        name="monthly",
        waivers=[
            FindingWaiver(
                finding_class=FindingClass.VALUE_CHANGED,
                reason="r",
                expires=dt.date(2030, 6, 1),
                sheet="Data",
            )
        ],
    )
    compat = configuration_compatible(previous, current)
    assert not compat.fully_compatible()


def test_per_sheet_availability_rule_change_only_affects_that_sheet() -> None:
    previous = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(
            sheets={
                "Data": SheetProfile(availability_rules=[_availability_rule("A1")]),
                "Other": SheetProfile(),
            }
        ),
    )
    current = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(
            sheets={
                "Data": SheetProfile(availability_rules=[_availability_rule("B1")]),
                "Other": SheetProfile(),
            }
        ),
    )
    compat = configuration_compatible(previous, current)
    assert compat.fully_compatible()  # global gate is unaffected
    assert not compat.comparable(_excel_finding(sheet="Data"))
    assert compat.comparable(_excel_finding(sheet="Other"))
    assert compat.comparable(_ppt_finding())


def _availability_rule(required_through: str) -> ExcelAvailabilityRule:
    return ExcelAvailabilityRule(
        range="A1:A10", periods="A1:A10", required_through=required_through
    )


def test_sheet_rename_alone_stays_comparable_via_logical_address() -> None:
    """Step 3's whole point: a confirmed rename must never itself break
    comparability. The legacy per-physical-name entry differs (different
    dict key) but the finding carries a stable ``sheet_id``, so only the
    saved contract entry for that id is consulted.
    """
    contract = WorkbookInputContract(
        members=(
            LogicalMemberContract(
                member_id="primary",
                sheets=(
                    LogicalSheetContract(
                        sheet_id="ledger",
                        regions=(
                            LogicalRegionContract(
                                region_id="r1", anchor_cell="A1", mode="keyed"
                            ),
                        ),
                    ),
                ),
            ),
        )
    )
    previous = DeliverableProfile(
        name="monthly",
        input_contract=contract,
        excel=ExcelProfile(sheets={"Ledger2025": SheetProfile(ignore=True)}),
    )
    current = DeliverableProfile(
        name="monthly",
        input_contract=contract,
        excel=ExcelProfile(sheets={"Ledger2026": SheetProfile(ignore=True)}),
    )
    address = LogicalFindingAddress(member_id="primary", sheet_id="ledger")
    compat = configuration_compatible(previous, current)
    assert compat.comparable(_excel_finding(sheet="Ledger2026", logical_address=address))


def test_logical_sheet_region_change_makes_that_scope_not_comparable() -> None:
    previous = DeliverableProfile(
        name="monthly",
        input_contract=WorkbookInputContract(
            members=(
                LogicalMemberContract(
                    member_id="primary",
                    sheets=(
                        LogicalSheetContract(
                            sheet_id="ledger",
                            regions=(
                                LogicalRegionContract(
                                    region_id="r1", anchor_cell="A1", mode="keyed"
                                ),
                            ),
                        ),
                    ),
                ),
            )
        ),
    )
    current = DeliverableProfile(
        name="monthly",
        input_contract=WorkbookInputContract(
            members=(
                LogicalMemberContract(
                    member_id="primary",
                    sheets=(
                        LogicalSheetContract(
                            sheet_id="ledger",
                            regions=(
                                LogicalRegionContract(
                                    region_id="r1",
                                    anchor_cell="A1",
                                    mode="positional",
                                ),
                            ),
                        ),
                    ),
                ),
            )
        ),
    )
    address = LogicalFindingAddress(member_id="primary", sheet_id="ledger")
    compat = configuration_compatible(previous, current)
    assert not compat.comparable(_excel_finding(sheet="Ledger", logical_address=address))


def test_logical_sheet_alias_only_change_remains_comparable() -> None:
    def _contract(alias: str) -> WorkbookInputContract:
        return WorkbookInputContract(
            members=(
                LogicalMemberContract(
                    member_id="primary",
                    sheets=(LogicalSheetContract(sheet_id="ledger", alias=alias),),
                ),
            )
        )

    previous = DeliverableProfile(name="monthly", input_contract=_contract("Ledger"))
    current = DeliverableProfile(name="monthly", input_contract=_contract("Ledger (v2)"))
    address = LogicalFindingAddress(member_id="primary", sheet_id="ledger")
    compat = configuration_compatible(previous, current)
    assert compat.comparable(_excel_finding(sheet="Ledger", logical_address=address))


def test_compatible_compare_findings_excludes_only_the_changed_scope() -> None:
    previous_profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(
            sheets={
                "Data": SheetProfile(availability_rules=[_availability_rule("A1")]),
                "Other": SheetProfile(),
            }
        ),
    )
    current_profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(
            sheets={
                "Data": SheetProfile(availability_rules=[_availability_rule("B1")]),
                "Other": SheetProfile(),
            }
        ),
    )
    # "Data" gained a finding purely because of the scope change (not a real
    # workbook difference) -- must not count as "new". "Other" is untouched
    # and its resolved finding must still count.
    previous_findings = [
        _excel_finding(sheet="Data"),
        _excel_finding(sheet="Other"),
    ]
    current_findings = [_excel_finding(sheet="Data")]
    delta, exclusion_summary = compatible_compare_findings(
        previous_findings,
        current_findings,
        previous_profile=previous_profile,
        current_profile=current_profile,
    )
    assert exclusion_summary.any_excluded
    assert exclusion_summary.previous_excluded == 1  # "Data" (previous side)
    assert exclusion_summary.current_excluded == 1  # "Data" (current side)
    assert delta.resolved == 1  # "Other" resolved
    assert delta.new == 0
    assert delta.persisting == 0


def test_compatible_compare_findings_is_a_true_delta_when_nothing_scope_related_changed() -> None:
    profile = DeliverableProfile(name="monthly")
    previous_findings = [_excel_finding(sheet="Data")]
    current_findings = [_excel_finding(sheet="Data"), _excel_finding(sheet="Other")]
    delta, exclusion_summary = compatible_compare_findings(
        previous_findings,
        current_findings,
        previous_profile=profile,
        current_profile=profile.model_copy(deep=True),
    )
    assert not exclusion_summary.any_excluded
    assert exclusion_summary.previous_excluded == 0
    assert exclusion_summary.current_excluded == 0
    assert delta.persisting == 1
    assert delta.new == 1
    assert delta.resolved == 0


def test_run_only_region_mode_change_makes_logical_scope_not_comparable() -> None:
    profile = DeliverableProfile(name="monthly")
    finding = _excel_finding(
        logical_address=LogicalFindingAddress(
            member_id="primary", sheet_id="ledger", region_id="r1"
        )
    )
    compatibility = configuration_compatible(
        profile,
        profile,
        previous_resolved=_resolved(mode="keyed"),
        current_resolved=_resolved(mode="positional"),
    )

    assert not compatibility.comparable(finding)


def test_source_hash_and_physical_movement_do_not_break_logical_comparability() -> None:
    profile = DeliverableProfile(name="monthly")
    finding = _excel_finding(
        logical_address=LogicalFindingAddress(
            member_id="primary",
            sheet_id="ledger",
            region_id="r1",
            column_id="account",
        )
    )
    compatibility = configuration_compatible(
        profile,
        profile,
        previous_resolved=_resolved(),
        current_resolved=_resolved(
            baseline_sheet="Renamed baseline",
            current_sheet="Renamed current",
            baseline_letter="C",
            current_letter="D",
            baseline_hash="c" * 64,
            current_hash="d" * 64,
        ),
    )

    assert compatibility.comparable(finding)


def test_selector_fact_change_makes_its_sheet_not_comparable() -> None:
    profile = DeliverableProfile(name="monthly")
    finding = _excel_finding(
        logical_address=LogicalFindingAddress(
            member_id="primary", sheet_id="ledger"
        )
    )
    compatibility = configuration_compatible(
        profile,
        profile,
        previous_resolved=_resolved(selector_formula_backed=False),
        current_resolved=_resolved(selector_formula_backed=True),
    )

    assert not compatibility.comparable(finding)


def test_mixed_legacy_and_resolved_runs_are_not_comparable() -> None:
    profile = DeliverableProfile(name="monthly")
    compatibility = configuration_compatible(
        profile,
        profile,
        previous_resolved=None,
        current_resolved=_resolved(),
    )

    assert not compatibility.comparable(_excel_finding())
