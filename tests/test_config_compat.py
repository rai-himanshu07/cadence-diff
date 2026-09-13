"""Tests for the Step 2 compatibility/precedence/migration layer
(``qc_tool.config.compat``).

Plan: docs/plans/plan-20260913-mode-aware-configuration-wizard.md, Step 2.
"""

from __future__ import annotations

import pytest

from qc_tool.config.compat import (
    browser_save_requires_exclusion_upgrade,
    describe_legacy_excel_scope,
    region_authority_conflicts,
    resolve_legacy_configuration,
)
from qc_tool.config.input_contract import (
    LogicalMemberContract,
    LogicalRegionContract,
    LogicalSheetContract,
    WorkbookInputContract,
)
from qc_tool.config.lint import lint_profile
from qc_tool.config.profile import (
    CadenceBand,
    ComparisonPrerequisite,
    DeliverableProfile,
    ExcelMemberProfile,
    ExcelProfile,
    PptProfile,
    RegionOverride,
    RowIdentityRule,
    SheetProfile,
    profile_sha256,
)
from qc_tool.coverage import QCRunMode

# --- resolve_legacy_configuration ------------------------------------------


def test_resolve_legacy_configuration_returns_the_legacy_placeholder() -> None:
    profile = DeliverableProfile(name="monthly")
    resolved = resolve_legacy_configuration(profile, QCRunMode.CYCLE_COMPARISON)
    assert resolved.source == "legacy_default"
    assert resolved.profile_name == "monthly"
    assert resolved.profile_sha256 == profile_sha256(profile)
    assert resolved.mode is QCRunMode.CYCLE_COMPARISON


def test_resolve_legacy_configuration_rejects_a_profile_with_a_contract() -> None:
    profile = DeliverableProfile(
        name="monthly",
        input_contract=WorkbookInputContract(
            members=(LogicalMemberContract(member_id="primary"),)
        ),
    )
    with pytest.raises(ValueError, match="only applies to a profile with no"):
        resolve_legacy_configuration(profile, QCRunMode.CYCLE_COMPARISON)


# --- region_authority_conflicts --------------------------------------------


def test_no_contract_means_no_conflicts() -> None:
    profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(
            sheets={
                "Data": SheetProfile(
                    regions=[RegionOverride(range="A1:D10", orientation="wide")]
                )
            }
        ),
    )
    assert region_authority_conflicts(profile) == ()


def test_a_bare_identity_hint_with_no_structural_decision_never_conflicts() -> None:
    """A logical sheet that only records a preferred name (no regions, the
    default 'automatic' mode) coexists peacefully with legacy fields on the
    same physical sheet -- it has not actually claimed structural authority.
    """
    profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(
            sheets={
                "Data": SheetProfile(
                    regions=[RegionOverride(range="A1:D10", orientation="wide")]
                )
            }
        ),
        input_contract=WorkbookInputContract(
            members=(
                LogicalMemberContract(
                    member_id="primary",
                    sheets=(
                        LogicalSheetContract(
                            sheet_id="data", preferred_sheet_name="Data"
                        ),
                    ),
                ),
            )
        ),
    )
    assert region_authority_conflicts(profile) == ()


def test_a_migrated_region_conflicts_with_a_legacy_region_override_on_the_same_sheet() -> None:
    profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(
            sheets={
                "Data": SheetProfile(
                    regions=[RegionOverride(range="A1:D10", orientation="wide")]
                )
            }
        ),
        input_contract=WorkbookInputContract(
            members=(
                LogicalMemberContract(
                    member_id="primary",
                    sheets=(
                        LogicalSheetContract(
                            sheet_id="data",
                            preferred_sheet_name="Data",
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
    conflicts = region_authority_conflicts(profile)
    assert len(conflicts) == 1
    assert "Data" in conflicts[0]
    assert "RegionOverride" in conflicts[0]


def test_a_migrated_region_conflicts_with_legacy_row_identity_rules() -> None:
    profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(
            sheets={
                "Data": SheetProfile(
                    row_identity_rules=[
                        RowIdentityRule(anchor_cell="A1", identity_columns=["B"])
                    ]
                )
            }
        ),
        input_contract=WorkbookInputContract(
            members=(
                LogicalMemberContract(
                    member_id="primary",
                    sheets=(
                        LogicalSheetContract(
                            sheet_id="data",
                            preferred_sheet_name="Data",
                            default_mode="positional",
                        ),
                    ),
                ),
            )
        ),
    )
    conflicts = region_authority_conflicts(profile)
    assert len(conflicts) == 1
    assert "row_identity_rules" in conflicts[0]


def test_conflicts_are_detected_for_package_members_too() -> None:
    profile = DeliverableProfile(
        name="pkg",
        excel=ExcelProfile(
            members={
                "ops": ExcelMemberProfile(
                    sheets={
                        "Ops": SheetProfile(
                            regions=[RegionOverride(range="A1:D10", orientation="wide")]
                        )
                    }
                )
            }
        ),
        input_contract=WorkbookInputContract(
            members=(
                LogicalMemberContract(
                    member_id="ops",
                    sheets=(
                        LogicalSheetContract(
                            sheet_id="ops_data",
                            preferred_sheet_name="Ops",
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
    conflicts = region_authority_conflicts(profile)
    assert len(conflicts) == 1
    assert "'ops'" in conflicts[0]


def test_lint_profile_reports_an_error_for_mixed_structural_authority() -> None:
    profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(
            sheets={
                "Data": SheetProfile(
                    regions=[RegionOverride(range="A1:D10", orientation="wide")]
                )
            }
        ),
        input_contract=WorkbookInputContract(
            members=(
                LogicalMemberContract(
                    member_id="primary",
                    sheets=(
                        LogicalSheetContract(
                            sheet_id="data",
                            preferred_sheet_name="Data",
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
    issues = lint_profile(profile)
    errors = [issue for issue in issues if issue.level == "error"]
    assert any("RegionOverride" in issue.message for issue in errors)


def test_lint_profile_stays_clean_for_an_unmigrated_legacy_profile() -> None:
    profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(
            sheets={
                "Data": SheetProfile(
                    regions=[RegionOverride(range="A1:D10", orientation="wide")]
                )
            }
        ),
    )
    assert lint_profile(profile) == []


# --- browser_save_requires_exclusion_upgrade -------------------------------


def test_no_legacy_exclusions_requires_no_upgrade() -> None:
    profile = DeliverableProfile(name="monthly")
    assert browser_save_requires_exclusion_upgrade(profile) == ()


def test_bare_ignore_sheets_requires_upgrade() -> None:
    profile = DeliverableProfile(
        name="monthly", excel=ExcelProfile(ignore_sheets=["Notes"])
    )
    violations = browser_save_requires_exclusion_upgrade(profile)
    assert len(violations) == 1
    assert "ignore_sheets" in violations[0]


def test_bare_sheet_ignore_flag_requires_upgrade() -> None:
    profile = DeliverableProfile(
        name="monthly", excel=ExcelProfile(sheets={"Notes": SheetProfile(ignore=True)})
    )
    violations = browser_save_requires_exclusion_upgrade(profile)
    assert any("ignore" in violation for violation in violations)


def test_bare_ignore_ranges_requires_upgrade() -> None:
    profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(sheets={"Data": SheetProfile(ignore_ranges=["A1:A5"])}),
    )
    violations = browser_save_requires_exclusion_upgrade(profile)
    assert any("ignore_ranges" in violation for violation in violations)


def test_member_scoped_legacy_exclusions_also_require_upgrade() -> None:
    profile = DeliverableProfile(
        name="pkg",
        excel=ExcelProfile(
            members={"ops": ExcelMemberProfile(ignore_sheets=["Notes"])}
        ),
    )
    violations = browser_save_requires_exclusion_upgrade(profile)
    assert any("members" in violation and "ops" in violation for violation in violations)


def test_this_predicate_is_independent_of_input_contract_presence() -> None:
    """The wizard demands the upgrade the moment a bare legacy exclusion is
    present, whether or not a contract has already begun migrating.
    """
    without_contract = DeliverableProfile(
        name="monthly", excel=ExcelProfile(ignore_sheets=["Notes"])
    )
    with_contract = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(ignore_sheets=["Notes"]),
        input_contract=WorkbookInputContract(
            members=(LogicalMemberContract(member_id="primary"),)
        ),
    )
    assert browser_save_requires_exclusion_upgrade(without_contract)
    assert browser_save_requires_exclusion_upgrade(with_contract)


# --- describe_legacy_excel_scope -------------------------------------------


def test_describe_legacy_excel_scope_is_empty_for_a_bare_profile() -> None:
    assert describe_legacy_excel_scope(DeliverableProfile(name="monthly")) == ()


def test_describe_legacy_excel_scope_covers_every_legacy_field_kind() -> None:
    profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(
            ignore_sheets=["Archive"],
            sheets={
                "Data": SheetProfile(
                    regions=[RegionOverride(range="A1:D10", orientation="wide")],
                    row_identity_rules=[
                        RowIdentityRule(anchor_cell="A1", identity_columns=["B"])
                    ],
                    cadence_bands=[CadenceBand(range="E1:P1", kind="month")],
                )
            },
            comparison_prerequisites=[
                ComparisonPrerequisite(name="Scenario", sheet="Config", cell="B2")
            ],
        ),
        ppt=PptProfile(slide_pins={"Intro": "Overview"}, required_slides=["Summary"]),
    )
    lines = describe_legacy_excel_scope(profile)
    joined = "\n".join(lines)
    assert "Archive" in joined
    assert "RegionOverride" in joined
    assert "RowIdentityRule" in joined
    assert "CadenceBand" in joined
    assert "Scenario" in joined
    assert "Intro" in joined and "Overview" in joined
    assert "Summary" in joined


def test_describe_legacy_excel_scope_never_widens_or_mutates_the_profile() -> None:
    """A pure display projection: calling it must not change canonical
    profile bytes.
    """
    profile = DeliverableProfile(
        name="monthly", excel=ExcelProfile(ignore_sheets=["Archive"])
    )
    before = profile_sha256(profile)
    describe_legacy_excel_scope(profile)
    assert profile_sha256(profile) == before
