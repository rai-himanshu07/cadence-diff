"""Versioned package identity and member-scoped profile contracts."""

from __future__ import annotations

import importlib
from pathlib import Path

import pytest
from pydantic import ValidationError

from qc_tool.config.profile import (
    ComparisonPrerequisite,
    DeliverableProfile,
    ExcelMemberProfile,
    ExcelProfile,
    RowIdentityRule,
    SheetProfile,
    excel_profile_for_member,
    profile_for_excel_member,
    profile_sha256,
)
from qc_tool.package import (
    MAX_WORKBOOKS_PER_SIDE,
    PackageArtifact,
    PackageManifest,
    PackageMember,
    PackageSide,
    paths_by_member,
    structural_member_aliases,
)


def _member(
    member_id: str,
    *,
    side: PackageSide = PackageSide.CURRENT,
    artifact: PackageArtifact = PackageArtifact.EXCEL,
) -> PackageMember:
    suffix = ".xlsx" if artifact is PackageArtifact.EXCEL else ".pptx"
    return PackageMember(
        member_id=member_id,
        side=side,
        artifact=artifact,
        display_name=f"{member_id}{suffix}",
    )


def test_manifest_caps_excel_members_at_eight_per_side() -> None:
    members = tuple(
        _member(f"wb{index}")
        for index in range(MAX_WORKBOOKS_PER_SIDE + 1)
    )

    with pytest.raises(ValidationError, match="no more than 8 Excel members"):
        PackageManifest(members=members)


def test_manifest_rejects_duplicate_identity_and_multiple_ppt_per_side() -> None:
    duplicate = _member("ops")
    with pytest.raises(ValidationError, match="identities must be unique"):
        PackageManifest(members=(duplicate, duplicate))

    with pytest.raises(ValidationError, match="one PowerPoint"):
        PackageManifest(
            members=(
                _member("primary", artifact=PackageArtifact.PPT),
                _member("appendix", artifact=PackageArtifact.PPT),
            )
        )


def test_manifest_canonicalizes_side_artifact_and_member_order() -> None:
    manifest = PackageManifest(
        members=(
            _member("zeta"),
            _member("primary", artifact=PackageArtifact.PPT),
            _member("alpha", side=PackageSide.BASELINE),
            _member("alpha"),
            _member("primary", side=PackageSide.BASELINE, artifact=PackageArtifact.PPT),
        )
    )

    assert manifest.role_keys == (
        "baseline_excel:alpha",
        "baseline_ppt",
        "current_excel:alpha",
        "current_excel:zeta",
        "current_ppt",
    )


def test_primary_and_non_primary_role_keys_are_lossless() -> None:
    primary = _member("primary")
    other = _member("ops")

    assert primary.role_key == "current_excel"
    assert other.role_key == "current_excel:ops"
    manifest = PackageManifest.from_role_files(
        {
            primary.role_key: Path("core.xlsx"),
            other.role_key: Path("ops.xlsx"),
        }
    )
    assert {member.role_key: member.member_id for member in manifest.members} == {
        "current_excel": "primary",
        "current_excel:ops": "ops",
    }


def test_legacy_projection_round_trips_and_nonlegacy_refuses_projection() -> None:
    legacy = PackageManifest.from_legacy_files(
        {
            "current_excel": Path("core.xlsx"),
            "current_ppt": Path("deck.pptx"),
        }
    )

    assert legacy.is_legacy_projection is True
    assert legacy.to_legacy_roles() == {
        "current_excel": "core.xlsx",
        "current_ppt": "deck.pptx",
    }
    with pytest.raises(ValueError, match="not a primary-role projection"):
        PackageManifest(members=(_member("ops"),)).to_legacy_roles()
    with pytest.raises(ValueError, match="unknown legacy package roles"):
        PackageManifest.from_legacy_files({"current_excel:ops": Path("ops.xlsx")})


def test_paths_by_member_requires_exact_role_parity() -> None:
    manifest = PackageManifest(members=(_member("ops"),))

    with pytest.raises(ValueError, match=r"missing=.*current_excel:ops"):
        paths_by_member({"current_excel": Path("core.xlsx")}, manifest)
    assert paths_by_member(
        {"current_excel:ops": Path("ops.xlsx")},
        manifest,
    ) == {"current_excel:ops": Path("ops.xlsx")}


def test_member_ids_are_bounded_lowercase_identifiers() -> None:
    for invalid in ("Ops", "2ops", "has space", "a" * 33):
        with pytest.raises(ValidationError):
            _member(invalid)


def test_structural_aliases_preserve_cross_side_pairing_without_ids() -> None:
    manifest = PackageManifest(
        members=(
            _member("client_ops", side=PackageSide.BASELINE),
            _member("client_ops"),
            _member("client_core"),
            _member("primary", artifact=PackageArtifact.PPT),
        )
    )

    aliases = structural_member_aliases(manifest)

    assert aliases[(PackageArtifact.EXCEL, "client_core")] == "excel_001"
    assert aliases[(PackageArtifact.EXCEL, "client_ops")] == "excel_002"
    assert aliases[(PackageArtifact.PPT, "primary")] == "primary"


def test_single_primary_workbook_uses_legacy_unscoped_profile() -> None:
    profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(sheets={"Data": SheetProfile(ignore=True)}),
    )

    projected = excel_profile_for_member(
        profile,
        "primary",
        workbook_count=1,
    )

    assert projected.sheets["Data"].ignore is True


def test_unscoped_excel_rules_are_refused_for_multi_workbook_package() -> None:
    profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(sheets={"Data": SheetProfile(ignore=True)}),
    )

    with pytest.raises(ValueError, match="ambiguous for this package"):
        excel_profile_for_member(profile, "ops", workbook_count=2)


def test_member_scoped_profile_projects_only_the_selected_member() -> None:
    profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(
            members={
                "core": ExcelMemberProfile(
                    sheets={"Core": SheetProfile(ignore=True)}
                ),
                "ops": ExcelMemberProfile(
                    sheets={"Ops": SheetProfile(refresh_ranges=["A1:B2"])}
                ),
            }
        ),
    )

    projected = profile_for_excel_member(profile, "ops", 2)

    assert set(projected.excel.sheets) == {"Ops"}
    assert projected.excel.sheets["Ops"].refresh_ranges == ["A1:B2"]
    assert projected.excel.members == {}


def test_empty_contract_id_preserves_canonical_profile_hash() -> None:
    implicit = DeliverableProfile(name="monthly")
    explicit = DeliverableProfile(name="monthly", contract_id="")

    assert profile_sha256(implicit) == profile_sha256(explicit)


def test_empty_comparison_prerequisites_preserve_canonical_profile_hash() -> None:
    implicit = DeliverableProfile(name="monthly")
    explicit = DeliverableProfile(
        name="monthly", excel=ExcelProfile(comparison_prerequisites=[])
    )

    assert profile_sha256(implicit) == profile_sha256(explicit)
    # Regression pin: fails if any new field ever changes this default hash.
    assert (
        profile_sha256(implicit)
        == "e7259f869ad0381dad15d726fab606b2c7438f9c8cf9fa9332596f75dfa1ace2"
    )


def test_empty_row_identity_rules_preserve_canonical_profile_hash() -> None:
    implicit_sheet = DeliverableProfile(
        name="monthly", excel=ExcelProfile(sheets={"Data": SheetProfile()})
    )
    explicit_empty = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(sheets={"Data": SheetProfile(row_identity_rules=[])}),
    )

    assert profile_sha256(implicit_sheet) == profile_sha256(explicit_empty)


def test_absent_row_identity_header_preserves_legacy_canonical_shape() -> None:
    rule = RowIdentityRule(anchor_cell="A1", identity_columns=["B"])

    assert "header_row" not in rule.model_dump(mode="json")


def test_new_profile_rules_reject_malformed_locations_at_model_boundary() -> None:
    from pydantic import ValidationError

    with pytest.raises(ValidationError):
        ComparisonPrerequisite(name="Scenario", sheet="Config", cell="not-a-cell")
    with pytest.raises(ValidationError):
        RowIdentityRule(anchor_cell="A1", identity_columns=["?"])
    with pytest.raises(ValidationError):
        RowIdentityRule(
            anchor_cell="A1",
            identity_columns=["B"],
            ordinal_columns=["b"],
        )


def test_comparison_prerequisites_project_into_a_package_member() -> None:
    profile = DeliverableProfile(
        name="pkg",
        excel=ExcelProfile(
            members={
                "ops": ExcelMemberProfile(
                    comparison_prerequisites=[
                        ComparisonPrerequisite(
                            name="Scenario", sheet="Config", cell="B2"
                        )
                    ]
                )
            }
        ),
    )

    projected = profile_for_excel_member(profile, "ops", 2)

    assert [item.name for item in projected.excel.comparison_prerequisites] == [
        "Scenario"
    ]


def test_empty_formula_engine_preserves_canonical_profile_hash() -> None:
    implicit = DeliverableProfile(name="monthly")
    explicit_default = DeliverableProfile(
        name="monthly", excel=ExcelProfile(formula_engine="auto")
    )

    assert profile_sha256(implicit) == profile_sha256(explicit_default)
    # Same regression pin as test_empty_comparison_prerequisites_preserve_
    # canonical_profile_hash: fails if formula_engine's default ever leaks
    # into the canonical JSON.
    assert (
        profile_sha256(implicit)
        == "e7259f869ad0381dad15d726fab606b2c7438f9c8cf9fa9332596f75dfa1ace2"
    )


def test_formula_engine_projects_into_a_package_member() -> None:
    profile = DeliverableProfile(
        name="pkg",
        excel=ExcelProfile(
            members={"ops": ExcelMemberProfile(formula_engine="native")}
        ),
    )

    projected = profile_for_excel_member(profile, "ops", 2)

    assert projected.excel.formula_engine == "native"


def test_unscoped_formula_engine_is_ambiguous_for_multi_workbook_package() -> None:
    profile = DeliverableProfile(
        name="monthly",
        excel=ExcelProfile(formula_engine="native"),
    )

    with pytest.raises(ValueError, match="ambiguous for this package"):
        excel_profile_for_member(profile, "ops", workbook_count=2)


MODULES = [
    "qc_tool",
    "qc_tool.config",
    "qc_tool.io",
    "qc_tool.excel",
    "qc_tool.ppt",
    "qc_tool.crosscheck",
    "qc_tool.triage",
    "qc_tool.report",
    "qc_tool.history",
    "qc_tool.ui",
]


@pytest.mark.parametrize("name", MODULES)
def test_module_imports(name: str) -> None:
    importlib.import_module(name)
