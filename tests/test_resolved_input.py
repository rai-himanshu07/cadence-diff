"""Frozen contract tests for the per-run resolved input configuration
(``ResolvedInputConfigurationV1``, ``qc_tool.config.resolved_input``).

Plan: docs/plans/plan-20260913-mode-aware-configuration-wizard.md, Step 1.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook
from pydantic import ValidationError

from qc_tool.config.input_contract import INPUT_CONTRACT_VERSION
from qc_tool.config.profile import DeliverableProfile, profile_sha256
from qc_tool.config.resolved_input import (
    RESOLVED_INPUT_CONFIGURATION_VERSION,
    ConfigurationCoverageState,
    ResolvedColumn,
    ResolvedInputConfigurationV1,
    ResolvedMember,
    ResolvedRegion,
    ResolvedSelector,
    ResolvedSheet,
    StaleResolvedConfigurationError,
    validate_freshness,
)
from qc_tool.coverage import QCRunMode


def _member(**overrides: object) -> ResolvedMember:
    defaults: dict[str, object] = {
        "member_id": "primary",
        "baseline_source_sha256": "a" * 64,
        "current_source_sha256": "b" * 64,
    }
    defaults.update(overrides)
    return ResolvedMember(**defaults)  # type: ignore[arg-type]


# --- canonical digest is independent of profile_sha256 -------------------


def test_resolved_digest_lives_in_a_separate_hash_space_from_profile_sha256() -> None:
    profile_hash = profile_sha256(DeliverableProfile(name="monthly"))
    resolved = ResolvedInputConfigurationV1.legacy_default(
        profile_name="monthly", profile_sha256=profile_hash
    )
    assert resolved.canonical_sha256() != profile_hash
    assert resolved.profile_sha256 == profile_hash


def test_resolved_digest_changes_when_a_sheet_resolution_changes_but_not_profile_hash() -> None:
    profile_hash = profile_sha256(DeliverableProfile(name="monthly"))
    first = ResolvedInputConfigurationV1(
        profile_name="monthly",
        profile_sha256=profile_hash,
        members=(_member(sheets=(ResolvedSheet(sheet_id="data", current_sheet_name="Data"),)),),
    )
    second = ResolvedInputConfigurationV1(
        profile_name="monthly",
        profile_sha256=profile_hash,
        members=(_member(sheets=(ResolvedSheet(sheet_id="data", current_sheet_name="Ledger"),)),),
    )
    assert first.canonical_sha256() != second.canonical_sha256()
    assert first.profile_sha256 == second.profile_sha256 == profile_hash


def test_resolved_digest_is_deterministic() -> None:
    first = ResolvedInputConfigurationV1(members=(_member(),))
    second = ResolvedInputConfigurationV1(members=(_member(),))
    assert first.canonical_sha256() == second.canonical_sha256()
    assert len(first.canonical_sha256()) == 64


# --- one-to-one sheet resolution -----------------------------------------


def test_baseline_sheet_names_must_resolve_one_to_one_within_a_member() -> None:
    with pytest.raises(ValidationError):
        _member(
            sheets=(
                ResolvedSheet(sheet_id="a", baseline_sheet_name="Data"),
                ResolvedSheet(sheet_id="b", baseline_sheet_name="Data"),
            )
        )


def test_current_sheet_names_must_resolve_one_to_one_within_a_member() -> None:
    with pytest.raises(ValidationError):
        _member(
            sheets=(
                ResolvedSheet(sheet_id="a", current_sheet_name="Data"),
                ResolvedSheet(sheet_id="b", current_sheet_name="Data"),
            )
        )


def test_distinct_logical_sheets_may_resolve_to_distinct_physical_names() -> None:
    member = _member(
        sheets=(
            ResolvedSheet(
                sheet_id="a", baseline_sheet_name="Data", current_sheet_name="Data"
            ),
            ResolvedSheet(
                sheet_id="b",
                baseline_sheet_name="Summary",
                current_sheet_name="Summary2026",
            ),
        )
    )
    assert len(member.sheets) == 2


def test_member_ids_must_be_unique_within_a_resolved_configuration() -> None:
    member = _member()
    with pytest.raises(ValidationError):
        ResolvedInputConfigurationV1(members=(member, member))


# --- absent-contract legacy default ---------------------------------------


def test_legacy_default_represents_a_run_with_no_saved_input_contract() -> None:
    resolved = ResolvedInputConfigurationV1.legacy_default(
        profile_name="default", profile_sha256="", mode=QCRunMode.CYCLE_COMPARISON
    )
    assert resolved.source == "legacy_default"
    assert resolved.members == ()
    assert resolved.mode is QCRunMode.CYCLE_COMPARISON


def test_default_source_is_input_contract_when_not_specified() -> None:
    resolved = ResolvedInputConfigurationV1(members=(_member(),))
    assert resolved.source == "input_contract"


# --- configuration coverage states ----------------------------------------


@pytest.mark.parametrize(
    "state",
    [
        "confirmed",
        "inherited",
        "automatic_confirmed",
        "positional",
        "excluded",
        "degraded_acknowledged",
    ],
)
def test_all_six_configuration_coverage_states_are_accepted(
    state: ConfigurationCoverageState,
) -> None:
    region = ResolvedRegion(region_id="r1", coverage=state)
    assert region.coverage == state


# --- no filesystem paths, no selector/preview values ----------------------


def test_resolved_member_never_carries_a_filesystem_path_field() -> None:
    for field_name in ResolvedMember.model_fields:
        assert "path" not in field_name.lower()


def test_resolved_selector_never_carries_the_actual_cell_value() -> None:
    fields = set(ResolvedSelector.model_fields)
    for forbidden in ("value", "baseline_value", "current_value", "expected_value"):
        assert forbidden not in fields
    selector = ResolvedSelector(
        selector_id="scenario",
        baseline_cell="B2",
        current_cell="B2",
        equal=True,
    )
    assert "value" not in selector.model_dump()


def test_resolved_region_never_carries_revealed_formula_text() -> None:
    for field_name in ResolvedRegion.model_fields:
        assert "formula" not in field_name.lower()


# --- staleness rejection ---------------------------------------------------


def test_validate_freshness_passes_when_hashes_and_version_match() -> None:
    resolved = ResolvedInputConfigurationV1(
        inspection_contract_version=INPUT_CONTRACT_VERSION,
        members=(_member(),),
    )
    validate_freshness(
        resolved,
        current_source_sha256={"primary": ("a" * 64, "b" * 64)},
    )


def test_validate_freshness_rejects_a_changed_source_hash() -> None:
    resolved = ResolvedInputConfigurationV1(
        inspection_contract_version=INPUT_CONTRACT_VERSION,
        members=(_member(),),
    )
    with pytest.raises(StaleResolvedConfigurationError):
        validate_freshness(
            resolved,
            current_source_sha256={"primary": ("a" * 64, "c" * 64)},
        )


def test_validate_freshness_rejects_a_missing_member() -> None:
    resolved = ResolvedInputConfigurationV1(
        inspection_contract_version=INPUT_CONTRACT_VERSION,
        members=(_member(),),
    )
    with pytest.raises(StaleResolvedConfigurationError):
        validate_freshness(resolved, current_source_sha256={})


def test_validate_freshness_rejects_a_stale_contract_schema_version() -> None:
    resolved = ResolvedInputConfigurationV1(
        inspection_contract_version=INPUT_CONTRACT_VERSION - 1 if INPUT_CONTRACT_VERSION > 0 else 0,
        members=(_member(),),
    )
    with pytest.raises(StaleResolvedConfigurationError):
        validate_freshness(
            resolved,
            current_source_sha256={"primary": ("a" * 64, "b" * 64)},
        )


# --- frozen / immutable ----------------------------------------------------


def test_resolved_configuration_is_frozen() -> None:
    resolved = ResolvedInputConfigurationV1()
    with pytest.raises(ValidationError):
        resolved.profile_name = "changed"  # type: ignore[misc]


def test_resolved_column_is_frozen() -> None:
    column = ResolvedColumn(column_id="c1")
    with pytest.raises(ValidationError):
        column.alignment_role = "identity"  # type: ignore[misc]


def test_resolved_configuration_version_is_frozen_at_one() -> None:
    assert RESOLVED_INPUT_CONFIGURATION_VERSION == 1
    assert ResolvedInputConfigurationV1().version == 1


# --- perform_run() actually calls validate_freshness() (plan Step 10) -----


def _write_pair(tmp_path: Path) -> tuple[Path, Path]:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    for path, value in ((baseline_path, 1), (current_path, 2)):
        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "Data"
        sheet["A1"] = value
        workbook.save(path)
    return baseline_path, current_path


def test_perform_run_rejects_a_resolved_configuration_with_a_stale_source_hash(
    tmp_path: Path,
) -> None:
    """Step 10's "source-hash-bound request rejection" criterion: a saved
    resolved configuration whose declared source hash no longer matches the
    file actually being loaded (e.g. the file changed since the workspace
    resolved it) must be refused before any QC work happens -- not silently
    applied against different data. `validate_freshness()` existed since
    Step 1 but was never actually called by any production caller until now.
    """
    from qc_tool.history.store import sha256_file
    from qc_tool.run_service import perform_run

    baseline_path, current_path = _write_pair(tmp_path)
    real_current_hash = sha256_file(current_path)
    resolved = ResolvedInputConfigurationV1(
        inspection_contract_version=INPUT_CONTRACT_VERSION,
        members=(
            ResolvedMember(
                member_id="primary",
                baseline_source_sha256=sha256_file(baseline_path),
                # Deliberately wrong -- simulates the current file having
                # changed since this configuration was resolved.
                current_source_sha256="0" * 64,
            ),
        ),
    )
    assert resolved.members[0].current_source_sha256 != real_current_hash

    with pytest.raises(StaleResolvedConfigurationError):
        perform_run(
            tmp_path / "work",
            {"baseline_excel": baseline_path, "current_excel": current_path},
            {},
            DeliverableProfile(name="default"),
            mode=QCRunMode.CYCLE_COMPARISON,
            resolved_input_configuration=resolved,
        )


def test_perform_run_accepts_a_resolved_configuration_with_matching_hashes(
    tmp_path: Path,
) -> None:
    """The inverse control: a resolved configuration whose declared hashes
    DO match the real files must not be rejected by the freshness check.
    """
    from qc_tool.history.store import sha256_file
    from qc_tool.run_service import perform_run

    baseline_path, current_path = _write_pair(tmp_path)
    resolved = ResolvedInputConfigurationV1(
        inspection_contract_version=INPUT_CONTRACT_VERSION,
        members=(
            ResolvedMember(
                member_id="primary",
                baseline_source_sha256=sha256_file(baseline_path),
                current_source_sha256=sha256_file(current_path),
            ),
        ),
    )

    artifacts = perform_run(
        tmp_path / "work",
        {"baseline_excel": baseline_path, "current_excel": current_path},
        {},
        DeliverableProfile(name="default"),
        mode=QCRunMode.CYCLE_COMPARISON,
        resolved_input_configuration=resolved,
    )
    assert artifacts.result is not None
