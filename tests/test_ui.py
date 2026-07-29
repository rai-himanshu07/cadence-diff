"""UI smoke tests (criterion 12): pages render; perform_run produces artifacts."""

import datetime as dt
import os
from pathlib import Path

import pytest
from nicegui.testing import User

import qc_tool.ui.app as app_module
from qc_tool.config.profile import DeliverableProfile, save_profile
from qc_tool.coverage import QCRunMode
from qc_tool.crosscheck.trace import MappingSuggestion, SuggestedSource
from qc_tool.findings import Severity
from qc_tool.history.store import RunHistory
from qc_tool.progress import CancellationToken, ProgressEvent, RunCancelled, RunPhase
from qc_tool.security import secure_managed_tree
from qc_tool.server_config import NetworkMode
from qc_tool.ui.app import (
    _files_for_mode,
    _profile_path,
    _safe_upload_name,
    _storage_secret,
    create_pages,
    list_profiles,
    load_profile_by_name,
    perform_run,
    persist_confirmed_mapping,
)
from qc_tool.ui.theme import CSS
from tests.conftest import fixture_profile

pytest_plugins = ["nicegui.testing.user_plugin"]


def test_mode_toggle_pins_content_color_against_quasar() -> None:
    assert '.mode-select .q-btn .q-btn__content' in CSS
    assert '.mode-select .q-btn[aria-pressed="true"] .q-btn__content' in CSS


def test_managed_names_cannot_escape_storage(tmp_path: Path) -> None:
    assert _safe_upload_name("../../quarterly.xlsx") == "quarterly.xlsx"
    assert _safe_upload_name(r"..\..\quarterly.xlsx") == "quarterly.xlsx"
    assert _profile_path(tmp_path, "monthly-pack") == tmp_path / "monthly-pack.yaml"
    with pytest.raises(ValueError, match="profile name"):
        _profile_path(tmp_path, "../../outside")


def test_storage_secret_is_private_and_stable(tmp_path: Path) -> None:
    first = _storage_secret(tmp_path)
    second = _storage_secret(tmp_path)

    assert first == second
    assert len(first) >= 32
    if os.name == "posix":
        assert (tmp_path / ".nicegui-storage-secret").stat().st_mode & 0o777 == 0o600


def test_existing_managed_tree_is_migrated_private(tmp_path: Path) -> None:
    nested = tmp_path / "runs" / "old"
    nested.mkdir(parents=True)
    artifact = nested / "report.html"
    artifact.write_text("private", encoding="utf-8")
    if os.name == "posix":
        nested.chmod(0o775)
        artifact.chmod(0o644)

    secure_managed_tree(tmp_path)

    if os.name == "posix":
        assert nested.stat().st_mode & 0o777 == 0o700
        assert artifact.stat().st_mode & 0o777 == 0o600


def test_profile_listing_and_loading(tmp_path: Path) -> None:
    profiles_dir = tmp_path / "profiles"
    assert list_profiles(profiles_dir) == ["default"]
    save_profile(DeliverableProfile(name="monthly-pack"), profiles_dir / "monthly-pack.yaml")
    assert list_profiles(profiles_dir) == ["default", "monthly-pack"]
    loaded = load_profile_by_name(profiles_dir, "monthly-pack")
    assert loaded.name == "monthly-pack"
    assert load_profile_by_name(profiles_dir, "default").name == "default"


def test_mode_file_requirements(tmp_path: Path) -> None:
    files = {
        "baseline_excel": tmp_path / "base.xlsx",
        "current_excel": tmp_path / "current.xlsx",
        "current_ppt": tmp_path / "current.pptx",
    }
    assert set(_files_for_mode(QCRunMode.CYCLE_COMPARISON, files)) == {
        "baseline_excel",
        "current_excel",
    }
    assert set(_files_for_mode(QCRunMode.CURRENT_FILE_PREFLIGHT, files)) == {
        "current_excel",
        "current_ppt",
    }
    assert set(_files_for_mode(QCRunMode.FINAL_PACKAGE, files)) == {
        "current_excel",
        "current_ppt",
    }
    with pytest.raises(ValueError, match="both current"):
        _files_for_mode(
            QCRunMode.FINAL_PACKAGE, {"current_excel": tmp_path / "current.xlsx"}
        )


def test_confirmed_mapping_persists_to_named_profile(tmp_path: Path) -> None:
    profiles_dir = tmp_path / "profiles"
    profiles_dir.mkdir()
    save_profile(DeliverableProfile(name="monthly"), profiles_dir / "monthly.yaml")
    suggestion = MappingSuggestion(
        slide="Summary",
        line="Revenue $1.2M",
        line_skeleton="Revenue $#M",
        figure_index=0,
        figure_raw="$1.2M",
    )
    candidate = SuggestedSource(
        sheet="Dashboard",
        cell="B2",
        value=1_200_000,
        display_match=True,
        label_score=100,
        rel_diff=0,
        row_label="Revenue",
        column_label="Jun-26",
    )

    mapping = persist_confirmed_mapping(
        profiles_dir, "monthly", suggestion, candidate
    )

    assert mapping.source_sheet == "Dashboard" and mapping.source_cell == "B2"
    stored = load_profile_by_name(profiles_dir, "monthly")
    assert stored.crosscheck.mappings == [mapping]
    with pytest.raises(ValueError, match="named profile"):
        persist_confirmed_mapping(profiles_dir, "default", suggestion, candidate)


def test_perform_run_produces_artifacts(fixture_dir: Path, tmp_path: Path) -> None:
    work_dir = tmp_path / "work"
    files = {
        "baseline_excel": fixture_dir / "baseline.xlsx",
        "current_excel": fixture_dir / "current.xlsx",
        "baseline_ppt": fixture_dir / "baseline.pptx",
        "current_ppt": fixture_dir / "current.pptx",
    }
    artifacts = perform_run(work_dir, files, {}, fixture_profile())

    assert artifacts.result.counts[Severity.CRITICAL] > 0
    for path in artifacts.report_paths.values():
        assert path.exists() and path.stat().st_size > 0

    history = RunHistory(work_dir / "history.sqlite3")
    record = history.get_run(artifacts.run_id)
    assert record.profile == "fixture"
    assert set(record.file_hashes) == set(files)


def test_perform_run_reports_ordered_progress(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    events: list[ProgressEvent] = []

    perform_run(
        tmp_path / "work",
        {"current_excel": fixture_dir / "current.xlsx"},
        {},
        DeliverableProfile(name="progress"),
        mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
        on_progress=events.append,
    )

    completed_phases = [
        event.phase
        for event in events
        if event.total and event.processed == event.total
    ]
    assert completed_phases == [
        RunPhase.PREPARING,
        RunPhase.LOADING_CURRENT_EXCEL,
        RunPhase.ANALYZING_EXCEL,
        RunPhase.WRITING_REPORTS,
        RunPhase.RECORDING_HISTORY,
        RunPhase.COMPLETE,
    ]


def test_cancelled_report_cleans_owned_directory_and_records_no_history(
    fixture_dir: Path,
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    work_dir = tmp_path / "work"
    token = CancellationToken()

    def cancel_after_partial_report(result, path: Path) -> None:
        path.write_text("partial", encoding="utf-8")
        token.cancel()

    monkeypatch.setattr(app_module, "write_excel_report", cancel_after_partial_report)

    with pytest.raises(RunCancelled):
        perform_run(
            work_dir,
            {"current_excel": fixture_dir / "current.xlsx"},
            {},
            DeliverableProfile(name="cancelled"),
            mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
            cancellation_token=token,
        )

    runs_dir = work_dir / "runs"
    assert runs_dir.is_dir()
    assert list(runs_dir.iterdir()) == []
    assert not (work_dir / "history.sqlite3").exists()


def test_pre_cancelled_run_creates_no_artifacts(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    token = CancellationToken()
    token.cancel()
    work_dir = tmp_path / "work"

    with pytest.raises(RunCancelled):
        perform_run(
            work_dir,
            {"current_excel": fixture_dir / "current.xlsx"},
            {},
            DeliverableProfile(name="cancelled"),
            mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
            cancellation_token=token,
        )

    assert not (work_dir / "runs").exists()
    assert not (work_dir / "history.sqlite3").exists()


def test_immediate_runs_get_distinct_private_report_paths(
    fixture_dir: Path, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    files = {"current_excel": fixture_dir / "current.xlsx"}
    first = perform_run(
        work_dir,
        files,
        {},
        DeliverableProfile(name="first"),
        mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
    )
    second = perform_run(
        work_dir,
        files,
        {},
        DeliverableProfile(name="second"),
        mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
    )

    assert first.report_paths != second.report_paths
    assert first.report_paths["html"].parent != second.report_paths["html"].parent
    if os.name == "posix":
        assert first.report_paths["html"].stat().st_mode & 0o777 == 0o600
        assert first.report_paths["html"].parent.stat().st_mode & 0o777 == 0o700


def test_perform_run_with_encrypted_current(
    fixture_dir: Path, tmp_path: Path, manifest
) -> None:
    files = {
        "baseline_excel": fixture_dir / "baseline.xlsx",
        "current_excel": fixture_dir / "current_encrypted.xlsx",
    }
    artifacts = perform_run(
        tmp_path / "work",
        files,
        {"current_excel": manifest.password},
        DeliverableProfile(name="enc-test"),
    )
    assert artifacts.result.counts[Severity.CRITICAL] > 0


@pytest.mark.asyncio
async def test_main_page_renders(user: User, tmp_path: Path) -> None:
    create_pages(tmp_path / "work")
    await user.open("/")
    await user.should_see("QC Tool")
    await user.should_see("Run QC")
    await user.should_see("Deliverable profile")
    await user.should_see("Override large-workbook refusal")
    await user.should_see("Current-file preflight")
    await user.should_see("Cycle comparison")
    await user.should_see("Final-package QC")
    await user.should_see("Guide")
    await user.should_see("How to choose a mode")


@pytest.mark.asyncio
async def test_guide_page_renders_packaged_operator_content(
    user: User, tmp_path: Path
) -> None:
    create_pages(tmp_path / "work")
    await user.open("/guide")
    await user.should_see("QC Tool guide")
    await user.should_see("Choose the right QC mode")
    await user.should_see("Profiles, controls, and waivers")
    await user.should_see("Availability controls blankness only")
    await user.should_see("Coverage and severity")
    await user.should_see("Safeguards are visible")
    await user.should_see("Excel to PowerPoint mappings")
    await user.should_see("Privacy, sharing, and attestations")
    await user.should_see("CLI and automation")
    await user.should_see("Troubleshooting")
    await user.should_see("Before sign-off")
    await user.should_see("No built-in authentication or TLS")


@pytest.mark.asyncio
async def test_history_page_renders(user: User, tmp_path: Path) -> None:
    create_pages(tmp_path / "work")
    await user.open("/history")
    await user.should_see("Run history")
    await user.should_see("No runs recorded yet.")


@pytest.mark.asyncio
async def test_lan_mode_shows_exposure_warning(user: User, tmp_path: Path) -> None:
    create_pages(
        tmp_path / "work",
        network_mode=NetworkMode.LAN,
        expires_at=dt.datetime.now(dt.UTC) + dt.timedelta(minutes=10),
    )
    await user.open("/")
    await user.should_see("NETWORK ACCESS ENABLED")
    await user.should_see("no built-in authentication or TLS")


@pytest.mark.asyncio
async def test_dark_mode_toggle_persists(user: User, tmp_path: Path) -> None:
    from nicegui import app

    app.storage.general.pop("dark_mode", None)
    create_pages(tmp_path / "work")
    await user.open("/")
    user.find(marker="dark-toggle").click()
    await user.should_see(marker="dark-toggle")  # page still alive after toggle
    assert app.storage.general.get("dark_mode") is True
    user.find(marker="dark-toggle").click()
    await user.should_see(marker="dark-toggle")
    assert app.storage.general.get("dark_mode") is False


@pytest.mark.asyncio
async def test_run_detail_page(user: User, fixture_dir: Path, tmp_path: Path) -> None:
    work_dir = tmp_path / "work"
    artifacts = perform_run(
        work_dir,
        {
            "baseline_excel": fixture_dir / "baseline.xlsx",
            "current_excel": fixture_dir / "current.xlsx",
        },
        {},
        fixture_profile(),
    )
    create_pages(work_dir)
    await user.open(f"/runs/{artifacts.run_id}")
    await user.should_see(f"Run #{artifacts.run_id}")
    await user.should_see("profile 'fixture'")
    await user.should_see("cross-checks ok")


@pytest.mark.asyncio
async def test_run_detail_page_not_found(user: User, tmp_path: Path) -> None:
    create_pages(tmp_path / "work")
    await user.open("/runs/999")
    await user.should_see("Run not found")


@pytest.mark.asyncio
async def test_final_package_detail_shows_mapping_review(
    user: User, fixture_dir: Path, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    profiles_dir = work_dir / "profiles"
    profiles_dir.mkdir(parents=True)
    profile = fixture_profile()
    save_profile(profile, profiles_dir / "fixture.yaml")
    artifacts = perform_run(
        work_dir,
        {
            "current_excel": fixture_dir / "current.xlsx",
            "current_ppt": fixture_dir / "current.pptx",
        },
        {},
        profile,
        mode=QCRunMode.FINAL_PACKAGE,
    )

    create_pages(work_dir)
    await user.open(f"/runs/{artifacts.run_id}")
    await user.should_see("Final-package QC")
    await user.should_see("Check coverage")
    await user.should_see("Mapping review")
    await user.should_see("eligible")
