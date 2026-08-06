"""UI smoke tests (criterion 12): pages render; perform_run produces artifacts."""

import datetime as dt
import logging
import os
from pathlib import Path

import pytest
import yaml
from nicegui import app, events, ui
from nicegui.helpers import warnings as nicegui_warnings
from nicegui.testing import User

import qc_tool.run_service as run_service
import qc_tool.ui.app as app_module
from qc_tool.config.profile import DeliverableProfile, save_profile
from qc_tool.coverage import QCRunMode
from qc_tool.crosscheck.trace import MappingSuggestion, SuggestedSource
from qc_tool.engine import QCRunResult
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingEvidenceTag,
    FindingProvenance,
    FindingSubtype,
    FindingTemporalContext,
    GridExcerpt,
    Materiality,
    Severity,
)
from qc_tool.history.run_state import RunStateRecord, RunStateStore, RunStatus
from qc_tool.history.store import RunHistory
from qc_tool.progress import CancellationToken, ProgressEvent, RunCancelled, RunPhase
from qc_tool.review import build_pattern_groups, build_review_groups
from qc_tool.security import secure_managed_tree
from qc_tool.server_config import NetworkMode
from qc_tool.ui.app import (
    _acceptance_summary,
    _context_grid_html,
    _evidence_axes,
    _files_for_mode,
    _history_row,
    _input_cautions,
    _outcome_summary,
    _profile_path,
    _queue_status_line,
    _relative_time,
    _review_group_rows,
    _role_requirement,
    _run_blockers,
    _safe_upload_name,
    _scope_summary,
    _storage_secret,
    create_pages,
    list_profiles,
    load_profile_by_name,
    perform_run,
    persist_confirmed_mapping,
)
from qc_tool.ui.guide import PROFILE_CONTROLS_EXAMPLE
from qc_tool.ui.theme import CSS
from tests.conftest import fixture_profile

pytest_plugins = ["nicegui.testing.user_plugin"]


def test_mode_toggle_pins_content_color_against_quasar() -> None:
    # Quasar paints the pressed button bg-primary + text-white, and --q-primary
    # is light in dark mode, so every toggle needs its own pinned colours.
    assert ".q-btn-toggle .q-btn .q-btn__content" in CSS
    assert '.q-btn-toggle .q-btn[aria-pressed="true"] .q-btn__content' in CSS
    assert '.q-btn-toggle .q-btn[aria-pressed="true"] { background: var(--btn-bg)' in CSS
    for field in (
        "provenance",
        "subtype",
        "materiality",
        "temporal_context",
        "expected_reason",
        "evidence_tags",
    ):
        assert field in app_module.FINDINGS_BODY_SLOT


def test_evidence_axes_cover_every_typed_axis() -> None:
    finding = Finding(
        finding_id="F1",
        finding_class=FindingClass.VALUE_CHANGED,
        artifact="excel",
        message="value changed",
        severity=Severity.CRITICAL,
        sheet="Data",
        location="C7",
        baseline_value="1",
        current_value="2",
        element="cell",
        impacts=["Data!C8"],
        root_cause_key="rc-1",
        provenance=FindingProvenance.CHANGED,
        subtype=FindingSubtype.VALUE_REPLACEMENT,
        materiality=Materiality.MATERIAL,
        temporal_context=FindingTemporalContext.HISTORICAL,
        evidence_tags={FindingEvidenceTag.FORMULA_TEXT},
        analyst_comment="checked",
    )

    axes = dict(_evidence_axes(finding))

    for key in (
        "class",
        "artifact",
        "location",
        "baseline",
        "current",
        "element",
        "impacts",
        "root cause",
        "provenance",
        "subtype",
        "materiality",
        "temporal context",
        "evidence",
        "analyst comment",
    ):
        assert key in axes
    # Empty axes are dropped rather than rendered as blank rows.
    assert "waiver" not in axes
    assert "expected reason" not in axes


def test_profile_controls_guide_example_is_valid_yaml() -> None:
    payload = yaml.safe_load(PROFILE_CONTROLS_EXAMPLE)
    tie_outs = payload["excel"]["controls"]["tie_outs"]

    assert tie_outs[0]["components"] == ["Data!C2:C5"]
    assert [term["operation"] for term in tie_outs[1]["terms"]] == [
        "add",
        "subtract",
    ]


def test_review_group_rows_do_not_embed_atomic_member_payloads(qc_result) -> None:
    groups = build_review_groups(qc_result.findings)

    rows = _review_group_rows(groups)

    assert rows
    assert sum(group.member_count for group in groups) == len(qc_result.findings)
    assert all("member_ids" not in row and "excerpts" not in row for row in rows)


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

    monkeypatch.setattr(run_service, "write_excel_report", cancel_after_partial_report)

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


def test_queue_status_line_shows_state_without_source_paths() -> None:
    created = dt.datetime.now(dt.UTC) - dt.timedelta(seconds=12)
    record = RunStateRecord(
        request_id="0123456789abcdef",
        created_at=created,
        status=RunStatus.RUNNING,
        mode=QCRunMode.CYCLE_COMPARISON.value,
        profile="fixture",
        files={"current_excel": "current.xlsx"},
        phase=RunPhase.COMPARING_FORMULAS.value,
        processed=3,
        total=7,
    )

    line = _queue_status_line(record)

    assert "#01234567" in line
    assert "Running" in line
    assert "Comparing formulas (3/7)" in line
    assert "s elapsed" in line
    assert "current.xlsx" in line
    assert "\\" not in line and "/current.xlsx" not in line


def test_queue_status_line_reports_position_for_queued_requests() -> None:
    record = RunStateRecord(
        request_id="fedcba9876543210",
        created_at=dt.datetime.now(dt.UTC),
        status=RunStatus.QUEUED,
        mode=QCRunMode.CYCLE_COMPARISON.value,
        profile="fixture",
        queue_position=2,
    )

    assert "Queued" in _queue_status_line(record)
    assert "position 2" in _queue_status_line(record)


@pytest.mark.asyncio
async def test_run_page_reconnects_to_queue_state_from_another_tab(
    user: User, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    create_pages(work_dir)
    # A request submitted elsewhere: this page must show it, not resubmit it.
    RunStateStore(work_dir / "history.sqlite3").enqueue(
        "abcdef0123456789",
        mode=QCRunMode.CYCLE_COMPARISON.value,
        profile="fixture",
        files={"current_excel": "current.xlsx"},
        queue_position=1,
    )

    await user.open("/")

    await user.should_see("#abcdef01")
    await user.should_see("Cancel #abcdef01")


@pytest.mark.asyncio
async def test_guide_page_renders_packaged_operator_content(
    user: User, tmp_path: Path
) -> None:
    create_pages(tmp_path / "work")
    await user.open("/guide")
    await user.should_see("QC Tool guide")
    await user.should_see("Choose the right QC mode")
    await user.should_see("they do not block read-only QC")
    await user.should_see("Profiles, controls, and waivers")
    await user.should_see("Pattern review-item counts are analyst decisions")
    await user.should_see("Mass alone is a grouping detector")
    await user.should_see("implicit numeric refresh block is Warning")
    await user.should_see("Excel selected/total")
    await user.should_see("preflight, cycle-comparison, and final-package modes")
    await user.should_see("operation: subtract")
    await user.should_see("Availability controls blankness only")
    await user.should_see("Coverage and severity")
    await user.should_see("Safeguards are visible")
    await user.should_see("Excel to PowerPoint mappings")
    await user.should_see("Privacy, sharing, and attestations")
    await user.should_see("CLI and automation")
    await user.should_see("Troubleshooting")
    await user.should_see("Before sign-off")
    await user.should_see("No built-in authentication or TLS")
    # Search, common tasks, and the glossary are packaged and offline.
    await user.should_see("Search the guide")
    await user.should_see("Common tasks")
    await user.should_see("Glossary")
    await user.should_see("pattern group")
    # A worked example with dummy data, and the history shelf actions.
    await user.should_see("Worked example")
    await user.should_see("formula replaced by a constant")
    await user.should_see("Archive before you delete")


@pytest.mark.asyncio
async def test_history_page_renders(user: User, tmp_path: Path) -> None:
    create_pages(tmp_path / "work")
    await user.open("/history")
    await user.should_see("Run history")
    await user.should_see("No runs recorded yet.")


@pytest.mark.asyncio
async def test_history_page_offers_search_and_filters(
    user: User, fixture_dir: Path, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    perform_run(
        work_dir,
        {
            "baseline_excel": fixture_dir / "baseline.xlsx",
            "current_excel": fixture_dir / "current.xlsx",
        },
        {},
        fixture_profile(),
    )
    create_pages(work_dir)
    await user.open("/history")
    await user.should_see("search run id, profile, or file")
    for label in ("Mode", "Profile", "Capability", "Date"):
        await user.should_see(label)


def test_run_blockers_report_every_missing_role() -> None:
    assert _run_blockers(QCRunMode.CYCLE_COMPARISON, {}) == [
        "Upload a baseline and current Excel pair and/or PowerPoint pair"
    ]
    assert _run_blockers(QCRunMode.FINAL_PACKAGE, {"current_excel": Path("a.xlsx")}) == [
        "Upload both current Excel and current PowerPoint files"
    ]
    ready = {"baseline_excel": Path("b.xlsx"), "current_excel": Path("c.xlsx")}
    assert _run_blockers(QCRunMode.CYCLE_COMPARISON, ready) == []


def test_run_blockers_name_the_files_a_rerun_still_needs() -> None:
    blockers = _run_blockers(
        QCRunMode.CYCLE_COMPARISON,
        {"baseline_excel": Path("b.xlsx"), "current_excel": Path("c.xlsx")},
        rerun_of=7,
        rerun_required=frozenset({"baseline_excel", "current_excel", "current_ppt"}),
    )
    assert blockers == [
        "Re-QC of run #7 is blocked until these files are selected again: "
        "Current — PowerPoint deck"
    ]


def test_role_requirement_changes_with_mode() -> None:
    assert _role_requirement(QCRunMode.FINAL_PACKAGE, "current_excel") == "required"
    assert (
        _role_requirement(QCRunMode.CURRENT_FILE_PREFLIGHT, "current_ppt")
        == "Excel and/or PPT"
    )
    assert _role_requirement(QCRunMode.CYCLE_COMPARISON, "baseline_excel") == "Excel pair"
    assert _role_requirement(QCRunMode.CYCLE_COMPARISON, "baseline_ppt") == "PPT pair"


def test_policy_summaries_describe_what_the_engine_receives() -> None:
    state = app_module.SessionState()
    assert _scope_summary(state) == "scope: everything"
    assert _acceptance_summary(state) == "acceptance: strict"
    state.available_sheets = ["A", "B", "C"]
    state.selected_sheets = {"A"}
    state.acceptance_absolute = 2.5
    state.acceptance_percent = 0.1
    assert _scope_summary(state) == "scope: 1/3 sheets"
    assert _acceptance_summary(state) == "acceptance: ±2.5 or ±0.1%"
    # A full selection is not a narrowed scope, so it must not be disclosed as one.
    state.selected_sheets = {"A", "B", "C"}
    assert _scope_summary(state) == "scope: everything"


def test_outcome_summary_leads_with_the_decision() -> None:
    assert _outcome_summary({Severity.CRITICAL: 3, Severity.WARNING: 4})[:2] == (
        "attention",
        "Review required",
    )
    assert _outcome_summary({Severity.WARNING: 2})[:2] == ("limited", "Review required")
    assert _outcome_summary({Severity.INFO: 9})[:2] == ("ok", "No blocking findings")


def test_context_grid_html_escapes_source_values() -> None:
    excerpt = GridExcerpt(
        rows=[1],
        cols=["<A>"],
        cells=[["<script>alert(1)</script>"]],
        hit_row=0,
        hit_col=0,
    )
    markup = _context_grid_html(excerpt, "baseline")
    assert "<script>" not in markup
    assert "&lt;script&gt;alert(1)&lt;/script&gt;" in markup
    assert "&lt;A&gt;" in markup
    assert 'class="hit"' in markup


def test_relative_time_keeps_the_exact_stamp_available() -> None:
    now = dt.datetime.now(dt.UTC)
    assert _relative_time(now) == "just now"
    assert _relative_time(now - dt.timedelta(minutes=20)) == "20 min ago"
    assert _relative_time(now - dt.timedelta(hours=5)) == "5 h ago"
    old = now - dt.timedelta(days=400)
    assert _relative_time(old) == old.date().isoformat()


def test_history_row_exposes_mode_profile_capability_and_decisions(
    fixture_dir: Path, tmp_path: Path
) -> None:
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
    record = RunHistory(work_dir / "history.sqlite3").get_run(artifacts.run_id)

    row = _history_row(record)

    assert row["id"] == artifacts.run_id
    assert row["mode_key"] == QCRunMode.CYCLE_COMPARISON.value
    assert row["profile"] == "fixture"
    assert row["capability"] in {"limited", "complete"}
    decisions = row["decisions"]
    assert isinstance(decisions, list)
    assert decisions and all(set(entry) == {"k", "n"} for entry in decisions)
    assert row["started"] == record.started_at.isoformat(timespec="seconds")
    assert "baseline.xlsx" in str(row["files"])


@pytest.mark.asyncio
async def test_colophon_is_scoped_to_the_guide(user: User, tmp_path: Path) -> None:
    create_pages(tmp_path / "work")
    await user.open("/")
    await user.should_see("source files are never modified")
    await user.should_not_see("Himanshu")

    await user.open("/guide")
    await user.should_see("Curated by Himanshu")


def test_input_cautions_flag_a_baseline_current_mix_up() -> None:
    state = app_module.SessionState()
    assert _input_cautions(state) == []
    state.files = {
        "baseline_excel": Path("/tmp/a/pack.xlsx"),
        "current_excel": Path("/tmp/b/pack.xlsx"),
    }
    state.file_sizes = {"baseline_excel": 4096, "current_excel": 4096}

    cautions = _input_cautions(state)

    assert len(cautions) == 1
    assert "pack.xlsx" in cautions[0]
    assert "prove nothing" in cautions[0]
    # A genuinely different current file is not a caution.
    state.file_sizes["current_excel"] = 5120
    assert _input_cautions(state) == []


def test_input_groups_separate_baseline_from_current() -> None:
    groups = {group: roles for group, _, _, roles in app_module.ROLE_GROUPS}

    assert groups["baseline"] == ("baseline_excel", "baseline_ppt")
    assert groups["current"] == ("current_excel", "current_ppt")
    assert set(groups["baseline"]) | set(groups["current"]) == set(app_module.ROLES)


@pytest.mark.asyncio
async def test_inputs_render_as_baseline_and_current_panels(
    user: User, tmp_path: Path
) -> None:
    create_pages(tmp_path / "work")
    await user.open("/")
    await user.should_see("Baseline")
    await user.should_see("the previous cycle you compare against")
    await user.should_see("Current")
    await user.should_see("the cycle you are signing off")


@pytest.mark.asyncio
async def test_shutdown_control_confirms_before_stopping(
    user: User, tmp_path: Path
) -> None:
    create_pages(tmp_path / "work")
    await user.open("/")

    user.find(marker="quit").click()

    await user.should_see("Stop the QC Tool server?")
    await user.should_see("Stop server")
    await user.should_see("Keep running")


@pytest.mark.asyncio
async def test_shutdown_dialog_discloses_unfinished_work(
    user: User, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    create_pages(work_dir)
    RunStateStore(work_dir / "history.sqlite3").enqueue(
        "abcdef0123456789",
        mode=QCRunMode.CYCLE_COMPARISON.value,
        profile="fixture",
        files={"current_excel": "current.xlsx"},
        queue_position=1,
    )

    await user.open("/")
    user.find(marker="quit").click()

    await user.should_see("recorded as unfinished")
    await user.should_see("#abcdef01")


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
    await user.should_see("Review queue")
    await user.should_see("Atomic evidence")
    await user.should_see("atomic findings")
    # Review queue is the default view; evidence tabs are opt-in.
    panels = user.find(kind=ui.tab_panels).elements.pop()
    assert panels.value == "review"


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


def _emit(element: ui.element, event_type: str, args: object) -> None:
    """Drive a slot-emitted table event the way the browser would."""
    for listener in element._event_listeners.values():
        if listener.type == event_type and listener.handler is not None:
            listener.handler(
                events.GenericEventArguments(
                    sender=element, client=element.client, args=args
                )
            )
            return
    raise AssertionError(f"no {event_type!r} listener on {element}")


@pytest.mark.asyncio
async def test_group_review_survives_the_panel_refresh_it_triggers(
    user: User,
    fixture_dir: Path,
    tmp_path: Path,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Applying a group review refreshes the panel that owns the dialog.

    Reporting or closing after that refresh touches deleted elements, which
    NiceGUI surfaces as a use-after-free warning and a RuntimeError.
    """
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

    group_table = next(
        element
        for element in user.find(kind=ui.table).elements
        if "review-groups-table" in element.classes
    )
    _emit(group_table, "select", {"id": str(group_table.rows[0]["id"])})
    await user.should_see("Review group")

    user.find("Review group").click()
    await user.should_see("Apply review")
    severity_select = next(
        element
        for element in user.find(kind=ui.select).elements
        if element.props.get("label") == "Group severity"
    )
    severity_select.value = Severity.INFO.value

    # Closing or reporting after the refresh touches deregistered elements,
    # which NiceGUI reports by logging rather than by raising to the caller.
    nicegui_warnings.reset()  # the use-after-free warning is otherwise once-only
    with caplog.at_level(logging.WARNING, logger="nicegui"):
        user.find("Apply review").click()

    assert not caplog.records, [record.getMessage() for record in caplog.records]


def test_comment_only_review_is_visible_in_the_queue(qc_result: QCRunResult) -> None:
    """Choosing "keep" records a decision without setting severity_overridden.

    Nothing in the queue used to indicate that, so the decision looked lost.
    """
    findings = [finding.model_copy(deep=True) for finding in qc_result.findings]
    findings[0].analyst_comment = "checked against the source pack"

    rows = _review_group_rows(build_pattern_groups(findings))
    reviewed = {str(row["id"]): int(str(row["reviewed"])) for row in rows}
    members = {str(row["id"]): int(str(row["members"])) for row in rows}

    assert any(reviewed.values()), "a comment-only decision must mark its group reviewed"
    assert all(reviewed[key] <= members[key] for key in reviewed)
    assert sum(reviewed.values()) == 1


def test_untouched_findings_are_not_marked_reviewed(qc_result: QCRunResult) -> None:
    findings = [finding.model_copy(deep=True) for finding in qc_result.findings]

    rows = _review_group_rows(build_pattern_groups(findings))

    assert all(row["reviewed"] == 0 for row in rows)


def test_row_slots_show_a_note_without_a_severity_override() -> None:
    for slot in (app_module.REVIEW_MEMBER_ROWS_SLOT, app_module.FINDINGS_BODY_SLOT):
        assert 'v-else-if="props.row.comment"' in slot
    assert 'v-if="props.row.reviewed"' in app_module.REVIEW_GROUPS_BODY_SLOT
