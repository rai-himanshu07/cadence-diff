"""End-to-end NiceGUI tests for the full-page mode-aware configuration
workspace (plan-20260913, Step 7): the ``/configure`` route, auto-start
scan, region rendering, and bulk confirmation, exercised through the real
``qc_tool.ui.app.create_pages()`` page registration and a real (tiny)
subprocess-based setup scan -- not a mock.
"""

from __future__ import annotations

from pathlib import Path

import pytest
from nicegui import ui
from nicegui.testing import User
from openpyxl import Workbook

from qc_tool.coverage import QCRunMode
from qc_tool.history.config_session import ConfigSessionStore, session_key_for
from qc_tool.history.store import sha256_file
from qc_tool.ui.app import create_pages
from qc_tool.ui.config_workspace import build_session_choices, get_setup_job_registry

pytest_plugins = ["nicegui.testing.user_plugin"]

#: A real subprocess spawn is far slower than should_see's default ~0.3s
#: retry budget; 50 retries (~5s) matches this project's own established
#: convention for waiting on real subprocess-backed UI updates.
_SUBPROCESS_RETRIES = 50


def _write_simple_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet.append(["ID", "Value"])
    sheet.append(["A0", 10])
    sheet.append(["A1", 20])
    sheet.append(["A2", 30])
    workbook.save(path)


def _stage_session(work_dir: Path) -> str:
    files = {
        "baseline_excel": work_dir / "uploads" / "baseline_excel" / "baseline.xlsx",
        "current_excel": work_dir / "uploads" / "current_excel" / "current.xlsx",
    }
    for path in files.values():
        path.parent.mkdir(parents=True, exist_ok=True)
        _write_simple_workbook(path)
    file_hashes = {role: sha256_file(path) for role, path in files.items()}
    choices = build_session_choices(
        mode=QCRunMode.CYCLE_COMPARISON,
        profile_name="default",
        files={role: str(path) for role, path in files.items()},
        file_hashes=file_hashes,
        output_mode="decision",
        allow_large_workbooks=False,
        allow_dependency_indexing=False,
        acceptance_absolute=0.0,
        acceptance_percent=0.0,
        rerun_of=None,
    )
    session_key = session_key_for(file_hashes)
    ConfigSessionStore(work_dir / "history.sqlite3").save_choices(
        session_key, profile_name="default", choices=choices
    )
    return session_key


def _write_simple_deck(path: Path, titles: list[str]) -> None:
    from pptx import Presentation

    deck = Presentation()
    layout = deck.slide_layouts[0]
    for title in titles:
        slide = deck.slides.add_slide(layout)
        title_shape = slide.shapes.title
        assert title_shape is not None
        title_shape.text = title
    deck.save(str(path))


def _stage_session_with_ppt(work_dir: Path) -> str:
    files = {
        "baseline_excel": work_dir / "uploads" / "baseline_excel" / "baseline.xlsx",
        "current_excel": work_dir / "uploads" / "current_excel" / "current.xlsx",
        "baseline_ppt": work_dir / "uploads" / "baseline_ppt" / "baseline.pptx",
        "current_ppt": work_dir / "uploads" / "current_ppt" / "current.pptx",
    }
    for role, path in files.items():
        path.parent.mkdir(parents=True, exist_ok=True)
        if role.endswith("_ppt"):
            continue
        _write_simple_workbook(path)
    _write_simple_deck(files["baseline_ppt"], ["Cover", "Old Summary"])
    _write_simple_deck(files["current_ppt"], ["Cover", "New Summary"])
    file_hashes = {role: sha256_file(path) for role, path in files.items()}
    choices = build_session_choices(
        mode=QCRunMode.CYCLE_COMPARISON,
        profile_name="default",
        files={role: str(path) for role, path in files.items()},
        file_hashes=file_hashes,
        output_mode="decision",
        allow_large_workbooks=False,
        allow_dependency_indexing=False,
        acceptance_absolute=0.0,
        acceptance_percent=0.0,
        rerun_of=None,
    )
    session_key = session_key_for(file_hashes)
    ConfigSessionStore(work_dir / "history.sqlite3").save_choices(
        session_key, profile_name="default", choices=choices
    )
    return session_key


@pytest.mark.asyncio
async def test_configure_page_reports_an_unknown_session(
    user: User, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    create_pages(work_dir)

    await user.open("/configure?session=does-not-exist")

    await user.should_see("This configuration session was not found")


@pytest.mark.asyncio
async def test_configure_page_lists_input_roles_and_completes_the_scan(
    user: User, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    session_key = _stage_session(work_dir)
    create_pages(work_dir)

    await user.open(f"/configure?session={session_key}")

    await user.should_see("baseline excel")
    await user.should_see("current excel")
    await user.should_see("Analysis complete", retries=_SUBPROCESS_RETRIES)

    registry = get_setup_job_registry(work_dir)
    job = registry.get_or_create(session_key)
    assert job.overall_status == "done"
    assert job.result is not None
    member = job.result.members["primary"]
    data_sheet = next(s for s in member.current_sheets if s.sheet_name == "Data")
    assert len(data_sheet.regions) == 1


@pytest.mark.asyncio
async def test_confirm_all_regions_persists_confirmed_true_for_a_valid_region(
    user: User, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    session_key = _stage_session(work_dir)
    create_pages(work_dir)

    await user.open(f"/configure?session={session_key}")
    await user.should_see("Analysis complete", retries=_SUBPROCESS_RETRIES)
    await user.should_see("Confirm all detected regions", retries=_SUBPROCESS_RETRIES)

    user.find("Confirm all detected regions").click()

    store = ConfigSessionStore(work_dir / "history.sqlite3")
    record = store.get(session_key)
    assert record is not None
    region_decisions = record.choices.get("region_decisions")
    assert isinstance(region_decisions, dict)
    assert region_decisions
    assert all(decision["confirmed"] is True for decision in region_decisions.values())


@pytest.mark.asyncio
async def test_selector_prerequisite_is_added_and_survives_a_session_reload(
    user: User, tmp_path: Path
) -> None:
    """Step 8: an analyst-declared selector prerequisite must round-trip
    through the session store, not just live in the running page's
    in-memory state -- otherwise resuming a saved configuration session
    (this store's whole purpose) would silently drop it.
    """
    work_dir = tmp_path / "work"
    session_key = _stage_session(work_dir)
    create_pages(work_dir)

    await user.open(f"/configure?session={session_key}")
    await user.should_see("Analysis complete", retries=_SUBPROCESS_RETRIES)
    await user.should_see("Add selector prerequisite:", retries=_SUBPROCESS_RETRIES)

    cell_input = next(
        element
        for element in user.find(kind=ui.input).elements
        if element.props.get("label") == "Cell (A1)"
    )
    label_input = next(
        element
        for element in user.find(kind=ui.input).elements
        if element.props.get("label") == "Label"
    )
    cell_input.value = "B5"
    label_input.value = "Scenario"
    user.find(kind=ui.button, content="Add").click()

    await user.should_see("Scenario (B5)")

    store = ConfigSessionStore(work_dir / "history.sqlite3")
    record = store.get(session_key)
    assert record is not None
    selectors = record.choices.get("selectors")
    assert isinstance(selectors, dict)
    saved_entries = selectors.get("primary")
    assert isinstance(saved_entries, list) and len(saved_entries) == 1
    assert saved_entries[0]["sheet_name"] == "Data"
    assert saved_entries[0]["label"] == "Scenario"
    assert saved_entries[0]["cell"] == "B5"

    # Simulate resuming the session later (a fresh page load re-runs the
    # scan from scratch): the persisted selector must be restored, not
    # silently dropped.
    await user.open(f"/configure?session={session_key}")
    await user.should_see("Analysis complete", retries=_SUBPROCESS_RETRIES)
    await user.should_see("Scenario (B5)")


@pytest.mark.asyncio
async def test_ppt_slide_review_shows_inventory_and_added_removed_warning(
    user: User, tmp_path: Path
) -> None:
    """Step 9: cycle mode's PowerPoint slide review shows the deck's slide
    titles and flags an unpaired added/removed slide as an unresolved
    (blocking) warning, exactly mirroring the Excel sheet-pairing
    criterion. Slides are peeked directly (no subprocess), so this
    exercises the section immediately after the Excel scan completes.
    """
    work_dir = tmp_path / "work"
    session_key = _stage_session_with_ppt(work_dir)
    create_pages(work_dir)

    await user.open(f"/configure?session={session_key}")
    await user.should_see("Analysis complete", retries=_SUBPROCESS_RETRIES)
    await user.should_see("PowerPoint slides")
    await user.should_see("Cover")
    await user.should_see("New Summary")
    # An unpaired added/removed slide is an unresolved blocker until the
    # analyst either declares a rename or acknowledges it.
    await user.should_see("Old Summary")


@pytest.mark.asyncio
async def test_ppt_slide_inclusion_toggle_persists_across_a_session_reload(
    user: User, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    session_key = _stage_session_with_ppt(work_dir)
    create_pages(work_dir)

    await user.open(f"/configure?session={session_key}")
    await user.should_see("Analysis complete", retries=_SUBPROCESS_RETRIES)
    await user.should_see("Cover")

    user.find(marker="slide-include-1").click()

    store = ConfigSessionStore(work_dir / "history.sqlite3")
    record = store.get(session_key)
    assert record is not None
    assert record.choices.get("excluded_slides") == [1]

    # Reopening the same session must restore the exclusion, not silently
    # re-include the slide.
    await user.open(f"/configure?session={session_key}")
    await user.should_see("Analysis complete", retries=_SUBPROCESS_RETRIES)
    await user.should_see("Cover", retries=_SUBPROCESS_RETRIES)
    excluded_checkbox = next(iter(user.find(marker="slide-include-1").elements))
    assert isinstance(excluded_checkbox, ui.checkbox)
    assert excluded_checkbox.value is False


@pytest.mark.asyncio
async def test_final_package_flags_a_stale_saved_slide_anchor_as_a_warning(
    user: User, tmp_path: Path
) -> None:
    """Step 9: final-package mode resolves a saved crosscheck mapping's
    slide title against the current deck. A title with zero matches is
    disclosed as a stale, warning-only mismatch -- never a blocker, and
    never guessed at (no other slide is substituted for it).
    """
    from qc_tool.config.profile import (
        CrosscheckMapping,
        CrosscheckProfile,
        DeliverableProfile,
        save_profile,
    )
    from qc_tool.config.profile import profile_path as _profile_path

    work_dir = tmp_path / "work"
    profiles_dir = work_dir / "profiles"
    stale_profile = DeliverableProfile(
        name="stale-mapping",
        crosscheck=CrosscheckProfile(
            mappings=[
                CrosscheckMapping(
                    slide="Old Summary",
                    line_skeleton="Revenue was #",
                    source_sheet="Data",
                    source_cell="B2",
                )
            ]
        ),
    )
    save_profile(stale_profile, _profile_path(profiles_dir, "stale-mapping"))

    files = {
        "current_excel": work_dir / "uploads" / "current_excel" / "current.xlsx",
        "current_ppt": work_dir / "uploads" / "current_ppt" / "current.pptx",
    }
    for path in files.values():
        path.parent.mkdir(parents=True, exist_ok=True)
    _write_simple_workbook(files["current_excel"])
    _write_simple_deck(files["current_ppt"], ["Cover", "New Summary"])
    file_hashes = {role: sha256_file(path) for role, path in files.items()}
    choices = build_session_choices(
        mode=QCRunMode.FINAL_PACKAGE,
        profile_name="stale-mapping",
        files={role: str(path) for role, path in files.items()},
        file_hashes=file_hashes,
        output_mode="decision",
        allow_large_workbooks=False,
        allow_dependency_indexing=False,
        acceptance_absolute=0.0,
        acceptance_percent=0.0,
        rerun_of=None,
    )
    session_key = session_key_for(file_hashes)
    ConfigSessionStore(work_dir / "history.sqlite3").save_choices(
        session_key, profile_name="stale-mapping", choices=choices
    )
    create_pages(work_dir)

    await user.open(f"/configure?session={session_key}")
    await user.should_see("Analysis complete", retries=_SUBPROCESS_RETRIES)
    await user.should_see("New Summary", retries=_SUBPROCESS_RETRIES)
    await user.should_see("Warnings")
    await user.should_see("Old Summary")
    await user.should_see("stale")


@pytest.mark.asyncio
async def test_save_profile_rejects_a_concurrent_edit_from_another_session(
    user: User, tmp_path: Path
) -> None:
    """plan-20260913 Step 11: "optimistic conflict protection" -- saving a
    profile this workspace session opened must refuse, not silently
    clobber, when the file on disk changed since it was opened (e.g. a
    second browser tab saved first).
    """
    from qc_tool.config.profile import DeliverableProfile, NumericTolerance, save_profile
    from qc_tool.config.profile import profile_path as _profile_path

    work_dir = tmp_path / "work"
    profiles_dir = work_dir / "profiles"
    profiles_dir.mkdir(parents=True)
    save_profile(
        DeliverableProfile(name="acme"), _profile_path(profiles_dir, "acme")
    )

    files = {
        "baseline_excel": work_dir / "uploads" / "baseline_excel" / "baseline.xlsx",
        "current_excel": work_dir / "uploads" / "current_excel" / "current.xlsx",
    }
    for path in files.values():
        path.parent.mkdir(parents=True, exist_ok=True)
        _write_simple_workbook(path)
    file_hashes = {role: sha256_file(path) for role, path in files.items()}
    choices = build_session_choices(
        mode=QCRunMode.CYCLE_COMPARISON,
        profile_name="acme",
        files={role: str(path) for role, path in files.items()},
        file_hashes=file_hashes,
        output_mode="decision",
        allow_large_workbooks=False,
        allow_dependency_indexing=False,
        acceptance_absolute=0.0,
        acceptance_percent=0.0,
        rerun_of=None,
    )
    session_key = session_key_for(file_hashes)
    ConfigSessionStore(work_dir / "history.sqlite3").save_choices(
        session_key, profile_name="acme", choices=choices
    )
    create_pages(work_dir)

    await user.open(f"/configure?session={session_key}")
    await user.should_see("Analysis complete", retries=_SUBPROCESS_RETRIES)
    await user.should_see("Confirm all detected regions", retries=_SUBPROCESS_RETRIES)
    user.find("Confirm all detected regions").click()

    # A "concurrent" edit: another session/tab saves the same profile name
    # with different content while this workspace stays open.
    save_profile(
        DeliverableProfile(
            name="acme", tolerance=NumericTolerance(absolute=99.0)
        ),
        _profile_path(profiles_dir, "acme"),
    )

    user.find(kind=ui.button, content="Save profile").click()

    await user.should_see("changed elsewhere while this workspace was open")
    # The concurrent edit's tolerance must survive untouched -- the click
    # above must not have overwritten it.
    from qc_tool.config.profile import load_profile_by_name

    reloaded = load_profile_by_name(profiles_dir, "acme")
    assert reloaded.tolerance.absolute == 99.0


@pytest.mark.asyncio
async def test_save_profile_succeeds_when_nothing_changed_underneath(
    user: User, tmp_path: Path
) -> None:
    """The inverse control: a save with no concurrent edit must still
    succeed normally -- conflict protection must not become a false
    positive on the ordinary, unchanged-file path.
    """
    from qc_tool.config.profile import DeliverableProfile, save_profile
    from qc_tool.config.profile import profile_path as _profile_path

    work_dir = tmp_path / "work"
    profiles_dir = work_dir / "profiles"
    profiles_dir.mkdir(parents=True)
    save_profile(
        DeliverableProfile(name="acme"), _profile_path(profiles_dir, "acme")
    )

    files = {
        "baseline_excel": work_dir / "uploads" / "baseline_excel" / "baseline.xlsx",
        "current_excel": work_dir / "uploads" / "current_excel" / "current.xlsx",
    }
    for path in files.values():
        path.parent.mkdir(parents=True, exist_ok=True)
        _write_simple_workbook(path)
    file_hashes = {role: sha256_file(path) for role, path in files.items()}
    choices = build_session_choices(
        mode=QCRunMode.CYCLE_COMPARISON,
        profile_name="acme",
        files={role: str(path) for role, path in files.items()},
        file_hashes=file_hashes,
        output_mode="decision",
        allow_large_workbooks=False,
        allow_dependency_indexing=False,
        acceptance_absolute=0.0,
        acceptance_percent=0.0,
        rerun_of=None,
    )
    session_key = session_key_for(file_hashes)
    ConfigSessionStore(work_dir / "history.sqlite3").save_choices(
        session_key, profile_name="acme", choices=choices
    )
    create_pages(work_dir)

    await user.open(f"/configure?session={session_key}")
    await user.should_see("Analysis complete", retries=_SUBPROCESS_RETRIES)
    await user.should_see("Confirm all detected regions", retries=_SUBPROCESS_RETRIES)
    user.find("Confirm all detected regions").click()

    user.find(kind=ui.button, content="Save profile").click()

    await user.should_see("Profile 'acme' saved")
