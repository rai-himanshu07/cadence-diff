"""End-to-end NiceGUI tests for the full-page mode-aware configuration
workspace (plan-20260913, Step 7): the ``/configure`` route, auto-start
scan, region rendering, and bulk confirmation, exercised through the real
``qc_tool.ui.app.create_pages()`` page registration and a real (tiny)
subprocess-based setup scan -- not a mock.
"""

from __future__ import annotations

from pathlib import Path

import pytest
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
