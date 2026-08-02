"""Context excerpts and re-QC feature tests."""

import shutil
import sqlite3
from pathlib import Path

import pytest
from nicegui import ui
from nicegui.testing import User

from qc_tool.engine import QCRunResult, compare_findings
from qc_tool.findings import FindingClass, Severity
from qc_tool.history.store import RunHistory
from qc_tool.ui.app import create_pages, perform_run
from tests.conftest import fixture_profile
from tests.fixtures.manifest_schema import FixtureManifest

pytest_plugins = ["nicegui.testing.user_plugin"]


# --- context excerpts ---------------------------------------------------------


def test_cell_findings_carry_excerpts(qc_result: QCRunResult, manifest: FixtureManifest) -> None:
    e01 = manifest.defect("E01")
    finding = next(
        f
        for f in qc_result.findings
        if f.finding_class is FindingClass.VALUE_CHANGED
        and (f.sheet, f.location) == ("Long_Monthly", e01.cell)
    )
    assert finding.current_excerpt is not None
    assert finding.baseline_excerpt is not None

    excerpt = finding.current_excerpt
    hit_value = excerpt.cells[excerpt.hit_row or 0][excerpt.hit_col or 0]
    assert float(hit_value) == float(e01.current or "")
    base_excerpt = finding.baseline_excerpt
    base_hit = base_excerpt.cells[base_excerpt.hit_row or 0][base_excerpt.hit_col or 0]
    assert float(base_hit) == float(e01.baseline or "")
    # Neighborhood includes the label columns for orientation.
    assert "South" in {cell for row in excerpt.cells for cell in row}


def test_formula_excerpt_shows_formula_text(qc_result: QCRunResult) -> None:
    finding = next(
        f
        for f in qc_result.findings
        if f.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
        and f.sheet == "Long_Monthly"
    )
    excerpt = finding.current_excerpt
    assert excerpt is not None
    hit = excerpt.cells[excerpt.hit_row or 0][excerpt.hit_col or 0]
    assert hit == "=C14-D14*1.1"  # uncached formula cells display their text


def test_growth_findings_have_no_excerpts(qc_result: QCRunResult) -> None:
    growth = [f for f in qc_result.findings if f.finding_class is FindingClass.ROW_GROWTH]
    assert growth
    assert all(f.current_excerpt is None and f.baseline_excerpt is None for f in growth)


def test_excerpts_roundtrip_through_history(
    qc_result: QCRunResult, tmp_path: Path
) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    run_id = history.record_run(qc_result, file_hashes={}, report_paths={})
    record = history.get_run(run_id)
    stored = next(f for f in record.findings if f.current_excerpt is not None)
    original = next(f for f in qc_result.findings if f.finding_id == stored.finding_id)
    assert stored.current_excerpt == original.current_excerpt


# --- history schema migration ---------------------------------------------------


def test_legacy_history_db_migrates(tmp_path: Path, qc_result: QCRunResult) -> None:
    db = tmp_path / "history.sqlite3"
    legacy_schema = """
    CREATE TABLE runs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        started_at TEXT NOT NULL, profile TEXT NOT NULL, files TEXT NOT NULL,
        file_hashes TEXT NOT NULL, counts TEXT NOT NULL, disclosures TEXT NOT NULL,
        verified_crosschecks INTEGER NOT NULL, findings TEXT NOT NULL,
        report_paths TEXT NOT NULL
    );
    """
    with sqlite3.connect(db) as conn:
        conn.executescript(legacy_schema)
        conn.execute(
            "INSERT INTO runs (started_at, profile, files, file_hashes, counts,"
            " disclosures, verified_crosschecks, findings, report_paths)"
            " VALUES ('2026-07-01T00:00:00+00:00', 'legacy', '{}', '{}', '{}',"
            " '[]', 0, '[]', '{}')"
        )

    history = RunHistory(db)  # must migrate, not fail
    legacy = history.get_run(1)
    assert legacy.file_paths == {} and legacy.rerun_of is None
    assert legacy.mode.value == "cycle_comparison" and legacy.coverage == []
    assert legacy.mapping_coverage is None and legacy.mapping_suggestions == []
    new_id = history.record_run(
        qc_result, file_hashes={}, report_paths={}, file_paths={"a": "/x"}, rerun_of=1
    )
    record = history.get_run(new_id)
    assert record.file_paths == {"a": "/x"} and record.rerun_of == 1


# --- re-QC -----------------------------------------------------------------------


def test_rerun_delta_after_fixing_files(fixture_dir: Path, tmp_path: Path) -> None:
    work_dir = tmp_path / "work"
    first = perform_run(
        work_dir,
        {
            "baseline_excel": fixture_dir / "baseline.xlsx",
            "current_excel": fixture_dir / "current.xlsx",
        },
        {},
        fixture_profile(),
    )
    first_open = sum(
        1 for f in first.result.findings if f.severity is not Severity.EXPECTED
    )
    assert first_open > 0

    # The analyst "fixes" the workbook and saves it under a new name (suffix).
    fixed = tmp_path / "current_v2_fixed.xlsx"
    shutil.copyfile(fixture_dir / "baseline.xlsx", fixed)

    second = perform_run(
        work_dir,
        {"baseline_excel": fixture_dir / "baseline.xlsx", "current_excel": fixed},
        {},
        fixture_profile(),
        rerun_of=first.run_id,
    )
    assert second.rerun_of == first.run_id
    assert second.delta is not None
    assert second.delta.resolved == first_open  # everything fixed
    assert second.delta.new == 0
    assert second.delta.persisting == 0

    record = RunHistory(work_dir / "history.sqlite3").get_run(second.run_id)
    assert record.rerun_of == first.run_id
    assert record.file_paths["current_excel"].endswith("current_v2_fixed.xlsx")


def test_rerun_same_files_all_persisting(fixture_dir: Path, tmp_path: Path) -> None:
    work_dir = tmp_path / "work"
    files = {
        "baseline_excel": fixture_dir / "baseline.xlsx",
        "current_excel": fixture_dir / "current.xlsx",
    }
    first = perform_run(work_dir, files, {}, fixture_profile())
    second = perform_run(work_dir, files, {}, fixture_profile(), rerun_of=first.run_id)
    assert second.delta is not None
    assert second.delta.resolved == 0 and second.delta.new == 0
    assert second.delta.persisting > 0


def test_compare_findings_ignores_expected(qc_result: QCRunResult) -> None:
    delta = compare_findings(qc_result.findings, qc_result.findings)
    non_expected = sum(
        1 for f in qc_result.findings if f.severity is not Severity.EXPECTED
    )
    assert delta.persisting == non_expected
    assert delta.resolved == 0 and delta.new == 0


@pytest.mark.asyncio
async def test_rerun_prefill_page(user: User, fixture_dir: Path, tmp_path: Path) -> None:
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
    await user.open(f"/?rerun={artifacts.run_id}")
    await user.should_see(f"Re-QC of run #{artifacts.run_id}")
    await user.should_see(f"verified, reused from run #{artifacts.run_id}")


@pytest.mark.asyncio
async def test_rerun_missing_file_shows_error(
    user: User, fixture_dir: Path, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    # Stage uploads inside the work dir (as the UI would), then remove one —
    # simulating a renamed/cleaned stored copy.
    staged = work_dir / "uploads" / "current_excel" / "current.xlsx"
    staged.parent.mkdir(parents=True)
    shutil.copyfile(fixture_dir / "current.xlsx", staged)
    artifacts = perform_run(
        work_dir,
        {"baseline_excel": fixture_dir / "baseline.xlsx", "current_excel": staged},
        {},
        fixture_profile(),
    )
    staged.unlink()  # the analyst renamed/removed the file

    create_pages(work_dir)
    await user.open(f"/?rerun={artifacts.run_id}")
    await user.should_see("file not found: current.xlsx — select the file again")
    await user.should_see(f"verified, reused from run #{artifacts.run_id}")  # baseline ok


@pytest.mark.asyncio
async def test_rerun_changed_file_shows_error(
    user: User, fixture_dir: Path, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    staged = work_dir / "uploads" / "current_excel" / "current.xlsx"
    staged.parent.mkdir(parents=True)
    shutil.copyfile(fixture_dir / "current.xlsx", staged)
    artifacts = perform_run(
        work_dir,
        {"baseline_excel": fixture_dir / "baseline.xlsx", "current_excel": staged},
        {},
        fixture_profile(),
    )
    # A later upload with the same name replaces the stored copy.
    shutil.copyfile(fixture_dir / "baseline.xlsx", staged)

    create_pages(work_dir)
    await user.open(f"/?rerun={artifacts.run_id}")
    await user.should_see(
        f"current.xlsx changed since run #{artifacts.run_id} — select the file again"
    )


@pytest.mark.asyncio
async def test_rerun_blocked_until_files_reselected(
    user: User, fixture_dir: Path, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    staged = work_dir / "uploads" / "current_excel" / "current.xlsx"
    staged.parent.mkdir(parents=True)
    shutil.copyfile(fixture_dir / "current.xlsx", staged)
    artifacts = perform_run(
        work_dir,
        {"baseline_excel": fixture_dir / "baseline.xlsx", "current_excel": staged},
        {},
        fixture_profile(),
    )
    staged.unlink()

    create_pages(work_dir)
    await user.open(f"/?rerun={artifacts.run_id}")
    # The readiness bar states the blocker before the analyst presses Run QC.
    await user.should_see(f"Re-QC of run #{artifacts.run_id} is blocked")
    run_button = user.find("Run QC").elements.pop()
    assert isinstance(run_button, ui.button)
    assert run_button.enabled is False
