"""Disposable population-excerpt worker: safe synthetic end-to-end tests."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from qc_tool.excel.population_excerpt_worker import (
    PopulationExcerptRequest,
    run_population_excerpt_worker,
)
from qc_tool.history.store import sha256_file


def _write_workbook(path: Path, *, value: str) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet["A1"] = "label"
    sheet["B2"] = value
    workbook.save(path)


def test_worker_builds_excerpts_for_requested_samples(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path, value="old")
    _write_workbook(current_path, value="new")

    request = PopulationExcerptRequest(
        baseline_path=str(baseline_path),
        current_path=str(current_path),
        baseline_hash=sha256_file(baseline_path),
        current_hash=sha256_file(current_path),
        sheet="Data",
        samples=(("B2", "B2"),),
    )

    result = run_population_excerpt_worker(request)

    assert result.disclosure == ""
    assert set(result.excerpts) == {"B2"}
    baseline_excerpt, current_excerpt = result.excerpts["B2"]
    assert baseline_excerpt is not None and current_excerpt is not None
    assert baseline_excerpt.hit_row is not None and baseline_excerpt.hit_col is not None
    assert any("old" in row for row in baseline_excerpt.cells)
    assert any("new" in row for row in current_excerpt.cells)


def test_worker_discloses_a_hash_mismatch_without_touching_the_source_content(
    tmp_path: Path,
) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path, value="old")
    _write_workbook(current_path, value="new")

    request = PopulationExcerptRequest(
        baseline_path=str(baseline_path),
        current_path=str(current_path),
        baseline_hash="0" * 64,  # deliberately wrong
        current_hash=sha256_file(current_path),
        sheet="Data",
        samples=(("B2", "B2"),),
    )

    result = run_population_excerpt_worker(request)

    assert result.excerpts == {}
    assert "changed since this run recorded it" in result.disclosure


def test_worker_discloses_a_missing_source_file(tmp_path: Path) -> None:
    current_path = tmp_path / "current.xlsx"
    _write_workbook(current_path, value="new")

    request = PopulationExcerptRequest(
        baseline_path=str(tmp_path / "does-not-exist.xlsx"),
        current_path=str(current_path),
        baseline_hash="",
        current_hash=sha256_file(current_path),
        sheet="Data",
        samples=(("B2", "B2"),),
    )

    result = run_population_excerpt_worker(request)

    assert result.excerpts == {}
    assert "no longer at its recorded location" in result.disclosure


def test_worker_discloses_a_missing_sheet(tmp_path: Path) -> None:
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path, value="old")
    _write_workbook(current_path, value="new")

    request = PopulationExcerptRequest(
        baseline_path=str(baseline_path),
        current_path=str(current_path),
        baseline_hash=sha256_file(baseline_path),
        current_hash=sha256_file(current_path),
        sheet="NoSuchSheet",
        samples=(("B2", "B2"),),
    )

    result = run_population_excerpt_worker(request)

    assert result.excerpts == {}
    assert "was not found in the reopened sources" in result.disclosure


def test_worker_times_out_gracefully_instead_of_hanging(tmp_path: Path) -> None:
    """A near-zero timeout must degrade to a plain disclosure -- never raise,
    hang, or leave a zombie process -- regardless of how fast the real child
    could have finished.
    """
    baseline_path = tmp_path / "baseline.xlsx"
    current_path = tmp_path / "current.xlsx"
    _write_workbook(baseline_path, value="old")
    _write_workbook(current_path, value="new")

    request = PopulationExcerptRequest(
        baseline_path=str(baseline_path),
        current_path=str(current_path),
        baseline_hash=sha256_file(baseline_path),
        current_hash=sha256_file(current_path),
        sheet="Data",
        samples=(("B2", "B2"),),
    )

    result = run_population_excerpt_worker(request, timeout_seconds=0.0001)

    assert result.excerpts == {}
    assert result.disclosure == "excerpt worker timed out"
