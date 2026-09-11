"""Contracts for the privacy-safe XLSB load-locality diagnostic
(plan-20260908-phase-b-guest-performance-followup.md, Step 2)."""

import argparse
from pathlib import Path

import pytest

import scripts.xlsb_load_locality_diagnostic as diagnostic
from qc_tool.io.native_kernel import native_kernel_available
from tests.fixtures.xlsb_writer import StyledCell, write_xlsb


@pytest.fixture
def source_xlsb(tmp_path: Path) -> Path:
    path = tmp_path / "source.xlsb"
    write_xlsb(path, {"Data": [[StyledCell(1.0), StyledCell(2.0, is_formula=True)]]})
    return path


def test_stage_local_copy_is_byte_identical(source_xlsb: Path, tmp_path: Path) -> None:
    work_dir = tmp_path / "work"
    staged = diagnostic.stage_local_copy(source_xlsb, work_dir)

    assert staged.exists()
    assert staged.name != source_xlsb.name
    assert staged.read_bytes() == source_xlsb.read_bytes()


def test_run_probe_reports_only_safe_aggregate_fields(
    source_xlsb: Path, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    result = diagnostic.run_probe(
        source_xlsb,
        work_dir,
        xlsb_values_engine="pyxlsb",
        keep_staged_copy=False,
    )

    assert result["xlsb_values_engine"] == "pyxlsb"
    assert result["copy_matches_source"] is True
    assert result["source_hash_unchanged"] is True
    shared_seconds = result["shared_path_load_seconds"]
    local_seconds = result["local_copy_load_seconds"]
    assert isinstance(shared_seconds, float) and shared_seconds >= 0.0
    assert isinstance(local_seconds, float) and local_seconds >= 0.0
    assert result["speedup_ratio_shared_over_local"] is not None
    # Never the source path/filename in any reported field.
    assert all(
        source_xlsb.name not in str(value) for value in result.values()
    )


def test_run_probe_deletes_staged_copy_by_default(
    source_xlsb: Path, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    diagnostic.run_probe(
        source_xlsb, work_dir, xlsb_values_engine="pyxlsb", keep_staged_copy=False
    )

    assert list(work_dir.iterdir()) == []


def test_run_probe_keeps_staged_copy_when_requested(
    source_xlsb: Path, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    diagnostic.run_probe(
        source_xlsb, work_dir, xlsb_values_engine="pyxlsb", keep_staged_copy=True
    )

    assert len(list(work_dir.iterdir())) == 1


def test_run_probe_leaves_source_bytes_unchanged(
    source_xlsb: Path, tmp_path: Path
) -> None:
    before = source_xlsb.read_bytes()
    diagnostic.run_probe(
        source_xlsb,
        tmp_path / "work",
        xlsb_values_engine="pyxlsb",
        keep_staged_copy=False,
    )

    assert source_xlsb.read_bytes() == before


def test_safe_label_rejects_free_text() -> None:
    with pytest.raises(argparse.ArgumentTypeError):
        diagnostic._safe_label("has a space")
    with pytest.raises(argparse.ArgumentTypeError):
        diagnostic._safe_label("C:/some/path.xlsb")
    assert diagnostic._safe_label("large-workbook-current-shared-vs-local") == (
        "large-workbook-current-shared-vs-local"
    )


@pytest.mark.skipif(
    not native_kernel_available(),
    reason="native/xlsbkernel/ not built in this environment (optional accelerator)",
)
def test_run_probe_accepts_native_values_engine(
    source_xlsb: Path, tmp_path: Path
) -> None:
    result = diagnostic.run_probe(
        source_xlsb,
        tmp_path / "work",
        xlsb_values_engine="native",
        keep_staged_copy=False,
    )

    assert result["xlsb_values_engine"] == "native"
    assert result["copy_matches_source"] is True

