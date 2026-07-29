"""End-to-end artifact, history, and read-only guarantees for new QC modes."""

import hashlib
from pathlib import Path

import pytest
from openpyxl import load_workbook

from qc_tool.config.profile import DeliverableProfile
from qc_tool.coverage import QCRunMode
from qc_tool.history.store import RunHistory
from qc_tool.ui.app import perform_run


@pytest.mark.parametrize(
    ("mode", "roles"),
    [
        (QCRunMode.CURRENT_FILE_PREFLIGHT, ("current_excel",)),
        (QCRunMode.CURRENT_FILE_PREFLIGHT, ("current_ppt",)),
        (QCRunMode.FINAL_PACKAGE, ("current_excel", "current_ppt")),
    ],
)
def test_new_modes_write_reports_history_and_preserve_sources(
    mode: QCRunMode,
    roles: tuple[str, ...],
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    source_by_role = {
        "current_excel": fixture_dir / "current.xlsx",
        "current_ppt": fixture_dir / "current.pptx",
    }
    files = {role: source_by_role[role] for role in roles}
    before = {
        role: hashlib.sha256(path.read_bytes()).hexdigest()
        for role, path in files.items()
    }

    artifacts = perform_run(
        tmp_path / "work",
        files,
        {},
        DeliverableProfile(name="mode-e2e"),
        mode=mode,
    )

    assert artifacts.result.mode is mode
    assert artifacts.result.coverage
    assert all(
        path.exists() and path.stat().st_size > 0
        for path in artifacts.report_paths.values()
    )
    workbook = load_workbook(artifacts.report_paths["excel"], read_only=True)
    assert "Coverage" in workbook.sheetnames
    html = artifacts.report_paths["html"].read_text(encoding="utf-8")
    assert mode.value in html and "Check coverage" in html
    record = RunHistory(tmp_path / "work" / "history.sqlite3").get_run(
        artifacts.run_id
    )
    assert record.mode is mode
    assert record.coverage == artifacts.result.coverage
    after = {
        role: hashlib.sha256(path.read_bytes()).hexdigest()
        for role, path in files.items()
    }
    assert after == before
