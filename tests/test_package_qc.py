"""Current Excel-to-PowerPoint final-package QC."""

from pathlib import Path

from qc_tool.coverage import CoverageState, QCRunMode
from qc_tool.engine import run_qc
from qc_tool.findings import FindingClass
from qc_tool.history.store import RunHistory
from tests.conftest import fixture_profile


def test_final_package_reports_mapping_coverage_and_suggestions(
    fixture_dir: Path, tmp_path: Path
) -> None:
    result = run_qc(
        current_excel=fixture_dir / "current.xlsx",
        current_ppt=fixture_dir / "current.pptx",
        profile=fixture_profile(),
        mode=QCRunMode.FINAL_PACKAGE,
    )

    assert result.mode is QCRunMode.FINAL_PACKAGE
    assert result.mapping_coverage is not None
    mapping = result.mapping_coverage
    assert mapping.eligible > mapping.mapped == 2
    assert mapping.verified == 1
    assert mapping.mismatched == 1
    assert mapping.unmapped == mapping.eligible - mapping.mapped
    assert result.mapping_suggestions
    assert any(suggestion.candidates for suggestion in result.mapping_suggestions)
    assert any(
        finding.finding_class is FindingClass.CROSSCHECK_MISMATCH
        for finding in result.findings
    )
    package_coverage = next(
        item for item in result.coverage if item.check_id == "excel-ppt-crosscheck"
    )
    assert package_coverage.state is CoverageState.CHECKED

    history = RunHistory(tmp_path / "history.sqlite3")
    run_id = history.record_run(result, file_hashes={}, report_paths={})
    stored = history.get_run(run_id)
    assert stored.mapping_coverage == result.mapping_coverage
    assert stored.mapping_suggestions == result.mapping_suggestions

    result.mapping_coverage.mapped += 1
    result.mapping_coverage.verified += 1
    result.mapping_coverage.unmapped -= 1
    result.mapping_suggestions = result.mapping_suggestions[1:]
    package_coverage.detail = "updated mapping coverage"
    history.update_mapping_review(
        run_id,
        coverage=result.mapping_coverage,
        suggestions=result.mapping_suggestions,
        check_coverage=result.coverage,
    )
    updated = history.get_run(run_id)
    assert updated.mapping_coverage == result.mapping_coverage
    assert updated.mapping_suggestions == result.mapping_suggestions
    assert next(
        item for item in updated.coverage if item.check_id == "excel-ppt-crosscheck"
    ).detail == "updated mapping coverage"
