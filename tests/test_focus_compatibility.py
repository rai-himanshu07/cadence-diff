"""Step 8: cumulative compatibility of desktop focus with existing evidence.

Focus is additive. A run recorded with a target sidecar must still produce the
same atomic identity, default JSON, reports, attestation, Re-QC delta, review
groups, and annotations as one recorded without it, and no focus action may be
rendered on a platform or network mode where the feature refuses.
"""

import json
from pathlib import Path

import pytest
from nicegui.testing import User

from qc_tool.attestation import create_attestation, verify_attestation
from qc_tool.engine import compare_findings
from qc_tool.focus.model import FOCUS_SIDECAR_VERSION, FocusTargetSidecar
from qc_tool.focus.service import FocusService, FocusUnavailable
from qc_tool.history.store import RunHistory, export_runs_archive
from qc_tool.report.excel_report import write_excel_report
from qc_tool.report.html_report import write_html_report
from qc_tool.review import build_pattern_groups, build_review_groups
from qc_tool.run_service import perform_run
from qc_tool.server_config import NetworkMode
from qc_tool.ui.app import create_pages
from tests.conftest import fixture_profile

pytest_plugins = ["nicegui.testing.user_plugin"]


@pytest.fixture(scope="module")
def cycle_run(fixture_dir: Path, tmp_path_factory: pytest.TempPathFactory):
    work_dir = tmp_path_factory.mktemp("focus-compat") / "work"
    artifacts = perform_run(
        work_dir,
        {
            "baseline_excel": fixture_dir / "baseline.xlsx",
            "current_excel": fixture_dir / "current.xlsx",
            "baseline_ppt": fixture_dir / "baseline.pptx",
            "current_ppt": fixture_dir / "current.pptx",
        },
        {},
        fixture_profile(),
    )
    return work_dir, artifacts


def test_a_real_run_records_a_versioned_sidecar(cycle_run) -> None:
    work_dir, artifacts = cycle_run
    record = RunHistory(work_dir / "history.sqlite3").get_run(artifacts.run_id)
    assert record.focus_targets.version == FOCUS_SIDECAR_VERSION
    assert record.focus_targets.usable
    identifiers = {finding.finding_id for finding in record.findings}
    assert set(record.focus_targets.targets) <= identifiers


def test_targets_never_reference_an_omitted_finding(cycle_run) -> None:
    work_dir, artifacts = cycle_run
    record = RunHistory(work_dir / "history.sqlite3").get_run(artifacts.run_id)
    for finding_id in record.focus_targets.targets:
        assert any(f.finding_id == finding_id for f in record.findings)


def test_default_json_is_unchanged_by_focus(cycle_run) -> None:
    _work_dir, artifacts = cycle_run
    payload = [
        finding.model_dump(mode="json") for finding in artifacts.result.findings
    ]
    text = json.dumps(payload)
    assert "focus" not in text
    assert "focus_targets" not in text
    for key in ("address", "path_salt", "expected_path_digest"):
        assert key not in text


def test_identity_review_groups_and_reqc_are_unchanged(cycle_run) -> None:
    work_dir, artifacts = cycle_run
    history = RunHistory(work_dir / "history.sqlite3")
    record = history.get_run(artifacts.run_id)
    assert [f.finding_id for f in record.findings] == [
        f.finding_id for f in artifacts.result.findings
    ]
    assert len(build_review_groups(record.findings)) == len(
        build_review_groups(artifacts.result.findings)
    )
    assert len(build_pattern_groups(record.findings)) == len(
        build_pattern_groups(artifacts.result.findings)
    )
    delta = compare_findings(record.findings, artifacts.result.findings)
    assert (delta.resolved, delta.new) == (0, 0)


def test_reports_regenerate_from_a_stored_record(cycle_run, tmp_path: Path) -> None:
    work_dir, artifacts = cycle_run
    record = RunHistory(work_dir / "history.sqlite3").get_run(artifacts.run_id)
    from qc_tool.ui.app import _result_from_record

    result = _result_from_record(record)
    excel_path = tmp_path / "report.xlsx"
    html_path = tmp_path / "report.html"
    write_excel_report(result, excel_path)
    write_html_report(result, html_path)
    assert excel_path.stat().st_size > 0
    html = html_path.read_text(encoding="utf-8")
    assert "focus_targets" not in html


def test_attestation_still_verifies(cycle_run, fixture_dir: Path, tmp_path: Path) -> None:
    _work_dir, artifacts = cycle_run
    destination = tmp_path / "run.qca"
    key = b"\x01" * 32
    create_attestation(
        destination,
        result=artifacts.result,
        profile=fixture_profile(),
        input_files={
            "baseline_excel": fixture_dir / "baseline.xlsx",
            "current_excel": fixture_dir / "current.xlsx",
        },
        report_paths=artifacts.report_paths,
        key=key,
    )
    verification = verify_attestation(destination, key=key)
    assert verification.valid


def test_annotations_survive_a_recorded_sidecar(cycle_run) -> None:
    work_dir, artifacts = cycle_run
    history = RunHistory(work_dir / "history.sqlite3")
    record = history.get_run(artifacts.run_id)
    target = record.findings[0].finding_id
    history.set_annotation(
        artifacts.run_id, target, severity="info", comment="reviewed"
    )
    reloaded = history.get_run(artifacts.run_id)
    annotated = next(f for f in reloaded.findings if f.finding_id == target)
    assert annotated.analyst_comment == "reviewed"
    assert annotated.severity_overridden
    assert reloaded.focus_targets.usable


def test_exported_history_carries_no_focus_target(cycle_run, tmp_path: Path) -> None:
    work_dir, artifacts = cycle_run
    history = RunHistory(work_dir / "history.sqlite3")
    record = history.get_run(artifacts.run_id)
    destination = tmp_path / "runs.zip"
    export_runs_archive([record], destination, managed_root=work_dir / "runs")
    payload = destination.read_bytes()
    assert b"focus_targets" not in payload
    assert b"expected_path_digest" not in payload


def test_a_legacy_run_row_still_loads(cycle_run) -> None:
    import sqlite3

    work_dir, artifacts = cycle_run
    db_path = work_dir / "history.sqlite3"
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE runs SET focus_targets = '{}' WHERE id = ?", (artifacts.run_id,)
        )
    record = RunHistory(db_path).get_run(artifacts.run_id)
    assert record.focus_targets == FocusTargetSidecar()
    assert record.findings
    # Restore the sidecar so module-scoped siblings keep their fixture.
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE runs SET focus_targets = ? WHERE id = ?",
            (
                json.dumps(
                    {
                        "version": FOCUS_SIDECAR_VERSION,
                        "targets": {
                            finding_id: [seed.model_dump(mode="json") for seed in seeds]
                            for finding_id, seeds in (
                                RunHistory(db_path)
                                .get_run(artifacts.run_id)
                                .focus_targets.targets.items()
                            )
                        },
                    }
                ),
                artifacts.run_id,
            ),
        )


def test_generation_failure_leaves_the_run_successful(
    fixture_dir: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import qc_tool.run_service as run_service

    def explode(*_args: object, **_kwargs: object) -> dict[str, object]:
        raise RuntimeError("target generation failed")

    monkeypatch.setattr(run_service, "build_focus_targets", explode)
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
    assert record.findings
    assert not record.focus_targets.usable
    assert artifacts.report_paths["html"].exists()


# --------------------------------------------------------------------------
# rendered guards
# --------------------------------------------------------------------------


def test_service_refuses_on_every_disallowed_configuration(tmp_path: Path) -> None:
    combinations = [
        (False, "win32", NetworkMode.LOCAL, FocusUnavailable.FEATURE_OFF),
        (True, "linux", NetworkMode.LOCAL, FocusUnavailable.UNSUPPORTED_PLATFORM),
        (True, "darwin", NetworkMode.LOCAL, FocusUnavailable.UNSUPPORTED_PLATFORM),
        (True, "win32", NetworkMode.LAN, FocusUnavailable.NETWORK_NOT_LOOPBACK),
        (False, "linux", NetworkMode.LAN, FocusUnavailable.FEATURE_OFF),
    ]
    for enabled, platform, mode, expected in combinations:
        service = FocusService(
            tmp_path, enabled=enabled, platform=platform, network_mode=mode
        )
        assert service.availability() is expected
        assert not service.available


@pytest.mark.asyncio
async def test_no_focus_action_is_rendered_with_the_feature_off(
    user: User, cycle_run
) -> None:
    work_dir, artifacts = cycle_run
    create_pages(work_dir)
    await user.open(f"/runs/{artifacts.run_id}")
    await user.should_see(f"Run #{artifacts.run_id}")
    await user.should_not_see("Open in desktop Office")
    await user.should_not_see("Bind current Excel")


@pytest.mark.asyncio
async def test_no_focus_action_is_rendered_on_this_platform_when_enabled(
    user: User, cycle_run
) -> None:
    work_dir, artifacts = cycle_run
    create_pages(work_dir, desktop_focus=True)
    await user.open(f"/runs/{artifacts.run_id}")
    await user.should_see(f"Run #{artifacts.run_id}")
    await user.should_not_see("Open in desktop Office")
