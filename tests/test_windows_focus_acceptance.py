"""Contracts for the Step 9 Windows focus acceptance probe.

The probe itself is untracked and only runs in the guest. These tests pin its
argument surface, its privacy rules, and its non-Windows refusal so a defect is
found here rather than during a manual acceptance session.
"""

import json
import os
from pathlib import Path

import pytest

from scripts import windows_focus_acceptance as probe

pytestmark = pytest.mark.skipif(
    os.name == "nt", reason="these contracts assert the non-Windows refusal path"
)


def test_scenarios_cover_every_required_acceptance_case() -> None:
    required = {
        "excel-current",
        "excel-baseline",
        "ppt-current",
        "ppt-baseline",
        "duplicate-hash",
        "managed-copy",
        "dirty-matching",
        "dirty-changed",
        "autosave-on",
        "two-excel-processes",
        "multi-window",
        "protected-view",
        "active-content",
        "helper-timeout",
        "foreground-denial",
        "latency",
    }
    assert required <= set(probe.SCENARIOS)


def test_every_role_maps_to_an_application() -> None:
    assert set(probe._ROLES) == {
        "current_excel",
        "baseline_excel",
        "current_ppt",
        "baseline_ppt",
    }
    assert set(probe._APPLICATIONS) == set(probe._ROLES.values())


def test_parser_requires_a_declared_scenario_role_source_and_expectation() -> None:
    parser = probe.build_parser()
    with pytest.raises(SystemExit):
        parser.parse_args(["--scenario", "excel-current"])
    namespace = parser.parse_args(
        [
            "--scenario",
            "excel-current",
            "--role",
            "current_excel",
            "--source",
            "C:/pilot/current.xlsx",
            "--expect",
            "focused",
            "--output",
            "out.json",
        ]
    )
    assert namespace.repeats == 1
    assert namespace.timeout == 25.0
    assert namespace.managed_root is None
    assert namespace.shape_id is None


def test_probe_refuses_off_windows_and_writes_a_fixed_code(tmp_path: Path) -> None:
    output = tmp_path / "result.json"
    code = probe.main(
        [
            "--scenario",
            "excel-current",
            "--role",
            "current_excel",
            "--source",
            str(tmp_path / "missing.xlsx"),
            "--expect",
            "focused",
            "--output",
            str(output),
        ]
    )
    assert code == 2
    payload = json.loads(output.read_text(encoding="utf-8"))
    assert payload["verdict"] == "unsupported_platform"
    assert payload["schema_version"] == probe._SCHEMA_VERSION


def test_observation_payload_carries_no_free_text_field() -> None:
    observation = probe.Observation(
        scenario="excel-current", role="current_excel", expected="focused"
    )
    payload = observation.payload()
    forbidden = {"path", "full_name", "sheet", "address", "sha256", "message"}
    assert not forbidden & set(payload)
    assert payload["verdict"] == ""
    assert payload["source_unchanged"] is True


def test_a_wrong_outcome_and_a_mutated_source_are_both_fatal() -> None:
    assert probe._VERDICT_FAIL in probe._FATAL_VERDICTS
    assert probe._VERDICT_MUTATED in probe._FATAL_VERDICTS
    assert probe._VERDICT_PASS not in probe._FATAL_VERDICTS


def test_verdict_requires_unchanged_source_and_process_count() -> None:
    observation = probe.Observation(
        scenario="excel-current",
        role="current_excel",
        expected="focused",
        processes_before=1,
        processes_after=1,
    )
    probe._set_verdict(observation, "focused")
    assert observation.verdict == probe._VERDICT_PASS

    observation.processes_after = 2
    probe._set_verdict(observation, "focused")
    assert observation.verdict == probe._VERDICT_MUTATED

    observation.processes_after = 1
    observation.source_unchanged = False
    probe._set_verdict(observation, "focused")
    assert observation.verdict == probe._VERDICT_MUTATED


def test_permission_denied_hash_uses_windows_shared_reader(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    path = tmp_path / "locked.xlsx"
    path.write_bytes(b"locked")
    monkeypatch.setattr(probe.sys, "platform", "win32")
    monkeypatch.setattr(
        Path,
        "open",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(PermissionError()),
    )
    monkeypatch.setattr(probe, "_sha256_windows_shared", lambda _path: "shared")
    assert probe._sha256(path) == "shared"


def test_duplicate_analyst_documents_are_counted_per_process(tmp_path: Path) -> None:
    from qc_tool.focus.discovery import FocusApplication, OpenDocument

    def document(process_id: int, name: str) -> OpenDocument:
        return OpenDocument(
            application=FocusApplication.EXCEL,
            process_id=process_id,
            process_created=1.0,
            windows_session_id=1,
            full_name=name,
            window_count=1,
            visible_window_count=1,
            visible_window_handles=(100 + process_id,),
            saved=True,
            autosave=False,
        )

    counts = probe._counts(
        [
            document(1, "C:/a.xlsx"),
            document(2, "C:/a.xlsx"),
            document(3, "C:/b.xlsx"),
        ]
    )
    assert counts["analyst_candidates"] == 3
    assert counts["duplicate_path_analyst_documents"] == 2
