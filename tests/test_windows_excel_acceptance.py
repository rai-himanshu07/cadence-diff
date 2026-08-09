"""Contracts for the privacy-safe Windows Excel acceptance probe."""

import argparse
import json
import subprocess
from pathlib import Path

import pytest

import scripts.windows_excel_acceptance as acceptance
from qc_tool.io.xlsb_formula import XlsbFormulaScan


def _scan() -> XlsbFormulaScan:
    return XlsbFormulaScan(formula_cells={"Data": frozenset({(1, 1)})})


def test_acceptance_probe_refuses_non_windows_without_reading_source(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(acceptance.os, "name", "posix")
    args = argparse.Namespace(
        xlsb=tmp_path / "missing.xlsb",
        output=tmp_path / "result.json",
        timeout_seconds=2.0,
    )

    exit_code, payload = acceptance._run_acceptance(args)

    assert exit_code == 2
    assert payload == {
        "schema_version": 1,
        "overall": "unsupported_platform",
        "passed": False,
    }


def test_timeout_probe_requires_observed_owned_excel_exit(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        acceptance,
        "_prepare_worker",
        lambda work_dir, data, scan: {"status_path": str(work_dir / "status.json")},
    )
    monkeypatch.setattr(
        acceptance,
        "_run_forced_timeout",
        lambda work_dir, request, timeout: (
            "timeout",
            "owned_excel_started",
        ),
    )
    monkeypatch.setattr(acceptance, "_status_identity", lambda path: (1234, 10.0))
    monkeypatch.setattr(
        acceptance, "_wait_for_process_exit", lambda identity: "exited"
    )

    payload = acceptance._probe_timeout(b"xlsb", _scan(), timeout_seconds=2.0)

    assert payload["outcome"] == "timeout"
    assert payload["status_observed"] is True
    assert payload["timeout_setup"] == "owned_excel_started"
    assert payload["owned_excel_process"] == "exited"
    assert payload["passed"] is True


def test_forced_timeout_starts_owned_excel_before_production_timeout(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    events: list[str] = []

    class FakeProcess:
        pid = 1234
        returncode: int | None = None

        def communicate(
            self, input: str | None = None, timeout: float | None = None
        ) -> tuple[str, str]:
            if timeout is not None:
                assert input
                events.append("timeout_wait")
                raise subprocess.TimeoutExpired("formula-worker", timeout)
            events.append("post_kill_drain")
            return "", ""

        def kill(self) -> None:
            events.append("worker_killed")
            self.returncode = 1

        def poll(self) -> int | None:
            return self.returncode

    process = FakeProcess()

    def original_popen(
        command: list[str], *args: object, **kwargs: object
    ) -> FakeProcess:
        assert "--timeout-child" in command
        events.append("timeout_child_spawned")
        return process

    monkeypatch.setattr(
        acceptance.excel_formula_module.subprocess, "Popen", original_popen
    )
    monkeypatch.setattr(
        acceptance,
        "_status_identity",
        lambda path: (4321, 10.0),
    )
    monkeypatch.setattr(
        acceptance.excel_formula_module,
        "_terminate_owned_excel",
        lambda path: events.append("owned_excel_cleanup"),
    )

    outcome, setup = acceptance._run_forced_timeout(
        tmp_path,
        {
            "status_path": str(tmp_path / "status.json"),
            "result_path": str(tmp_path / "result.json"),
        },
        2.0,
    )

    assert outcome == "timeout"
    assert setup == "owned_excel_started"
    assert events == [
        "timeout_child_spawned",
        "timeout_wait",
        "owned_excel_cleanup",
        "worker_killed",
        "post_kill_drain",
            "owned_excel_cleanup",
    ]
    assert acceptance.excel_formula_module.subprocess.Popen is original_popen


def test_acceptance_payload_is_aggregate_only(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    source = tmp_path / "private-client-name.xlsb"
    source.write_bytes(b"private workbook bytes")
    args = argparse.Namespace(
        xlsb=source,
        output=tmp_path / "result.json",
        timeout_seconds=2.0,
    )
    monkeypatch.setattr(acceptance.os, "name", "nt")
    monkeypatch.setattr(acceptance, "scan_xlsb_formulas", lambda data: _scan())
    monkeypatch.setattr(acceptance, "_probe_dacl", lambda data: {"passed": True})
    monkeypatch.setattr(
        acceptance,
        "_probe_timeout",
        lambda data, scan, timeout_seconds: {"passed": True},
    )
    monkeypatch.setattr(
        acceptance,
        "_probe_parent_death",
        lambda data, scan: {"passed": True},
    )

    exit_code, payload = acceptance._run_acceptance(args)
    serialized = json.dumps(payload)

    assert exit_code == 0
    assert payload["overall"] == "passed"
    assert payload["source_unchanged"] is True
    assert source.name not in serialized
    assert str(source) not in serialized


def test_acceptance_can_run_only_timeout_and_localizes_probe_error(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    source = tmp_path / "source.xlsb"
    source.write_bytes(b"xlsb")
    args = argparse.Namespace(
        xlsb=source,
        output=tmp_path / "result.json",
        timeout_seconds=2.0,
        probe="timeout",
    )
    monkeypatch.setattr(acceptance.os, "name", "nt")
    monkeypatch.setattr(acceptance, "scan_xlsb_formulas", lambda data: _scan())
    monkeypatch.setattr(
        acceptance,
        "_probe_timeout",
        lambda data, scan, timeout_seconds: (_ for _ in ()).throw(
            RuntimeError("private diagnostic")
        ),
    )
    monkeypatch.setattr(
        acceptance,
        "_probe_dacl",
        lambda data: (_ for _ in ()).throw(AssertionError("must not run")),
    )
    monkeypatch.setattr(
        acceptance,
        "_probe_parent_death",
        lambda data, scan: (_ for _ in ()).throw(AssertionError("must not run")),
    )

    exit_code, payload = acceptance._run_acceptance(args)

    assert exit_code == 1
    assert payload["requested_probes"] == ["timeout"]
    assert payload["probes"] == {
        "timeout": {
            "error_type": "RuntimeError",
            "passed": False,
            "stage": "probe_exception",
        }
    }
    assert "private diagnostic" not in json.dumps(payload)
