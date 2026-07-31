"""Headless `qc-tool run` tests: modes, exit codes, JSON export."""

import json
from pathlib import Path

import pytest

from qc_tool import cli


def _run(args: list[str]) -> int:
    return cli.main(["run", *args])


def test_cycle_run_exit_codes_and_json(
    fixture_dir: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    json_path = tmp_path / "findings.json"
    code = _run(
        [
            "--baseline-excel", str(fixture_dir / "baseline.xlsx"),
            "--current-excel", str(fixture_dir / "current.xlsx"),
            "--data-dir", str(tmp_path / "data"),
            "--json", str(json_path),
        ]
    )
    assert code == 2  # criticals present

    out = capsys.readouterr().out
    assert "mode: cycle_comparison" in out
    assert "review items:" in out
    assert "affected findings:" in out
    assert "critical=" in out and "reports:" in out

    payload = json.loads(json_path.read_text(encoding="utf-8"))
    assert payload["schema_version"] == 1
    assert payload["mode"] == "cycle_comparison"
    assert payload["counts"]["critical"] > 0
    assert any(f["finding_class"] == "value_changed" for f in payload["findings"])
    assert payload["context_included"] is False
    assert all("current_excerpt" not in finding for finding in payload["findings"])


def test_fail_on_never_returns_zero(fixture_dir: Path, tmp_path: Path) -> None:
    code = _run(
        [
            "--baseline-excel", str(fixture_dir / "baseline.xlsx"),
            "--current-excel", str(fixture_dir / "current.xlsx"),
            "--data-dir", str(tmp_path / "data"),
            "--fail-on", "never",
        ]
    )
    assert code == 0


def test_preflight_mode_inferred(
    fixture_dir: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    code = _run(
        [
            "--current-excel", str(fixture_dir / "current.xlsx"),
            "--data-dir", str(tmp_path / "data"),
        ]
    )
    out = capsys.readouterr().out
    assert "mode: current_file_preflight" in out
    assert "coverage:" in out
    assert code == 2  # the fixture contains seeded error literals


def test_progress_is_opt_in_and_stderr_only(
    fixture_dir: Path,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    code = _run(
        [
            "--current-excel",
            str(fixture_dir / "current.xlsx"),
            "--data-dir",
            str(tmp_path / "data"),
            "--progress",
            "--fail-on",
            "never",
        ]
    )

    captured = capsys.readouterr()
    assert code == 0
    assert "progress: loading current excel" in captured.err
    assert "progress: recording history (1/1)" in captured.err
    assert "progress:" not in captured.out
    assert "mode: current_file_preflight" in captured.out


def test_package_mode_inferred(
    fixture_dir: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    code = _run(
        [
            "--current-excel", str(fixture_dir / "current.xlsx"),
            "--current-ppt", str(fixture_dir / "current.pptx"),
            "--data-dir", str(tmp_path / "data"),
        ]
    )
    out = capsys.readouterr().out
    assert "mode: final_package" in out
    assert "mapping: eligible=" in out
    assert code in (0, 2)


def test_password_role_syntax(fixture_dir: Path, tmp_path: Path, manifest) -> None:
    code = _run(
        [
            "--baseline-excel", str(fixture_dir / "baseline.xlsx"),
            "--current-excel", str(fixture_dir / "current_encrypted.xlsx"),
            "--data-dir", str(tmp_path / "data"),
            "--password", f"current_excel={manifest.password}",
            "--fail-on", "never",
        ]
    )
    assert code == 0


def test_run_records_history(fixture_dir: Path, tmp_path: Path) -> None:
    from qc_tool.history.store import RunHistory

    data_dir = tmp_path / "data"
    _run(
        [
            "--current-excel", str(fixture_dir / "current.xlsx"),
            "--data-dir", str(data_dir),
            "--fail-on", "never",
        ]
    )
    runs = RunHistory(data_dir / "history.sqlite3").list_runs()
    assert len(runs) == 1
    assert runs[0].files["current_excel"] == "current.xlsx"


def test_invalid_mode_combination_returns_clean_error(
    fixture_dir: Path, tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    code = _run(
        [
            "--mode",
            "preflight",
            "--baseline-excel",
            str(fixture_dir / "baseline.xlsx"),
            "--current-excel",
            str(fixture_dir / "current.xlsx"),
            "--data-dir",
            str(tmp_path / "data"),
        ]
    )
    assert code == 1
    captured = capsys.readouterr()
    assert "current-file preflight does not accept baseline files" in captured.err
    assert "Traceback" not in captured.err


def test_password_can_be_read_from_environment(
    fixture_dir: Path,
    tmp_path: Path,
    manifest,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("QC_CURRENT_PASSWORD", manifest.password)
    code = _run(
        [
            "--baseline-excel",
            str(fixture_dir / "baseline.xlsx"),
            "--current-excel",
            str(fixture_dir / "current_encrypted.xlsx"),
            "--data-dir",
            str(tmp_path / "data"),
            "--password-env",
            "current_excel=QC_CURRENT_PASSWORD",
            "--fail-on",
            "never",
        ]
    )
    assert code == 0
