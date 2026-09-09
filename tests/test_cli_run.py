"""Headless `qc-tool run` tests: modes, exit codes, JSON export."""

import json
import zipfile
from pathlib import Path

import pytest

from qc_tool import cli
from qc_tool.history.store import RunHistory


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
    assert "pattern review items:" in out
    assert "spatial review items:" in out
    assert "atomic findings:" in out
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


def test_blocked_prerequisite_returns_exit_code_three(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    from openpyxl import Workbook

    from qc_tool.config.profile import (
        ComparisonPrerequisite,
        DeliverableProfile,
        ExcelProfile,
        save_profile,
    )

    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    for path, scenario in ((baseline, "Base Case"), (current, "Upside Case")):
        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "Config"
        sheet["B2"] = scenario
        workbook.save(path)

    profile = DeliverableProfile(
        name="prereq",
        excel=ExcelProfile(
            comparison_prerequisites=[
                ComparisonPrerequisite(name="Scenario", sheet="Config", cell="B2")
            ]
        ),
    )
    profile_path = tmp_path / "prereq.yaml"
    save_profile(profile, profile_path)
    data_dir = tmp_path / "data"

    code = _run(
        [
            "--baseline-excel", str(baseline),
            "--current-excel", str(current),
            "--data-dir", str(data_dir),
            "--profile", str(profile_path),
        ]
    )

    assert code == 3
    err = capsys.readouterr().err
    assert "blocked:" in err
    assert "Config!B2" in err
    assert "Base Case" not in err
    assert "Upside Case" not in err
    history_db = data_dir / "history.sqlite3"
    assert not history_db.exists() or RunHistory(history_db).list_runs() == []


def test_population_summary_line_printed_when_policy_enabled(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    from openpyxl import Workbook

    from qc_tool.config.profile import (
        DeliverableProfile,
        PopulationPolicy,
        ReviewPolicy,
        save_profile,
    )

    def build(path: Path, *, multiplier: int) -> None:
        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "Data"
        sheet.append(["Input", "Output"])
        for row in range(2, 21):  # 19 uniform formula changes
            sheet.append([100, f"=A{row}*{multiplier}"])
        workbook.save(path)

    baseline = tmp_path / "base.xlsx"
    current = tmp_path / "curr.xlsx"
    build(baseline, multiplier=2)
    build(current, multiplier=3)
    profile = DeliverableProfile(
        name="population-cli",
        review_policy=ReviewPolicy(
            populations=PopulationPolicy(enabled=True, threshold=10)
        ),
    )
    profile_path = tmp_path / "population.yaml"
    save_profile(profile, profile_path)

    code = _run(
        [
            "--baseline-excel", str(baseline),
            "--current-excel", str(current),
            "--data-dir", str(tmp_path / "data"),
            "--profile", str(profile_path),
        ]
    )

    out = capsys.readouterr().out
    assert code in (0, 1, 2)
    assert "populations (formula_logic_changed):" in out
    assert "19 cells summarised as 1 population" in out


def test_repeatable_current_workbook_flag_records_member_manifest_and_v2_json(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    data_dir = tmp_path / "data"
    output = tmp_path / "package.json"

    code = _run(
        [
            "--current-workbook",
            f"ops={fixture_dir / 'current.xlsx'}",
            "--data-dir",
            str(data_dir),
            "--json",
            str(output),
            "--fail-on",
            "never",
        ]
    )

    record = RunHistory(data_dir / "history.sqlite3").get_run(1)
    payload = json.loads(output.read_text(encoding="utf-8"))
    assert code == 0
    assert record.files == {"current_excel:ops": "current.xlsx"}
    assert record.package_manifest is not None
    assert record.package_manifest.role_keys == ("current_excel:ops",)
    assert payload["schema_version"] == 2
    assert payload["package_manifest"]["members"][0]["member_id"] == "ops"


def test_duplicate_or_reserved_cli_member_id_fails_before_recording(
    fixture_dir: Path,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    data_dir = tmp_path / "data"
    source = fixture_dir / "current.xlsx"

    duplicate = _run(
        [
            "--current-workbook",
            f"ops={source}",
            "--current-workbook",
            f"ops={source}",
            "--data-dir",
            str(data_dir),
        ]
    )
    duplicate_error = capsys.readouterr().err
    reserved = _run(
        [
            "--current-workbook",
            f"primary={source}",
            "--data-dir",
            str(data_dir),
        ]
    )
    reserved_error = capsys.readouterr().err

    assert duplicate == 1
    assert "duplicate member role current_excel:ops" in duplicate_error
    assert reserved == 1
    assert "member id 'primary' uses --baseline-excel/--current-excel" in (
        reserved_error
    )
    assert not (data_dir / "history.sqlite3").exists()


def test_byte_identical_cycle_refuses_before_reports_or_history(
    fixture_dir: Path,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    source = fixture_dir / "current.xlsx"
    data_dir = tmp_path / "data"

    code = _run(
        [
            "--baseline-excel",
            str(source),
            "--current-excel",
            str(source),
            "--data-dir",
            str(data_dir),
            "--fail-on",
            "never",
        ]
    )

    assert code == 1
    assert "byte-identical" in capsys.readouterr().err
    assert not (data_dir / "history.sqlite3").exists()
    assert not list((data_dir / "runs").glob("*/qc_report.*"))


def test_same_side_duplicate_member_bytes_refuse_before_history(
    fixture_dir: Path,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    source = fixture_dir / "current.xlsx"
    data_dir = tmp_path / "data"

    code = _run(
        [
            "--current-workbook",
            f"core={source}",
            "--current-workbook",
            f"ops={source}",
            "--data-dir",
            str(data_dir),
            "--fail-on",
            "never",
        ]
    )

    assert code == 1
    assert "duplicate bytes" in capsys.readouterr().err.casefold()
    assert not (data_dir / "history.sqlite3").exists()


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


@pytest.mark.parametrize(
    ("flag", "filename"),
    [
        ("--current-excel", "corrupt.xlsx"),
        ("--current-ppt", "corrupt.pptx"),
    ],
)
def test_corrupt_office_package_returns_bounded_error_without_history(
    flag: str,
    filename: str,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    source = tmp_path / filename
    source.write_bytes(b"PK\x03\x04not-a-valid-zip")
    data_dir = tmp_path / "data"

    code = _run(
        [
            flag,
            str(source),
            "--data-dir",
            str(data_dir),
            "--fail-on",
            "never",
        ]
    )

    captured = capsys.readouterr()
    assert code == 1
    assert filename in captured.err
    assert "invalid Office package" in captured.err
    assert "Traceback" not in captured.err
    assert str(tmp_path) not in captured.err
    assert not (data_dir / "history.sqlite3").exists()


@pytest.mark.parametrize(
    ("flag", "filename"),
    [
        ("--current-excel", "empty.xlsx"),
        ("--current-ppt", "empty.pptx"),
    ],
)
def test_incomplete_office_package_returns_bounded_error_without_history(
    flag: str,
    filename: str,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    source = tmp_path / filename
    with zipfile.ZipFile(source, "w") as archive:
        archive.writestr("placeholder", b"")
    data_dir = tmp_path / "data"

    code = _run(
        [
            flag,
            str(source),
            "--data-dir",
            str(data_dir),
            "--fail-on",
            "never",
        ]
    )

    captured = capsys.readouterr()
    assert code == 1
    assert filename in captured.err
    assert "required metadata is missing" in captured.err
    assert "Traceback" not in captured.err
    assert str(tmp_path) not in captured.err
    assert not (data_dir / "history.sqlite3").exists()


@pytest.mark.parametrize("password_args", [[], ["--password", "current_excel=wrong"]])
def test_encrypted_password_errors_are_bounded(
    password_args: list[str],
    fixture_dir: Path,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    data_dir = tmp_path / "data"

    code = _run(
        [
            "--current-excel",
            str(fixture_dir / "current_encrypted.xlsx"),
            "--data-dir",
            str(data_dir),
            "--fail-on",
            "never",
            *password_args,
        ]
    )

    captured = capsys.readouterr()
    assert code == 1
    assert "Traceback" not in captured.err
    assert str(fixture_dir) not in captured.err
    assert not (data_dir / "history.sqlite3").exists()


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
