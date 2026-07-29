"""Signed QC attestation bundle and tamper detection."""

import json
import os
import zipfile
from pathlib import Path

from qc_tool import cli
from qc_tool.attestation import (
    create_attestation,
    load_or_create_attestation_key,
    verify_attestation,
)
from qc_tool.coverage import QCRunMode
from qc_tool.ui.app import perform_run
from tests.conftest import fixture_profile


def test_attestation_verifies_and_detects_member_tampering(
    fixture_dir: Path, tmp_path: Path
) -> None:
    work_dir = tmp_path / "work"
    files = {"current_excel": fixture_dir / "current.xlsx"}
    profile = fixture_profile()
    artifacts = perform_run(
        work_dir,
        files,
        {},
        profile,
        mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
    )
    target = artifacts.result.findings[0]
    target.analyst_comment = "reviewed"
    target.severity_overridden = True
    key_path, key = load_or_create_attestation_key(work_dir)
    bundle = create_attestation(
        tmp_path / "run.qca",
        result=artifacts.result,
        profile=profile,
        input_files=files,
        report_paths=artifacts.report_paths,
        key=key,
    )

    assert verify_attestation(bundle, key=key).valid
    if os.name == "posix":
        assert key_path.stat().st_mode & 0o777 == 0o600
        assert bundle.stat().st_mode & 0o777 == 0o600
    with zipfile.ZipFile(bundle) as archive:
        manifest = json.loads(archive.read("manifest.json"))
        assert manifest["tool"]["version"]
        assert manifest["inputs"]["current_excel"]["sha256"]
        assert manifest["analyst_decisions"][0]["comment"] == "reviewed"
        members = {name: archive.read(name) for name in archive.namelist()}

    members["reports/qc_report.html"] += b"tampered"
    tampered = tmp_path / "tampered.qca"
    with zipfile.ZipFile(tampered, "w") as archive:
        for name, data in members.items():
            archive.writestr(name, data)
    result = verify_attestation(tampered, key=key)
    assert not result.valid
    assert any(issue.code in {"member-hash", "member-size"} for issue in result.issues)


def test_headless_attestation_cli(fixture_dir: Path, tmp_path: Path) -> None:
    data_dir = tmp_path / "data"
    bundle = tmp_path / "run.qca"
    code = cli.main(
        [
            "run",
            "--current-excel",
            str(fixture_dir / "current.xlsx"),
            "--data-dir",
            str(data_dir),
            "--fail-on",
            "never",
            "--attestation",
            str(bundle),
        ]
    )
    assert code == 0 and bundle.exists()
    assert (
        cli.main(
            [
                "verify-attestation",
                str(bundle),
                "--data-dir",
                str(data_dir),
            ]
        )
        == 0
    )
