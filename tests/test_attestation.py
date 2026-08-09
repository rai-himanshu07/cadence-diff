"""Signed QC attestation bundle and tamper detection."""

import json
import os
import zipfile
from pathlib import Path

from qc_tool import cli
from qc_tool.attestation import (
    AttestationSignoff,
    create_attestation,
    load_or_create_attestation_key,
    verify_attestation,
)
from qc_tool.coverage import QCRunMode
from qc_tool.engine import QCRunResult
from qc_tool.package import (
    PackageArtifact,
    PackageManifest,
    PackageMember,
    PackageSide,
)
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
        write_reports=True,
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
        assert manifest["schema_version"] == 1
        assert "package_manifest" not in manifest
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


def _package_manifest() -> PackageManifest:
    return PackageManifest(
        members=(
            PackageMember(
                member_id="core",
                side=PackageSide.CURRENT,
                artifact=PackageArtifact.EXCEL,
                display_name="core.xlsx",
            ),
            PackageMember(
                member_id="ops",
                side=PackageSide.CURRENT,
                artifact=PackageArtifact.EXCEL,
                display_name="ops.xlsx",
            ),
        )
    )


def test_signoff_without_true_package_remains_attestation_v2(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    source = fixture_dir / "current.xlsx"
    bundle = create_attestation(
        tmp_path / "v2.qca",
        result=QCRunResult(profile_name="signed"),
        profile=fixture_profile(),
        input_files={"current_excel": source},
        report_paths={},
        key=b"k" * 32,
        signoff=AttestationSignoff(
            finalized_at="2026-08-07T00:00:00+00:00",
            review_state_digest="review-state",
        ),
    )

    with zipfile.ZipFile(bundle) as archive:
        manifest = json.loads(archive.read("manifest.json"))
    assert manifest["schema_version"] == 2
    assert "signoff" in manifest
    assert "package_manifest" not in manifest
    assert verify_attestation(bundle, key=b"k" * 32).valid


def test_true_package_attestation_v3_binds_manifest_and_input_roles(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    manifest = _package_manifest()
    source = fixture_dir / "current.xlsx"
    files = {member.role_key: source for member in manifest.members}
    result = QCRunResult(
        profile_name="package",
        files={member.role_key: member.display_name for member in manifest.members},
        package_manifest=manifest,
    )

    bundle = create_attestation(
        tmp_path / "v3.qca",
        result=result,
        profile=fixture_profile(),
        input_files=files,
        report_paths={},
        key=b"p" * 32,
    )

    with zipfile.ZipFile(bundle) as archive:
        payload = json.loads(archive.read("manifest.json"))
    assert payload["schema_version"] == 3
    assert payload["package_manifest"] == manifest.model_dump(mode="json")
    assert set(payload["inputs"]) == set(manifest.role_keys)
    assert verify_attestation(bundle, key=b"p" * 32).valid


def test_v3_verifier_rejects_package_input_role_mismatch(
    fixture_dir: Path,
    tmp_path: Path,
) -> None:
    manifest = _package_manifest()
    result = QCRunResult(
        profile_name="package",
        package_manifest=manifest,
    )
    bundle = create_attestation(
        tmp_path / "role-mismatch.qca",
        result=result,
        profile=fixture_profile(),
        input_files={"current_excel:core": fixture_dir / "current.xlsx"},
        report_paths={},
        key=b"r" * 32,
    )

    verification = verify_attestation(bundle, key=b"r" * 32)

    assert not verification.valid
    assert any(issue.code == "package-inputs" for issue in verification.issues)


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
