"""Signed QC attestation bundle and tamper detection."""

import hashlib
import hmac
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


def test_attestation_binds_the_resolved_input_digest(tmp_path: Path) -> None:
    """plan-20260913 Step 4: the exact per-run resolved logical
    configuration digest is a signed, informational disclosure -- present
    for every run (empty string for a run with no saved input contract).
    """
    profile = fixture_profile()
    result = QCRunResult(profile_name=profile.name, resolved_input_digest="a" * 64)
    _, key = load_or_create_attestation_key(tmp_path)
    bundle = create_attestation(
        tmp_path / "run.qca",
        result=result,
        profile=profile,
        input_files={},
        report_paths={},
        key=key,
    )

    assert verify_attestation(bundle, key=key).valid
    with zipfile.ZipFile(bundle) as archive:
        manifest = json.loads(archive.read("manifest.json"))
    assert manifest["run"]["resolved_input_digest"] == "a" * 64


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


def test_attestation_discloses_resolved_formula_engines(
    fixture_dir: Path, tmp_path: Path
) -> None:
    """Criterion 5: the manifest carries the resolved formula-engine per
    excel role at every schema version -- a purely informational,
    always-present disclosure, not a new signed feature.
    """
    result = QCRunResult(
        profile_name="signed",
        formula_engines={
            "baseline_excel": "libreoffice:24.2.4.2",
            "current_excel": "native-biff12:1.2.3",
        },
        values_engines={
            "baseline_excel": "pyxlsb:1.0.10",
            "current_excel": "native-biff12:1.2.3",
        },
    )
    bundle = create_attestation(
        tmp_path / "engines.qca",
        result=result,
        profile=fixture_profile(),
        input_files={"current_excel": fixture_dir / "current.xlsx"},
        report_paths={},
        key=b"e" * 32,
    )

    with zipfile.ZipFile(bundle) as archive:
        manifest = json.loads(archive.read("manifest.json"))

    assert manifest["run"]["formula_engines"] == {
        "baseline_excel": "libreoffice:24.2.4.2",
        "current_excel": "native-biff12:1.2.3",
    }
    assert manifest["run"]["values_engines"] == {
        "baseline_excel": "pyxlsb:1.0.10",
        "current_excel": "native-biff12:1.2.3",
    }
    assert verify_attestation(bundle, key=b"e" * 32).valid


def test_verifier_rejects_resigned_invalid_run_metadata(
    fixture_dir: Path, tmp_path: Path
) -> None:
    key = b"m" * 32
    bundle = create_attestation(
        tmp_path / "run-metadata.qca",
        result=QCRunResult(profile_name="signed"),
        profile=fixture_profile(),
        input_files={"current_excel": fixture_dir / "current.xlsx"},
        report_paths={},
        key=key,
    )
    with zipfile.ZipFile(bundle) as archive:
        manifest = json.loads(archive.read("manifest.json"))
        members = {
            name: archive.read(name)
            for name in archive.namelist()
            if name != "manifest.json"
        }
    manifest["run"]["requested_output_mode"] = "random"
    unsigned = {name: value for name, value in manifest.items() if name != "signature"}
    payload = json.dumps(unsigned, sort_keys=True, separators=(",", ":")).encode()
    manifest["signature"]["value"] = hmac.new(key, payload, hashlib.sha256).hexdigest()
    tampered = tmp_path / "invalid-run-metadata.qca"
    with zipfile.ZipFile(tampered, "w") as archive:
        archive.writestr("manifest.json", json.dumps(manifest))
        for name, data in members.items():
            archive.writestr(name, data)

    verification = verify_attestation(tampered, key=key)

    assert not verification.valid
    assert any(issue.code == "run-metadata" for issue in verification.issues)


def _population_finding_for_signoff_test():
    from qc_tool.findings import Finding, FindingClass, MembershipCodec, PopulationEvidence

    return Finding(
        artifact="excel",
        finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
        sheet="Data",
        location="B2:B16",
        element="population",
        message="15 cells summarised as one population",
        population=PopulationEvidence(
            member_count=15,
            membership=MembershipCodec(
                current_rectangles=("B2:B16",),
                baseline_mode="shift",
                shift=(-1, 0),
                member_count=15,
            ),
            first="B2",
            last="B16",
            shape_before_digest="a" * 64,
            shape_after_digest="b" * 64,
        ),
    )


def test_v4_verifier_still_validates_a_malformed_signoff_alongside_a_population(
    fixture_dir: Path, tmp_path: Path
) -> None:
    """Criterion 10: schema_version reflects the HIGHEST tier feature
    present, not an exclusive mode -- a v4 (population) bundle can ALSO
    carry a sign-off, and that sign-off must still be validated, not
    silently skipped because the bundle's schema_version is 4, not 2.
    """
    result = QCRunResult(profile_name="pop", findings=[_population_finding_for_signoff_test()])
    key = b"s" * 32
    bundle = create_attestation(
        tmp_path / "pop-signoff.qca",
        result=result,
        profile=fixture_profile(),
        input_files={"current_excel": fixture_dir / "current.xlsx"},
        report_paths={},
        key=key,
        signoff=AttestationSignoff(
            finalized_at="2026-09-09T00:00:00+00:00",
            review_state_digest="review-state",
        ),
    )

    with zipfile.ZipFile(bundle) as archive:
        manifest = json.loads(archive.read("manifest.json"))
        members = {
            name: archive.read(name) for name in archive.namelist() if name != "manifest.json"
        }

    assert manifest["schema_version"] == 4
    assert "signoff" in manifest
    del manifest["signoff"]["review_state_digest"]  # now malformed

    unsigned = {k: v for k, v in manifest.items() if k != "signature"}
    payload = json.dumps(unsigned, sort_keys=True, separators=(",", ":")).encode("utf-8")
    manifest["signature"]["value"] = hmac.new(key, payload, hashlib.sha256).hexdigest()

    tampered = tmp_path / "tampered-signoff.qca"
    with zipfile.ZipFile(tampered, "w") as archive:
        archive.writestr("manifest.json", json.dumps(manifest))
        for name, data in members.items():
            archive.writestr(name, data)

    verification = verify_attestation(tampered, key=key)

    assert not verification.valid
    assert any(issue.code == "signoff" for issue in verification.issues)


def test_v4_verifier_rejects_an_unsupported_lineage_version(
    fixture_dir: Path, tmp_path: Path
) -> None:
    """Criterion 10: `lineage_version` is validated, not merely disclosed."""
    result = QCRunResult(profile_name="pop", findings=[_population_finding_for_signoff_test()])
    key = b"l" * 32
    bundle = create_attestation(
        tmp_path / "lineage.qca",
        result=result,
        profile=fixture_profile(),
        input_files={"current_excel": fixture_dir / "current.xlsx"},
        report_paths={},
        key=key,
    )

    with zipfile.ZipFile(bundle) as archive:
        manifest = json.loads(archive.read("manifest.json"))
        members = {
            name: archive.read(name) for name in archive.namelist() if name != "manifest.json"
        }

    assert manifest["lineage_version"] == 2
    manifest["lineage_version"] = 99

    unsigned = {k: v for k, v in manifest.items() if k != "signature"}
    payload = json.dumps(unsigned, sort_keys=True, separators=(",", ":")).encode("utf-8")
    manifest["signature"]["value"] = hmac.new(key, payload, hashlib.sha256).hexdigest()

    tampered = tmp_path / "tampered-lineage.qca"
    with zipfile.ZipFile(tampered, "w") as archive:
        archive.writestr("manifest.json", json.dumps(manifest))
        for name, data in members.items():
            archive.writestr(name, data)

    verification = verify_attestation(tampered, key=key)

    assert not verification.valid
    assert any(issue.code == "lineage-version" for issue in verification.issues)


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
