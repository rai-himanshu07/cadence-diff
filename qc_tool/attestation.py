"""Signed QC attestation bundles and tamper verification."""

import datetime as dt
import hashlib
import hmac
import json
import os
import secrets
import zipfile
from pathlib import Path

from pydantic import BaseModel, Field

from qc_tool import __version__
from qc_tool.config.profile import DeliverableProfile, canonical_profile_bytes
from qc_tool.engine import QCRunResult
from qc_tool.report.json_report import result_payload
from qc_tool.security import private_directory, private_file

_KEY_BYTES = 32


class AttestationIssue(BaseModel):
    code: str
    message: str


class AttestationVerification(BaseModel):
    valid: bool
    key_id: str = ""
    issues: list[AttestationIssue] = Field(default_factory=list)

    def add(self, code: str, message: str) -> None:
        self.valid = False
        self.issues.append(AttestationIssue(code=code, message=message))


class AttestationSignoff(BaseModel):
    finalized_at: str
    acknowledgements: tuple[str, ...] = ()
    review_state_digest: str
    annotation_lineage: list[dict[str, object]] = Field(default_factory=list)


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _canonical(payload: dict) -> bytes:
    return json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")


def _key_id(key: bytes) -> str:
    return hashlib.sha256(key).hexdigest()[:16]


def load_or_create_attestation_key(data_dir: Path) -> tuple[Path, bytes]:
    private_directory(data_dir)
    key_path = data_dir / "attestation.key"
    if key_path.exists():
        key = key_path.read_bytes()
    else:
        key = secrets.token_bytes(_KEY_BYTES)
        key_path.write_bytes(key)
        private_file(key_path)
    if len(key) < _KEY_BYTES:
        raise ValueError(f"attestation key {key_path} must contain at least {_KEY_BYTES} bytes")
    if os.name == "posix" and key_path.stat().st_mode & 0o077:
        raise ValueError(f"attestation key {key_path} must not be group/world accessible")
    return key_path, key


def load_attestation_key(path: Path) -> bytes:
    key = path.read_bytes()
    if len(key) < _KEY_BYTES:
        raise ValueError(f"attestation key {path} must contain at least {_KEY_BYTES} bytes")
    if os.name == "posix" and path.stat().st_mode & 0o077:
        raise ValueError(f"attestation key {path} must not be group/world accessible")
    return key


def create_attestation(
    output: Path,
    *,
    result: QCRunResult,
    profile: DeliverableProfile,
    input_files: dict[str, Path],
    report_paths: dict[str, Path],
    key: bytes,
    signoff: AttestationSignoff | None = None,
) -> Path:
    """Create a private signed bundle containing evidence and report members."""
    profile_bytes = canonical_profile_bytes(profile)
    findings_bytes = json.dumps(
        result_payload(result, include_context=True), indent=2
    ).encode("utf-8")
    members: dict[str, bytes] = {
        "evidence/profile.json": profile_bytes,
        "evidence/findings.json": findings_bytes,
    }
    for kind, path in sorted(report_paths.items()):
        members[f"reports/qc_report.{path.suffix.lstrip('.') or kind}"] = path.read_bytes()

    member_manifest = {
        name: {"sha256": _sha256_bytes(data), "size": len(data)}
        for name, data in sorted(members.items())
    }
    inputs = {
        role: {
            "name": path.name,
            "sha256": _sha256_bytes(path.read_bytes()),
            "size": path.stat().st_size,
        }
        for role, path in sorted(input_files.items())
    }
    analyst_decisions = [
        {
            "finding_id": finding.finding_id,
            "severity": finding.severity.value if finding.severity else None,
            "severity_overridden": finding.severity_overridden,
            "comment": finding.analyst_comment,
            "waiver_reason": finding.waiver_reason,
            "waiver_expires": finding.waiver_expires,
        }
        for finding in result.findings
        if finding.severity_overridden or finding.analyst_comment or finding.waiver_reason
    ]
    unsigned = {
        "schema_version": 2 if signoff is not None else 1,
        "generated_at": dt.datetime.now(dt.UTC).isoformat(timespec="seconds"),
        "tool": {"distribution": "cadence-diff", "version": __version__},
        "run": {
            "mode": result.mode.value,
            "profile": result.profile_name,
            "counts": {severity.value: count for severity, count in result.counts.items()},
            "coverage": [item.model_dump(mode="json") for item in result.coverage],
            "mapping_coverage": (
                result.mapping_coverage.model_dump(mode="json")
                if result.mapping_coverage is not None
                else None
            ),
            "verified_crosschecks": result.verified_crosschecks,
        },
        "profile_sha256": _sha256_bytes(profile_bytes),
        "inputs": inputs,
        "members": member_manifest,
        "waivers": [item.model_dump(mode="json") for item in profile.waivers],
        "analyst_decisions": analyst_decisions,
    }
    if signoff is not None:
        unsigned["signoff"] = signoff.model_dump(mode="json")
    signature = hmac.new(key, _canonical(unsigned), hashlib.sha256).hexdigest()
    manifest = {
        **unsigned,
        "signature": {
            "algorithm": "HMAC-SHA256",
            "key_id": _key_id(key),
            "value": signature,
        },
    }
    private_directory(output.parent)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("manifest.json", json.dumps(manifest, indent=2))
        for name, data in members.items():
            archive.writestr(name, data)
    private_file(output)
    return output


def verify_attestation(path: Path, *, key: bytes) -> AttestationVerification:
    result = AttestationVerification(valid=True, key_id=_key_id(key))
    try:
        archive = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        result.add("bad-zip", "attestation is not a valid ZIP archive")
        return result
    with archive:
        names = set(archive.namelist())
        if "manifest.json" not in names:
            result.add("missing-manifest", "manifest.json is missing")
            return result
        try:
            manifest = json.loads(archive.read("manifest.json"))
        except (json.JSONDecodeError, UnicodeDecodeError):
            result.add("invalid-manifest", "manifest.json is not valid JSON")
            return result
        schema_version = manifest.get("schema_version")
        if schema_version not in {1, 2}:
            result.add(
                "schema-version", f"unsupported attestation schema {schema_version!r}"
            )
        if schema_version == 2 and not isinstance(manifest.get("signoff"), dict):
            result.add("missing-signoff", "schema v2 sign-off evidence is missing")
        signature = manifest.pop("signature", None)
        if not isinstance(signature, dict):
            result.add("missing-signature", "manifest signature is missing")
            return result
        if signature.get("algorithm") != "HMAC-SHA256":
            result.add("signature-algorithm", "unsupported signature algorithm")
        if signature.get("key_id") != _key_id(key):
            result.add("key-id", "attestation was signed with a different key")
        expected = hmac.new(key, _canonical(manifest), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(str(signature.get("value", "")), expected):
            result.add("signature", "manifest signature does not verify")
        members = manifest.get("members")
        if not isinstance(members, dict):
            result.add("member-manifest", "member hash manifest is invalid")
            return result
        expected_names = {"manifest.json", *members}
        unexpected = names - expected_names
        missing = expected_names - names
        if unexpected:
            result.add("unexpected-members", f"unexpected members: {sorted(unexpected)}")
        if missing:
            result.add("missing-members", f"missing members: {sorted(missing)}")
        for name, metadata in members.items():
            if name not in names or not isinstance(metadata, dict):
                continue
            data = archive.read(name)
            if metadata.get("sha256") != _sha256_bytes(data):
                result.add("member-hash", f"hash mismatch for {name}")
            if metadata.get("size") != len(data):
                result.add("member-size", f"size mismatch for {name}")
    return result
