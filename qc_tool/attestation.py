"""Signed QC attestation bundles and tamper verification."""

import datetime as dt
import hashlib
import hmac
import json
import os
import secrets
import zipfile
from pathlib import Path

from pydantic import BaseModel, Field, field_validator, model_validator

from qc_tool import __version__
from qc_tool.config.profile import (
    DeliverableProfile,
    ResolvedOutputPolicy,
    canonical_profile_bytes,
)
from qc_tool.coverage import FindingOutputMode, QCRunMode
from qc_tool.engine import QCRunResult
from qc_tool.history.review_state import ANNOTATION_LINEAGE_VERSION, AnnotationLineage
from qc_tool.package import PackageManifest
from qc_tool.report.json_report import result_payload
from qc_tool.review import population_identity_digest
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


class PopulationManifestEntry(BaseModel):
    """Signed disclosure for one group-first population finding (schema v4)."""

    identity_key_digest: str
    member_count: int
    membership_digest: str


class AttestationRunMetadata(BaseModel):
    """Validated signed run contract, with defaults for legacy bundles."""

    mode: QCRunMode
    profile: str
    counts: dict[str, int]
    coverage: list[dict[str, object]]
    mapping_coverage: dict[str, object] | None = None
    verified_crosschecks: int
    formula_engines: dict[str, str] = Field(default_factory=dict)
    values_engines: dict[str, str] = Field(default_factory=dict)
    requested_output_mode: FindingOutputMode = FindingOutputMode.PROFILE
    resolved_output_policy: ResolvedOutputPolicy | None = None

    model_config = {"extra": "forbid"}

    @field_validator("formula_engines", "values_engines")
    @classmethod
    def validate_engine_map(cls, value: dict[str, str]) -> dict[str, str]:
        if any(not role or not engine for role, engine in value.items()):
            raise ValueError("engine provenance requires non-empty roles and identities")
        return value

    @model_validator(mode="after")
    def validate_output_policy(self) -> "AttestationRunMetadata":
        if self.resolved_output_policy is not None:
            if self.resolved_output_policy.output_mode is not self.requested_output_mode:
                raise ValueError("resolved output policy does not match requested output mode")
        elif self.requested_output_mode is not FindingOutputMode.PROFILE:
            raise ValueError("non-profile output mode requires a resolved output policy")
        return self


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _canonical(payload: dict) -> bytes:
    return json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")


def _population_manifest(result: QCRunResult) -> list[PopulationManifestEntry] | None:
    entries = [
        PopulationManifestEntry(
            identity_key_digest=population_identity_digest(finding),
            member_count=finding.population.member_count,
            membership_digest=_sha256_bytes(
                _canonical(finding.population.membership.model_dump(mode="json"))
            ),
        )
        for finding in result.findings
        if finding.population is not None
    ]
    return entries or None


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
    package_manifest = (
        result.package_manifest
        if result.package_manifest is not None
        and not result.package_manifest.is_legacy_projection
        else None
    )
    population_manifest = _population_manifest(result)
    unsigned = {
        "schema_version": (
            4
            if population_manifest is not None
            else (3 if package_manifest is not None else (2 if signoff is not None else 1))
        ),
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
            # Resolved formula-engine/adapter-fingerprint string per excel
            # role (Criterion 5): a purely informational disclosure so a
            # cross-run comparison can later be told an engine changed,
            # never a new schema feature -- present at every schema version.
            "formula_engines": dict(result.formula_engines),
            "values_engines": dict(result.values_engines),
            # Run-level finding-output contract (plan-20260910): purely
            # informational disclosure, present at every schema version,
            # never a verifier gate.
            "requested_output_mode": result.requested_output_mode.value,
            "resolved_output_policy": (
                result.resolved_output_policy.model_dump(mode="json")
                if result.resolved_output_policy is not None
                else None
            ),
        },
        "profile_sha256": _sha256_bytes(profile_bytes),
        "inputs": inputs,
        "members": member_manifest,
        "waivers": [item.model_dump(mode="json") for item in profile.waivers],
        "analyst_decisions": analyst_decisions,
    }
    if package_manifest is not None:
        unsigned["package_manifest"] = package_manifest.model_dump(mode="json")
    if signoff is not None:
        unsigned["signoff"] = signoff.model_dump(mode="json")
    if population_manifest is not None:
        unsigned["population_manifest"] = [
            entry.model_dump(mode="json") for entry in population_manifest
        ]
        unsigned["lineage_version"] = 2
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
        if not isinstance(manifest, dict):
            result.add("invalid-manifest", "manifest.json is not a JSON object")
            return result
        schema_version = manifest.get("schema_version")
        if schema_version not in {1, 2, 3, 4}:
            result.add(
                "schema-version", f"unsupported attestation schema {schema_version!r}"
            )

        try:
            AttestationRunMetadata.model_validate(manifest.get("run"))
        except (TypeError, ValueError):
            result.add("run-metadata", "signed run metadata is missing or invalid")

        # Cumulative, presence-driven feature validation: a v3/v4 bundle can
        # ALSO carry a sign-off (schema_version reflects the HIGHEST tier
        # feature present, per `create_attestation`, not an exclusive mode),
        # so every optional feature is validated whenever ITS OWN key is
        # present, never skipped just because a higher-tier feature also
        # exists in the same bundle.
        if "signoff" in manifest:
            try:
                AttestationSignoff.model_validate(manifest.get("signoff"))
            except (TypeError, ValueError):
                result.add("signoff", "sign-off evidence is missing or invalid")
        elif schema_version == 2:
            result.add("missing-signoff", "schema v2 sign-off evidence is missing")

        package_manifest: PackageManifest | None = None
        if "package_manifest" in manifest:
            try:
                package_manifest = PackageManifest.model_validate(
                    manifest.get("package_manifest")
                )
            except (TypeError, ValueError):
                result.add("package-manifest", "package manifest is missing or invalid")
        elif schema_version == 3:
            result.add(
                "package-manifest", "schema v3 package manifest is missing or invalid"
            )

        if "population_manifest" in manifest:
            population_manifest = manifest.get("population_manifest")
            if not isinstance(population_manifest, list):
                result.add(
                    "population-manifest", "population manifest is missing or invalid"
                )
            else:
                for entry in population_manifest:
                    try:
                        PopulationManifestEntry.model_validate(entry)
                    except (TypeError, ValueError):
                        result.add(
                            "population-manifest-entry",
                            "a population manifest entry does not match its schema",
                        )
                        break
        elif schema_version == 4:
            result.add(
                "population-manifest", "schema v4 population manifest is missing or invalid"
            )

        if "lineage_version" in manifest:
            if manifest.get("lineage_version") != ANNOTATION_LINEAGE_VERSION:
                result.add(
                    "lineage-version",
                    f"unsupported annotation-lineage schema {manifest.get('lineage_version')!r}",
                )
            signoff_payload = manifest.get("signoff")
            lineage_rows = (
                signoff_payload.get("annotation_lineage")
                if isinstance(signoff_payload, dict)
                else None
            )
            if lineage_rows is not None and not isinstance(lineage_rows, list):
                result.add("lineage-row", "annotation-lineage rows are not a list")
            elif isinstance(lineage_rows, list):
                for row in lineage_rows:
                    try:
                        AnnotationLineage.model_validate(row)
                    except (TypeError, ValueError):
                        result.add(
                            "lineage-row",
                            "an annotation-lineage row does not match its declared schema",
                        )
                        break

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
        inputs = manifest.get("inputs")
        if package_manifest is not None:
            expected_roles = {
                member.role_key for member in package_manifest.members
            }
            actual_roles = set(inputs) if isinstance(inputs, dict) else set()
            if actual_roles != expected_roles:
                result.add(
                    "package-inputs",
                    "attested input roles do not match the package manifest",
                )
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
