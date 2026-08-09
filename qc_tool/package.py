"""Versioned package/member identity without paths, hashes, or credentials."""

from __future__ import annotations

from collections.abc import Mapping
from enum import StrEnum
from pathlib import Path
from typing import Literal, Self

from pydantic import BaseModel, Field, field_validator, model_validator

MAX_WORKBOOKS_PER_SIDE = 8
MEMBER_ID_PATTERN = r"^[a-z][a-z0-9_-]{0,31}$"
_LEGACY_ROLES = frozenset(
    {"baseline_excel", "current_excel", "baseline_ppt", "current_ppt"}
)


class PackageSide(StrEnum):
    BASELINE = "baseline"
    CURRENT = "current"


class PackageArtifact(StrEnum):
    EXCEL = "excel"
    PPT = "ppt"


_SIDE_ORDER = {PackageSide.BASELINE: 0, PackageSide.CURRENT: 1}
_ARTIFACT_ORDER = {PackageArtifact.EXCEL: 0, PackageArtifact.PPT: 1}


class PackageMember(BaseModel):
    member_id: str = Field(pattern=MEMBER_ID_PATTERN)
    side: PackageSide
    artifact: PackageArtifact
    display_name: str = Field(min_length=1, max_length=255)

    model_config = {"frozen": True}

    @field_validator("display_name")
    @classmethod
    def normalize_display_name(cls, value: str) -> str:
        normalized = value.strip()
        if not normalized:
            raise ValueError("package display name cannot be blank")
        return normalized

    @property
    def role_key(self) -> str:
        prefix = f"{self.side.value}_{self.artifact.value}"
        return prefix if self.member_id == "primary" else f"{prefix}:{self.member_id}"


def _member_sort_key(
    member: PackageMember,
) -> tuple[int, int, str]:
    return (
        _SIDE_ORDER[member.side],
        _ARTIFACT_ORDER[member.artifact],
        member.member_id,
    )


class PackageManifest(BaseModel):
    version: Literal[1] = 1
    members: tuple[PackageMember, ...] = ()

    model_config = {"frozen": True}

    @field_validator("members", mode="after")
    @classmethod
    def canonicalize_members(
        cls,
        members: tuple[PackageMember, ...],
    ) -> tuple[PackageMember, ...]:
        return tuple(sorted(members, key=_member_sort_key))

    @model_validator(mode="after")
    def validate_population(self) -> Self:
        identities = {
            (member.side, member.artifact, member.member_id)
            for member in self.members
        }
        if len(identities) != len(self.members):
            raise ValueError("package member identities must be unique")
        role_keys = self.role_keys
        if len(set(role_keys)) != len(role_keys):
            raise ValueError("package role keys must be unique")
        for side in PackageSide:
            if (
                len(self.members_for(side, PackageArtifact.EXCEL))
                > MAX_WORKBOOKS_PER_SIDE
            ):
                raise ValueError(
                    f"no more than {MAX_WORKBOOKS_PER_SIDE} Excel members per side"
                )
            if len(self.members_for(side, PackageArtifact.PPT)) > 1:
                raise ValueError("no more than one PowerPoint member per side")
        return self

    @property
    def role_keys(self) -> tuple[str, ...]:
        return tuple(member.role_key for member in self.members)

    @property
    def is_legacy_projection(self) -> bool:
        return all(member.member_id == "primary" for member in self.members)

    def members_for(
        self,
        side: PackageSide,
        artifact: PackageArtifact,
    ) -> tuple[PackageMember, ...]:
        return tuple(
            member
            for member in self.members
            if member.side is side and member.artifact is artifact
        )

    def member(
        self,
        side: PackageSide,
        artifact: PackageArtifact,
        member_id: str,
    ) -> PackageMember:
        matches = (
            member
            for member in self.members
            if member.side is side
            and member.artifact is artifact
            and member.member_id == member_id
        )
        try:
            return next(matches)
        except StopIteration as exc:
            raise KeyError("no such package member") from exc

    def to_legacy_roles(self) -> dict[str, str]:
        if not self.is_legacy_projection:
            raise ValueError("manifest is not a primary-role projection")
        return {member.role_key: member.display_name for member in self.members}

    @classmethod
    def from_legacy_files(cls, files: Mapping[str, object]) -> Self:
        unknown = set(files) - _LEGACY_ROLES
        if unknown:
            raise ValueError(f"unknown legacy package roles: {sorted(unknown)}")
        return cls.from_role_files(files)

    @classmethod
    def from_role_files(cls, files: Mapping[str, object]) -> Self:
        members: list[PackageMember] = []
        for role_key, source in files.items():
            if source is None:
                continue
            prefix, separator, member_id = role_key.partition(":")
            if not separator:
                member_id = "primary"
            prefix_parts = prefix.split("_")
            if len(prefix_parts) != 2:
                raise ValueError(f"invalid package role {role_key!r}")
            side, artifact = prefix_parts
            raw_name = getattr(source, "name", None)
            display_name = str(raw_name or Path(str(source)).name)
            members.append(
                PackageMember(
                    member_id=member_id,
                    side=PackageSide(side),
                    artifact=PackageArtifact(artifact),
                    display_name=display_name,
                )
            )
        return cls(members=tuple(members))


def paths_by_member(
    files: Mapping[str, object],
    manifest: PackageManifest,
) -> dict[str, object]:
    """Validate exact parity and return the same objects by stable role key."""
    expected = set(manifest.role_keys)
    provided = {key for key, value in files.items() if value is not None}
    if expected != provided:
        raise ValueError(
            "package role mismatch; "
            f"missing={sorted(expected - provided)}; "
            f"extra={sorted(provided - expected)}"
        )
    return {member.role_key: files[member.role_key] for member in manifest.members}


def structural_member_aliases(
    manifest: PackageManifest,
) -> dict[tuple[PackageArtifact, str], str]:
    """Privacy-safe aliases that preserve same-ID pairing across package sides."""
    aliases: dict[tuple[PackageArtifact, str], str] = {}
    for artifact in PackageArtifact:
        ordinal = 0
        for member_id in sorted(
            {
                member.member_id
                for member in manifest.members
                if member.artifact is artifact
            }
        ):
            if member_id == "primary":
                aliases[(artifact, member_id)] = "primary"
                continue
            ordinal += 1
            aliases[(artifact, member_id)] = f"{artifact.value}_{ordinal:03d}"
    return aliases
