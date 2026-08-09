"""Versioned private focus-target contract stored beside a recorded run.

The sidecar maps a finding id to zero or more role-specific target seeds. It
carries no path, hash, formula, value, or Office object name: byte identity
always comes from the run's own ``file_hashes``.
"""

from __future__ import annotations

import json
import logging
from collections.abc import Mapping, Sequence
from enum import StrEnum
from typing import Self

from pydantic import BaseModel, ConfigDict, Field, ValidationError, model_validator

from qc_tool.package import MEMBER_ID_PATTERN

logger = logging.getLogger(__name__)

#: Bumped whenever a stored seed changes meaning. A run recorded under any other
#: version yields no focus actions instead of a guessed interpretation.
FOCUS_SIDECAR_VERSION = 1

#: Column default for legacy rows and for runs whose generation failed.
EMPTY_SIDECAR_JSON = "{}"


class FocusArtifact(StrEnum):
    EXCEL = "excel"
    PPT = "ppt"


class FocusRole(StrEnum):
    BASELINE_EXCEL = "baseline_excel"
    CURRENT_EXCEL = "current_excel"
    BASELINE_PPT = "baseline_ppt"
    CURRENT_PPT = "current_ppt"

    @property
    def artifact(self) -> FocusArtifact:
        if self in {FocusRole.BASELINE_EXCEL, FocusRole.CURRENT_EXCEL}:
            return FocusArtifact.EXCEL
        return FocusArtifact.PPT


def focus_role_key(role: FocusRole, member_id: str = "primary") -> str:
    return role.value if member_id == "primary" else f"{role.value}:{member_id}"


class FocusTargetSeed(BaseModel):
    """One role-specific location inside one already-open saved document."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    artifact: FocusArtifact
    role: FocusRole
    member_id: str = Field(default="primary", pattern=MEMBER_ID_PATTERN)
    sheet: str | None = None
    address: str | None = None
    slide_index: int | None = Field(default=None, ge=1)
    shape_id: int | None = Field(default=None, ge=1)

    @model_validator(mode="after")
    def check_shape(self) -> Self:
        if self.role.artifact is not self.artifact:
            raise ValueError("focus seed role and artifact disagree")
        if self.artifact is FocusArtifact.EXCEL:
            if self.slide_index is not None or self.shape_id is not None:
                raise ValueError("an Excel focus seed cannot carry slide fields")
            if not self.sheet:
                raise ValueError("an Excel focus seed requires a sheet name")
        else:
            if self.sheet is not None or self.address is not None:
                raise ValueError("a PowerPoint focus seed cannot carry sheet fields")
            if self.slide_index is None:
                raise ValueError("a PowerPoint focus seed requires a slide index")
        return self

    @property
    def role_key(self) -> str:
        return focus_role_key(self.role, self.member_id)


class FocusTargetSidecar(BaseModel):
    """Decoded sidecar; ``version`` 0 means legacy, absent, or unreadable."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    version: int = 0
    targets: dict[str, tuple[FocusTargetSeed, ...]] = Field(default_factory=dict)

    @property
    def usable(self) -> bool:
        return self.version == FOCUS_SIDECAR_VERSION and bool(self.targets)

    def seeds(self, finding_id: str) -> tuple[FocusTargetSeed, ...]:
        if self.version != FOCUS_SIDECAR_VERSION:
            return ()
        return self.targets.get(finding_id, ())

    def seed(
        self,
        finding_id: str,
        role: FocusRole,
        member_id: str = "primary",
    ) -> FocusTargetSeed | None:
        """The single seed for one finding and role, or ``None`` if ambiguous."""
        matches = [
            seed
            for seed in self.seeds(finding_id)
            if seed.role is role and seed.member_id == member_id
        ]
        if len(matches) != 1:
            return None
        return matches[0]


def encode_focus_targets(
    targets: Mapping[str, Sequence[FocusTargetSeed]],
) -> str:
    """Serialize generated seeds for the private ``runs.focus_targets`` column."""
    stored = {
        finding_id: [seed.model_dump(mode="json") for seed in seeds]
        for finding_id, seeds in sorted(targets.items())
        if seeds
    }
    if not stored:
        return EMPTY_SIDECAR_JSON
    return json.dumps({"version": FOCUS_SIDECAR_VERSION, "targets": stored})


def decode_focus_targets(raw: str | None) -> FocusTargetSidecar:
    """Read a stored sidecar; anything unreadable yields no targets at all."""
    if not raw or raw == EMPTY_SIDECAR_JSON:
        return FocusTargetSidecar()
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError:
        logger.warning("focus-sidecar-unreadable")
        return FocusTargetSidecar()
    if not isinstance(payload, dict):
        logger.warning("focus-sidecar-unreadable")
        return FocusTargetSidecar()
    if payload.get("version") != FOCUS_SIDECAR_VERSION:
        return FocusTargetSidecar()
    try:
        return FocusTargetSidecar.model_validate(payload)
    except ValidationError:
        logger.warning("focus-sidecar-invalid")
        return FocusTargetSidecar()
