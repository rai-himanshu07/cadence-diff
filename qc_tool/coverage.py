"""Explicit run-mode and check-coverage contracts shared by every surface."""

from enum import StrEnum

from pydantic import BaseModel, Field

from qc_tool.package import MEMBER_ID_PATTERN


class QCRunMode(StrEnum):
    CYCLE_COMPARISON = "cycle_comparison"
    CURRENT_FILE_PREFLIGHT = "current_file_preflight"
    FINAL_PACKAGE = "final_package"


class CoverageState(StrEnum):
    CHECKED = "checked"
    UNAVAILABLE = "unavailable"
    DEGRADED = "degraded"


class CoverageItem(BaseModel):
    check_id: str
    label: str
    artifact: str
    artifact_member: str = Field(
        default="primary",
        pattern=MEMBER_ID_PATTERN,
        exclude_if=lambda value: value == "primary",
    )
    state: CoverageState
    findings: int = 0
    detail: str = ""


class MappingCoverage(BaseModel):
    eligible: int = 0
    mapped: int = 0
    verified: int = 0
    mismatched: int = 0
    unresolved: int = 0
    unmapped: int = 0
    #: Material surfaces that carry claims nobody in this tool can read, such as
    #: figures rendered into a picture. Additive; zero for runs recorded before
    #: the population was made explicit.
    unavailable: int = 0

    @property
    def reconciles(self) -> bool:
        """Whether the mapped split accounts for exactly the eligible claims."""
        return self.mapped + self.unmapped == self.eligible


def capability_limited(coverage: list[CoverageItem]) -> bool:
    """Whether any required check could not run, so zero findings is not clean."""
    return any(item.state is CoverageState.UNAVAILABLE for item in coverage)
