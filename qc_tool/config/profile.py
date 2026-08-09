"""Deliverable profile: named, reusable per-deliverable QC configuration.

Profiles are YAML documents validated by these models. Every setting has a
sensible default so a run with no profile still works; recurring
deliverables save their corrections (region pins, keys, tolerances) once
and reuse them each cycle.
"""

import datetime as dt
import hashlib
import json
import os
import re
import tempfile
from pathlib import Path
from typing import Literal
from uuid import uuid4

import yaml
from pydantic import BaseModel, Field, field_validator, model_validator

from qc_tool.findings import FindingClass, Materiality, Severity
from qc_tool.package import MEMBER_ID_PATTERN
from qc_tool.security import private_directory, private_file

Orientation = Literal["long", "wide", "block"]
CadenceKind = Literal["month", "week", "quarter", "date"]
SeriesRole = Literal["actual", "forecast", "target"]
_PROFILE_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9 ._-]{0,79}$")


class NumericTolerance(BaseModel):
    """A numeric difference is a finding only if it exceeds both bounds."""

    absolute: float = 0.0
    relative: float = 0.0


class RestatementWindows(BaseModel):
    """Trailing distinct-period counts treated as restatement-prone.

    Constants restated inside the window are expected data-lag corrections
    (WARNING tier); outside it they threaten history integrity. A window of
    0 disables recency for that cadence. Date-keyed axes use the window of
    their inferred cadence (median spacing).
    """

    week: int = Field(default=8, ge=0)
    month: int = Field(default=2, ge=0)
    quarter: int = Field(default=1, ge=0)


class RegionOverride(BaseModel):
    """Pins one table region, replacing auto-detection for its sheet."""

    cell_range: str = Field(alias="range")
    orientation: Orientation
    header_row: int | None = None
    key_column: str | None = None  # column letter holding row identity

    model_config = {"populate_by_name": True}


class AcceptanceBand(BaseModel):
    """A declared per-range tolerance: in-band changes stay visible as INFO.

    Unlike the global suppress-style ``NumericTolerance``, an in-band change
    still produces a finding (tier ``within_tolerance``) so an auditor sees
    every accepted difference.
    """

    cell_range: str = Field(alias="range")
    absolute: float = Field(default=0.0, ge=0.0)
    relative: float = Field(default=0.0, ge=0.0)

    model_config = {"populate_by_name": True}


class CadenceBand(BaseModel):
    """An explicit contiguous period-axis range for one cadence kind."""

    cell_range: str = Field(alias="range")
    kind: CadenceKind

    model_config = {"populate_by_name": True}


class ExcelAvailabilityRule(BaseModel):
    """Period-aligned Excel cells whose future blankness is explicitly allowed."""

    name: str = ""
    cell_range: str = Field(alias="range")
    period_range: str = Field(alias="periods")
    required_through: str
    allow_blank_after: bool = True
    role: SeriesRole | None = None

    model_config = {"populate_by_name": True}


class SheetProfile(BaseModel):
    ignore: bool = False
    ignore_ranges: list[str] = Field(default_factory=list)
    #: Ranges holding cycle-snapshot aggregates: value changes there are
    #: expected each cadence (still reported, classed as expected), while
    #: format/style/formula findings remain fully active.
    refresh_ranges: list[str] = Field(default_factory=list)
    #: Declared numeric tolerance bands; in-band changes report as INFO.
    acceptance_bands: list[AcceptanceBand] = Field(default_factory=list)
    regions: list[RegionOverride] = Field(default_factory=list)
    cadence_bands: list[CadenceBand] = Field(default_factory=list)
    availability_rules: list[ExcelAvailabilityRule] = Field(default_factory=list)
    #: Chart-title or ``chart[N]`` overrides for rolling/full source windows.
    chart_windows: dict[str, Literal["rolling", "full"]] = Field(default_factory=dict)


class RangeControl(BaseModel):
    name: str = ""
    sheet: str
    cell_range: str = Field(alias="range")

    model_config = {"populate_by_name": True}


class UniqueRangeControl(RangeControl):
    skip_header: bool = True


class NumericBoundsControl(RangeControl):
    minimum: float | None = None
    maximum: float | None = None


class TieOutTerm(BaseModel):
    reference: str
    operation: Literal["add", "subtract"] = "add"


class TieOutControl(BaseModel):
    name: str
    target: str
    components: list[str] = Field(default_factory=list)
    terms: list[TieOutTerm] = Field(default_factory=list)
    absolute_tolerance: float = 0.0
    relative_tolerance: float = 0.0


class ExcelControls(BaseModel):
    required_ranges: list[RangeControl] = Field(default_factory=list)
    unique_ranges: list[UniqueRangeControl] = Field(default_factory=list)
    numeric_bounds: list[NumericBoundsControl] = Field(default_factory=list)
    tie_outs: list[TieOutControl] = Field(default_factory=list)


class ExcelMemberProfile(BaseModel):
    ignore_sheets: list[str] = Field(default_factory=list)
    sheets: dict[str, SheetProfile] = Field(default_factory=dict)
    controls: ExcelControls = Field(default_factory=ExcelControls)


class ExcelProfile(ExcelMemberProfile):
    members: dict[str, ExcelMemberProfile] = Field(
        default_factory=dict,
        exclude_if=lambda value: not value,
    )

    @field_validator("members")
    @classmethod
    def validate_members_keys(
        cls,
        value: dict[str, ExcelMemberProfile],
    ) -> dict[str, ExcelMemberProfile]:
        invalid = sorted(
            key for key in value if re.fullmatch(MEMBER_ID_PATTERN, key) is None
        )
        if invalid:
            raise ValueError(f"invalid Excel member ids: {invalid}")
        return value


class PptAvailabilityRule(BaseModel):
    """Blankness boundary for a table row or chart series on one slide."""

    name: str = ""
    slide: str
    scope: Literal["table", "chart"]
    element: str = ""
    series: str | None = None
    required_through: str
    allow_blank_after: bool = True
    role: SeriesRole | None = None


class PptProfile(BaseModel):
    #: Force slide pairings (baseline title -> current title); wins over fuzzy.
    slide_pins: dict[str, str] = Field(default_factory=dict)
    #: Minimum similarity score (0-100) to accept a fuzzy slide pair.
    match_threshold: float = 55.0
    #: Per-slide chart window override: "rolling" or "full" (else inferred).
    chart_windows: dict[str, Literal["rolling", "full"]] = Field(default_factory=dict)
    availability_rules: list[PptAvailabilityRule] = Field(default_factory=list)
    #: Slide titles that must exist in a final deck.
    required_slides: list[str] = Field(default_factory=list)
    #: Case-insensitive tokens that indicate unfinished presentation content.
    draft_tokens: list[str] = Field(
        default_factory=lambda: ["TBD", "XXX", "TODO", "PLACEHOLDER", "DRAFT"]
    )


class CrosscheckMapping(BaseModel):
    """Analyst-confirmed link from a deck figure to its workbook source cell.

    The figure itself changes every cycle, so the anchor is the stable
    numeric skeleton of its line (or ``table:<row>/<col>`` for table cells)
    plus the figure's ordinal within that line.
    """

    slide: str
    line_skeleton: str
    figure_index: int = 0
    label: str = ""
    source_sheet: str
    source_cell: str
    source_member: str = Field(
        default="primary",
        pattern=MEMBER_ID_PATTERN,
        exclude_if=lambda value: value == "primary",
    )


class CrosscheckProfile(BaseModel):
    mappings: list[CrosscheckMapping] = Field(default_factory=list)
    max_candidates: int = 5


class FindingWaiver(BaseModel):
    finding_class: FindingClass
    reason: str = Field(min_length=1)
    expires: dt.date
    sheet: str | None = None
    slide: str | None = None
    location: str | None = None
    element: str | None = None
    member: str = Field(
        default="primary",
        pattern=MEMBER_ID_PATTERN,
        exclude_if=lambda value: value == "primary",
    )


class DeliverableProfile(BaseModel):
    name: str
    #: Optional contract id assigned when promoting a profile to a contract.
    #: Empty string preserves legacy behaviour and must be omitted from the
    #: canonical profile JSON.
    contract_id: str = ""
    description: str = ""
    tolerance: NumericTolerance = Field(default_factory=NumericTolerance)
    restatement_windows: RestatementWindows = Field(default_factory=RestatementWindows)
    excel: ExcelProfile = Field(default_factory=ExcelProfile)
    ppt: PptProfile = Field(default_factory=PptProfile)
    crosscheck: CrosscheckProfile = Field(default_factory=CrosscheckProfile)
    waivers: list[FindingWaiver] = Field(default_factory=list)
    #: Per-class severity overrides applied to non-expected findings.
    severity: dict[FindingClass, Severity] = Field(default_factory=dict)
    #: Per-tier severity overrides; absent tiers use the built-in mapping
    #: (noise/within_tolerance -> info, recent_restatement -> warning).
    materiality_severity: dict[Materiality, Severity] = Field(default_factory=dict)

    def sheet_profile(self, sheet_name: str) -> SheetProfile | None:
        return self.excel.sheets.get(sheet_name)

    @model_validator(mode="after")
    def _validate_contract_id(self) -> "DeliverableProfile":
        cid = self.contract_id
        if cid and re.fullmatch(r"[0-9a-f]{32}", cid) is None:
            raise ValueError("contract_id must be a 32-character lowercase hex string")
        return self


def default_profile(name: str = "default") -> DeliverableProfile:
    return DeliverableProfile(name=name)


def new_profile(name: str) -> DeliverableProfile:
    """Create a named profile for UI-promoted contracts with a stable id.

    The produced profile sets `contract_id` to a uuid4().hex (lowercase
    hex) to be used as a contract scope. Legacy-created profiles retain an
    empty `contract_id`.
    """
    return DeliverableProfile(name=name, contract_id=uuid4().hex)


def canonical_profile_json(profile: DeliverableProfile) -> str:
    payload = profile.model_dump(mode="json", by_alias=True)
    # Preserve legacy canonicalization: omit empty contract_id from the
    # canonical representation so old profiles keep the same hash.
    if payload.get("contract_id") in (None, ""):
        payload.pop("contract_id", None)
    return json.dumps(payload, sort_keys=True, separators=(",", ":"))


def canonical_profile_bytes(profile: DeliverableProfile) -> bytes:
    return canonical_profile_json(profile).encode("utf-8")


def profile_sha256(profile: DeliverableProfile) -> str:
    return hashlib.sha256(canonical_profile_bytes(profile)).hexdigest()


def _legacy_excel_profile(profile: DeliverableProfile) -> ExcelMemberProfile:
    return ExcelMemberProfile(
        ignore_sheets=list(profile.excel.ignore_sheets),
        sheets=dict(profile.excel.sheets),
        controls=profile.excel.controls.model_copy(deep=True),
    )


def legacy_excel_profile_is_empty(profile: DeliverableProfile) -> bool:
    """Whether the unscoped compatibility profile carries no Excel rules."""
    return _legacy_excel_profile(profile) == ExcelMemberProfile()


def excel_profile_for_member(
    profile: DeliverableProfile,
    member_id: str,
    *,
    workbook_count: int,
) -> ExcelMemberProfile:
    """Resolve one member without applying unscoped rules ambiguously."""
    if workbook_count < 1:
        raise ValueError("workbook count must be positive")
    configured = profile.excel.members.get(member_id)
    if workbook_count == 1 and member_id == "primary" and configured is None:
        return _legacy_excel_profile(profile)
    if not legacy_excel_profile_is_empty(profile):
        raise ValueError(
            "legacy unscoped Excel rules are ambiguous for this package"
        )
    return configured.model_copy(deep=True) if configured else ExcelMemberProfile()


def profile_for_excel_member(
    profile: DeliverableProfile,
    member_id: str,
    workbook_count: int,
) -> DeliverableProfile:
    """Project one member into the existing single-workbook engine contract."""
    member_profile = excel_profile_for_member(
        profile,
        member_id,
        workbook_count=workbook_count,
    )
    projected = profile.model_copy(deep=True)
    projected.excel = ExcelProfile(
        ignore_sheets=list(member_profile.ignore_sheets),
        sheets=dict(member_profile.sheets),
        controls=member_profile.controls.model_copy(deep=True),
    )
    projected.waivers = [
        waiver for waiver in projected.waivers if waiver.member == member_id
    ]
    return projected


def profile_path(profiles_dir: Path, name: str) -> Path:
    if not _PROFILE_NAME_RE.fullmatch(name) or name in {".", ".."}:
        raise ValueError(
            "profile name must start with a letter or number and contain only "
            "letters, numbers, spaces, dots, dashes, or underscores"
        )
    return profiles_dir / f"{name}.yaml"


def load_profile(path: Path) -> DeliverableProfile:
    raw = yaml.safe_load(path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError(f"{path}: profile YAML must be a mapping")
    return DeliverableProfile.model_validate(raw)


def save_profile(profile: DeliverableProfile, path: Path) -> None:
    private_directory(path.parent)
    payload = profile.model_dump(mode="json", by_alias=True, exclude_defaults=True)
    payload["name"] = profile.name
    content = yaml.safe_dump(payload, sort_keys=False)
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=path.parent,
            prefix=f".{path.name}.",
            suffix=".tmp",
            delete=False,
        ) as handle:
            handle.write(content)
            temporary = Path(handle.name)
        private_file(temporary)
        if load_profile(temporary) != profile:
            raise ValueError("saved profile does not round-trip through YAML")
        os.replace(temporary, path)
        temporary = None
        private_file(path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
