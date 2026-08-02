"""Deliverable profile: named, reusable per-deliverable QC configuration.

Profiles are YAML documents validated by these models. Every setting has a
sensible default so a run with no profile still works; recurring
deliverables save their corrections (region pins, keys, tolerances) once
and reuse them each cycle.
"""

import datetime as dt
from pathlib import Path
from typing import Literal

import yaml
from pydantic import BaseModel, Field

from qc_tool.findings import FindingClass, Materiality, Severity
from qc_tool.security import private_directory, private_file

Orientation = Literal["long", "wide", "block"]
CadenceKind = Literal["month", "week", "quarter", "date"]
SeriesRole = Literal["actual", "forecast", "target"]


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


class ExcelProfile(BaseModel):
    ignore_sheets: list[str] = Field(default_factory=list)
    sheets: dict[str, SheetProfile] = Field(default_factory=dict)
    controls: ExcelControls = Field(default_factory=ExcelControls)


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


class DeliverableProfile(BaseModel):
    name: str
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


def default_profile(name: str = "default") -> DeliverableProfile:
    return DeliverableProfile(name=name)


def load_profile(path: Path) -> DeliverableProfile:
    raw = yaml.safe_load(path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError(f"{path}: profile YAML must be a mapping")
    return DeliverableProfile.model_validate(raw)


def save_profile(profile: DeliverableProfile, path: Path) -> None:
    private_directory(path.parent)
    payload = profile.model_dump(by_alias=True, exclude_defaults=True)
    payload["name"] = profile.name
    path.write_text(yaml.safe_dump(payload, sort_keys=False), encoding="utf-8")
    private_file(path)
