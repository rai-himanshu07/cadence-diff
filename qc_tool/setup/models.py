"""Typed, bounded output of the setup-analysis pipeline (plan-20260913,
Step 6).

Pure structural OBSERVATION of what is actually present in a baseline/
current file pair -- hidden/very-hidden sheets, flood-fill-detected
regions (with dynamic extents and period-axis orientation already computed
by ``qc_tool.excel.regions``), ranked-table composite-key evidence (via
``qc_tool.excel.ranked_identity``), a dependency-cost workload forecast
(via ``qc_tool.excel.complexity``), and -- for XLSB sources only -- the
existing structural risk classification (via ``qc_tool.io.xlsb_formula``)
that already answers "can this run without Office".

Deliberately does NOT resolve these observations against a saved
``WorkbookInputContract`` (``qc_tool.config.input_contract``) -- that
matching/confirmation step belongs to the configuration workspace UI and
its own resolution logic (a later step), once an analyst can review and
confirm ambiguous matches. Never carries raw cell values, formula text, or
selector values -- only bounded structural evidence already proven safe by
the existing detectors this module's own analysis pipeline composes.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from qc_tool.excel.complexity import WorkbookComplexity
from qc_tool.excel.ranked_identity import RankedTableCandidate
from qc_tool.excel.regions import TableRegion

#: Bump only when this shape changes in a way that invalidates an already
#: persisted sidecar (see ``qc_tool.setup.preview_store``).
SETUP_ANALYSIS_VERSION = 1

#: Strict output caps (Step 6's own "strict byte/query/output caps"
#: criterion) -- a pathological workbook degrades with a disclosed
#: truncation rather than growing the sidecar unboundedly.
MAX_SHEETS_PER_SIDE = 500
MAX_REGIONS_PER_SHEET = 64
MAX_FAILURE_DETAIL_CHARS = 256


@dataclass(frozen=True, slots=True)
class DetectedRegion:
    """One flood-fill-detected region plus its (optional) ranked-table
    composite-key evidence -- never a cell value, only the bounded,
    already-typed evidence ``TableRegion``/``RankedTableCandidate`` produce.
    """

    region: TableRegion
    #: Present only when a paired baseline region existed and the region
    #: cleared every ranked-candidate threshold; ``None`` means "not
    #: evaluated or nothing qualified", never a false negative to trust.
    ranked_candidate: RankedTableCandidate | None = None


@dataclass(frozen=True, slots=True)
class SheetSetupProfile:
    """One sheet's structural observation for one side (baseline/current)."""

    sheet_name: str
    hidden: bool = False
    very_hidden: bool = False
    regions: tuple[DetectedRegion, ...] = ()
    #: Set when THIS sheet's own detection failed (e.g. an unexpected shape
    #: tripped a detector); every other sheet's scan still completes --
    #: partial-analysis fallback, never a whole-scan abort for one sheet.
    failure_detail: str = ""


@dataclass(frozen=True, slots=True)
class XlsbRiskProfile:
    """Structural risk classification for an XLSB source, reusing the
    existing ``qc_tool.io.xlsb_formula`` scan verbatim -- ``None`` on the
    parent ``MemberSetupProfile`` field for xlsx/xlsm, where this does not
    apply. This scan needs no Excel/LibreOffice (the whole point of
    "XLSB no-Office behavior").
    """

    risky_features: tuple[str, ...] = ()
    passive_features: tuple[str, ...] = ()
    blocking_features: tuple[str, ...] = ()
    unknown_external_features: tuple[str, ...] = ()
    safe_for_external_engine: bool = False


@dataclass(frozen=True, slots=True)
class MemberSetupProfile:
    """One logical member's (or the single primary workbook's) structural
    observation across both sides of a cycle comparison.
    """

    member_id: str
    baseline_hash: str
    current_hash: str
    baseline_sheets: tuple[SheetSetupProfile, ...] = ()
    current_sheets: tuple[SheetSetupProfile, ...] = ()
    baseline_complexity: WorkbookComplexity | None = None
    current_complexity: WorkbookComplexity | None = None
    baseline_xlsb_risk: XlsbRiskProfile | None = None
    current_xlsb_risk: XlsbRiskProfile | None = None
    #: Set when the WHOLE member's scan failed (e.g. the file could not be
    #: opened at all -- wrong password, corrupt archive); every other
    #: member's scan still completes.
    failure_detail: str = ""


@dataclass(frozen=True, slots=True)
class SetupAnalysisResult:
    """Complete, bounded output of one setup-analysis scan."""

    version: int = SETUP_ANALYSIS_VERSION
    generated_at: str = ""
    #: Keyed by member id ("primary" for a single-workbook-pair run).
    members: dict[str, MemberSetupProfile] = field(default_factory=dict)
    #: Set only when the run-level attempt failed before any member could
    #: be scanned (e.g. a shared password prompt was never answered).
    failure_detail: str = ""
