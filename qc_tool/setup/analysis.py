"""Pure orchestration for the setup-analysis pipeline (plan-20260913,
Step 6): composes already-proven detectors over already-loaded snapshots.

Never opens a file itself -- every input here is an already-loaded
``WorkbookSnapshot`` (or a pre-computed ``XlsbRiskProfile`` the caller
obtained separately from raw bytes). This mirrors the established
"detectors are pure functions over loaded snapshots" convention already
used by ``qc_tool.excel.regions``/``qc_tool.excel.ranked_identity``/
``qc_tool.excel.complexity`` -- this module is a thin, testable composition
layer over them, not a new detector of its own. The impure I/O (loading
files, reading raw XLSB bytes) belongs to ``qc_tool.setup.preview_worker``.

Partial-analysis fallback: one sheet's (or one whole member's) detector
failure never aborts the rest of the scan -- see ``SheetSetupProfile.
failure_detail``/``MemberSetupProfile.failure_detail``.
"""

from __future__ import annotations

import time
from dataclasses import dataclass, field

from qc_tool.excel.complexity import (
    WorkbookComplexity,
    assess_workbook_complexity,
    finalize_workbook_complexity,
)
from qc_tool.excel.ranked_identity import detect_ranked_table_candidate
from qc_tool.excel.regions import TableRegion, detect_regions
from qc_tool.io.model import SheetSnapshot, WorkbookSnapshot
from qc_tool.progress import CancellationToken, check_cancelled
from qc_tool.setup.models import (
    MAX_FAILURE_DETAIL_CHARS,
    MAX_REGIONS_PER_SHEET,
    MAX_SHEETS_PER_SIDE,
    DetectedRegion,
    MemberSetupProfile,
    SheetSetupProfile,
    XlsbRiskProfile,
)
from qc_tool.setup.preview_store import SetupScanStore


def _bounded_detail(detail: str) -> str:
    return detail[:MAX_FAILURE_DETAIL_CHARS]


def _detect_sheet_regions(sheet: SheetSnapshot) -> tuple[tuple[TableRegion, ...], str]:
    """``(regions, failure_detail)``. Never raises -- a detector failure is
    isolated to this one sheet, never the whole scan.
    """
    try:
        return tuple(detect_regions(sheet)[:MAX_REGIONS_PER_SHEET]), ""
    except Exception as exc:
        # Isolate one sheet's detector failure -- never abort the whole scan.
        return (), _bounded_detail(str(exc))


def _detect_ranked_candidate(
    baseline_sheet: SheetSnapshot,
    current_sheet: SheetSnapshot,
    baseline_region: TableRegion,
    current_region: TableRegion,
    *,
    allow_manual_review: bool,
):
    if current_region.orientation != "block":
        return None
    try:
        return detect_ranked_table_candidate(
            baseline_sheet,
            current_sheet,
            baseline_region,
            current_region,
            allow_manual_review=allow_manual_review,
        )
    except Exception:
        return None


def _safe_complexity(
    workbook: WorkbookSnapshot, cancellation_token: CancellationToken | None
) -> WorkbookComplexity | None:
    """A forecast, never a refusal -- always requests the uncapped numbers
    (``allow_complex_workbook=True``) so the analyst sees the real cost
    even for a workbook the actual run would otherwise refuse by default.
    """
    try:
        return assess_workbook_complexity(
            workbook,
            allow_complex_workbook=True,
            cancellation_token=cancellation_token,
        )
    except Exception:
        # A forecast failure disables the forecast only, never the scan.
        return None


def analyze_member(
    *,
    member_id: str,
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    baseline_hash: str,
    current_hash: str,
    baseline_xlsb_risk: XlsbRiskProfile | None = None,
    current_xlsb_risk: XlsbRiskProfile | None = None,
    allow_manual_review: bool = False,
    cancellation_token: CancellationToken | None = None,
) -> MemberSetupProfile:
    """Structural scan of one member's already-loaded baseline/current pair.

    Regions are detected independently per side. A ranked composite-key
    candidate is evaluated only when a same-named sheet exists on the OTHER
    side too and both sides detect the exact same NUMBER of regions (a
    conservative, safe pairing heuristic -- this is a preview scan, not the
    authoritative alignment the engine performs at run time; when sheets or
    region counts do not line up this simply, the candidate is left
    unevaluated -- never guessed).
    """
    baseline_sheets_by_name = {sheet.name: sheet for sheet in baseline.sheets}

    baseline_regions_by_sheet: dict[str, tuple[TableRegion, ...]] = {}
    baseline_profiles: list[SheetSetupProfile] = []
    for sheet in baseline.sheets[:MAX_SHEETS_PER_SIDE]:
        check_cancelled(cancellation_token)
        regions, detail = _detect_sheet_regions(sheet)
        baseline_regions_by_sheet[sheet.name] = regions
        baseline_profiles.append(
            SheetSetupProfile(
                sheet_name=sheet.name,
                hidden=sheet.visibility == "hidden",
                very_hidden=sheet.visibility == "veryHidden",
                regions=tuple(DetectedRegion(region=region) for region in regions),
                failure_detail=detail,
            )
        )

    current_profiles: list[SheetSetupProfile] = []
    for sheet in current.sheets[:MAX_SHEETS_PER_SIDE]:
        check_cancelled(cancellation_token)
        regions, detail = _detect_sheet_regions(sheet)
        base_sheet = baseline_sheets_by_name.get(sheet.name)
        base_regions = baseline_regions_by_sheet.get(sheet.name, ())
        detected: list[DetectedRegion] = []
        can_pair = base_sheet is not None and not detail and len(base_regions) == len(regions)
        for index, region in enumerate(regions):
            candidate = None
            if can_pair and base_sheet is not None:
                candidate = _detect_ranked_candidate(
                    base_sheet,
                    sheet,
                    base_regions[index],
                    region,
                    allow_manual_review=allow_manual_review,
                )
            detected.append(DetectedRegion(region=region, ranked_candidate=candidate))
        current_profiles.append(
            SheetSetupProfile(
                sheet_name=sheet.name,
                hidden=sheet.visibility == "hidden",
                very_hidden=sheet.visibility == "veryHidden",
                regions=tuple(detected),
                failure_detail=detail,
            )
        )

    return MemberSetupProfile(
        member_id=member_id,
        baseline_hash=baseline_hash,
        current_hash=current_hash,
        baseline_sheets=tuple(baseline_profiles),
        current_sheets=tuple(current_profiles),
        baseline_complexity=_safe_complexity(baseline, cancellation_token),
        current_complexity=_safe_complexity(current, cancellation_token),
        baseline_xlsb_risk=baseline_xlsb_risk,
        current_xlsb_risk=current_xlsb_risk,
    )


def _merge_complexity(
    target: WorkbookComplexity, source: WorkbookComplexity | None
) -> None:
    if source is None:
        return
    target.formula_count += source.formula_count
    target.lexical_formula_count += source.lexical_formula_count
    target.reference_operands += source.reference_operands
    target.resolved_range_cells += source.resolved_range_cells
    target.projected_concrete_edges += source.projected_concrete_edges
    target.interaction_rule_count += source.interaction_rule_count
    target.sampled_formulas += source.sampled_formulas
    for key, value in source.extras.items():
        target.extras[key] = target.extras.get(key, 0) + value


def _finalize_streaming_complexity(
    complexity: WorkbookComplexity,
) -> WorkbookComplexity:
    return finalize_workbook_complexity(
        complexity,
        source_name="setup-sidecar",
        allow_complex_workbook=True,
    )


def _sheet_complexity(
    sheet: SheetSnapshot, cancellation_token: CancellationToken | None
) -> WorkbookComplexity | None:
    workbook = WorkbookSnapshot(
        source_name="setup-sidecar",
        file_format="setup",
        formulas_available=any(cell.formula for cell in sheet.cells.values()),
        styles_available=False,
        formula_presence_available=True,
        sheets=[sheet],
    )
    return _safe_complexity(workbook, cancellation_token)


@dataclass(slots=True)
class StreamingMemberAnalyzer:
    """Accumulate structural profiles while source sheets stream to the sidecar."""

    member_id: str
    baseline_hash: str
    current_hash: str
    allow_manual_review: bool = False
    cancellation_token: CancellationToken | None = None
    baseline_profiles: list[SheetSetupProfile] = field(default_factory=list)
    current_profiles: list[SheetSetupProfile] = field(default_factory=list)
    baseline_regions: dict[str, tuple[TableRegion, ...]] = field(default_factory=dict)
    baseline_complexity: WorkbookComplexity = field(default_factory=WorkbookComplexity)
    current_complexity: WorkbookComplexity = field(default_factory=WorkbookComplexity)
    baseline_region_seconds: float = 0.0
    current_region_seconds: float = 0.0
    baseline_complexity_seconds: float = 0.0
    current_complexity_seconds: float = 0.0
    ranked_candidate_seconds: float = 0.0
    ranked_candidate_calls: int = 0

    def add_baseline(self, sheet: SheetSnapshot) -> SheetSetupProfile:
        check_cancelled(self.cancellation_token)
        region_started = time.perf_counter()
        regions, detail = _detect_sheet_regions(sheet)
        self.baseline_region_seconds += time.perf_counter() - region_started
        self.baseline_regions[sheet.name] = regions
        complexity_started = time.perf_counter()
        _merge_complexity(
            self.baseline_complexity,
            _sheet_complexity(sheet, self.cancellation_token),
        )
        self.baseline_complexity_seconds += time.perf_counter() - complexity_started
        profile = SheetSetupProfile(
            sheet_name=sheet.name,
            hidden=sheet.visibility == "hidden",
            very_hidden=sheet.visibility == "veryHidden",
            regions=tuple(DetectedRegion(region=region) for region in regions),
            failure_detail=detail,
        )
        self.baseline_profiles.append(profile)
        return profile

    def add_current(
        self,
        sheet: SheetSnapshot,
        baseline_sheet: SheetSnapshot | None,
    ) -> SheetSetupProfile:
        check_cancelled(self.cancellation_token)
        region_started = time.perf_counter()
        regions, detail = _detect_sheet_regions(sheet)
        self.current_region_seconds += time.perf_counter() - region_started
        complexity_started = time.perf_counter()
        _merge_complexity(
            self.current_complexity,
            _sheet_complexity(sheet, self.cancellation_token),
        )
        self.current_complexity_seconds += time.perf_counter() - complexity_started
        base_regions = self.baseline_regions.get(sheet.name, ())
        can_pair = (
            baseline_sheet is not None
            and not detail
            and len(base_regions) == len(regions)
        )
        detected: list[DetectedRegion] = []
        for index, region in enumerate(regions):
            candidate = None
            if can_pair and baseline_sheet is not None:
                candidate_started = time.perf_counter()
                candidate = _detect_ranked_candidate(
                    baseline_sheet,
                    sheet,
                    base_regions[index],
                    region,
                    allow_manual_review=self.allow_manual_review,
                )
                self.ranked_candidate_seconds += (
                    time.perf_counter() - candidate_started
                )
                if region.orientation == "block":
                    self.ranked_candidate_calls += 1
            detected.append(DetectedRegion(region=region, ranked_candidate=candidate))
        profile = SheetSetupProfile(
            sheet_name=sheet.name,
            hidden=sheet.visibility == "hidden",
            very_hidden=sheet.visibility == "veryHidden",
            regions=tuple(detected),
            failure_detail=detail,
        )
        self.current_profiles.append(profile)
        return profile

    def result(
        self,
        *,
        baseline_xlsb_risk: XlsbRiskProfile | None = None,
        current_xlsb_risk: XlsbRiskProfile | None = None,
    ) -> MemberSetupProfile:
        _finalize_streaming_complexity(self.baseline_complexity)
        _finalize_streaming_complexity(self.current_complexity)
        return MemberSetupProfile(
            member_id=self.member_id,
            baseline_hash=self.baseline_hash,
            current_hash=self.current_hash,
            baseline_sheets=tuple(self.baseline_profiles),
            current_sheets=tuple(self.current_profiles),
            baseline_complexity=self.baseline_complexity,
            current_complexity=self.current_complexity,
            baseline_xlsb_risk=baseline_xlsb_risk,
            current_xlsb_risk=current_xlsb_risk,
        )


def analyze_member_from_sidecar(
    *,
    store: SetupScanStore,
    session_key: str,
    input_generation: int,
    member_id: str,
    baseline_hash: str,
    current_hash: str,
    baseline_xlsb_risk: XlsbRiskProfile | None = None,
    current_xlsb_risk: XlsbRiskProfile | None = None,
    allow_manual_review: bool = False,
    cancellation_token: CancellationToken | None = None,
) -> MemberSetupProfile:
    """Analyze persisted sheets while bounding memory to one paired sheet."""
    baseline_inventory = store.list_sheets(
        session_key,
        member_id,
        "baseline",
        expected_generation=input_generation,
        expected_source_hash=baseline_hash,
    )
    current_inventory = store.list_sheets(
        session_key,
        member_id,
        "current",
        expected_generation=input_generation,
        expected_source_hash=current_hash,
    )
    baseline_by_name = {sheet.sheet_name: sheet for sheet in baseline_inventory}
    baseline_regions_by_name: dict[str, tuple[TableRegion, ...]] = {}
    baseline_profiles: list[SheetSetupProfile] = []
    baseline_complexity = WorkbookComplexity()
    current_complexity = WorkbookComplexity()

    for inventory in baseline_inventory[:MAX_SHEETS_PER_SIDE]:
        check_cancelled(cancellation_token)
        sheet = store.load_sheet(
            session_key,
            member_id,
            "baseline",
            inventory.sheet_name,
            expected_generation=input_generation,
            expected_source_hash=baseline_hash,
        )
        regions, detail = _detect_sheet_regions(sheet)
        baseline_regions_by_name[sheet.name] = regions
        _merge_complexity(
            baseline_complexity, _sheet_complexity(sheet, cancellation_token)
        )
        baseline_profiles.append(
            SheetSetupProfile(
                sheet_name=sheet.name,
                hidden=sheet.visibility == "hidden",
                very_hidden=sheet.visibility == "veryHidden",
                regions=tuple(DetectedRegion(region=region) for region in regions),
                failure_detail=detail,
            )
        )

    current_profiles: list[SheetSetupProfile] = []
    for inventory in current_inventory[:MAX_SHEETS_PER_SIDE]:
        check_cancelled(cancellation_token)
        current_sheet = store.load_sheet(
            session_key,
            member_id,
            "current",
            inventory.sheet_name,
            expected_generation=input_generation,
            expected_source_hash=current_hash,
        )
        regions, detail = _detect_sheet_regions(current_sheet)
        _merge_complexity(
            current_complexity,
            _sheet_complexity(current_sheet, cancellation_token),
        )
        base_inventory = baseline_by_name.get(current_sheet.name)
        base_regions = baseline_regions_by_name.get(current_sheet.name, ())
        baseline_sheet = None
        if base_inventory is not None and not detail and len(base_regions) == len(regions):
            baseline_sheet = store.load_sheet(
                session_key,
                member_id,
                "baseline",
                base_inventory.sheet_name,
                expected_generation=input_generation,
                expected_source_hash=baseline_hash,
            )
        detected: list[DetectedRegion] = []
        for index, region in enumerate(regions):
            candidate = None
            if baseline_sheet is not None:
                candidate = _detect_ranked_candidate(
                    baseline_sheet,
                    current_sheet,
                    base_regions[index],
                    region,
                    allow_manual_review=allow_manual_review,
                )
            detected.append(DetectedRegion(region=region, ranked_candidate=candidate))
        current_profiles.append(
            SheetSetupProfile(
                sheet_name=current_sheet.name,
                hidden=current_sheet.visibility == "hidden",
                very_hidden=current_sheet.visibility == "veryHidden",
                regions=tuple(detected),
                failure_detail=detail,
            )
        )

    _finalize_streaming_complexity(baseline_complexity)
    _finalize_streaming_complexity(current_complexity)
    return MemberSetupProfile(
        member_id=member_id,
        baseline_hash=baseline_hash,
        current_hash=current_hash,
        baseline_sheets=tuple(baseline_profiles),
        current_sheets=tuple(current_profiles),
        baseline_complexity=baseline_complexity,
        current_complexity=current_complexity,
        baseline_xlsb_risk=baseline_xlsb_risk,
        current_xlsb_risk=current_xlsb_risk,
    )
