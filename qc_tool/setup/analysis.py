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

from qc_tool.excel.complexity import WorkbookComplexity, assess_workbook_complexity
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
                try:
                    candidate = detect_ranked_table_candidate(
                        base_sheet,
                        sheet,
                        base_regions[index],
                        region,
                        allow_manual_review=allow_manual_review,
                    )
                except Exception:
                    # Isolate one region's candidate-detection failure.
                    candidate = None
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
