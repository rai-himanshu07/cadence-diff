"""Severity triage: classify findings and finalize them for reporting.

Rules, in order:

1. Expected cadence growth is always ``expected`` — the whole point of the
   tool is that growth never drowns out real errors. Profile overrides do
   not escalate expected findings (use chart-window / refresh-range /
   region settings to change what counts as expected instead).
2. Profile per-class overrides apply to non-expected findings.
3. Otherwise the default class severity applies.

`triage` also orders findings deterministically (critical first) and
assigns stable ids — the finalized list is the contract every report
consumes.
"""

import datetime as dt
from collections import Counter

from qc_tool.config.profile import DeliverableProfile
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingEvidenceTag,
    FindingExpectedReason,
    FindingProvenance,
    FindingSubtype,
    FindingTemporalContext,
    Materiality,
    Severity,
)

DEFAULT_SEVERITIES: dict[FindingClass, Severity] = {
    # Historical data integrity: critical.
    FindingClass.VALUE_CHANGED: Severity.CRITICAL,
    FindingClass.FORMULA_ERROR: Severity.CRITICAL,
    FindingClass.FORMULA_HARDCODED: Severity.CRITICAL,
    FindingClass.FORMULA_REMOVED: Severity.CRITICAL,
    FindingClass.FORMULA_MISSING: Severity.CRITICAL,
    FindingClass.CIRCULAR_REFERENCE: Severity.CRITICAL,
    FindingClass.EXTERNAL_LINK: Severity.CRITICAL,
    FindingClass.ACTIVE_CONTENT: Severity.CRITICAL,
    #: Macro logic can rewrite any figure in the deliverable.
    FindingClass.VBA_MODULE_CHANGED: Severity.CRITICAL,
    #: Query and connection edits change where the numbers came from.
    FindingClass.POWER_QUERY_CHANGED: Severity.CRITICAL,
    FindingClass.CONNECTION_CHANGED: Severity.CRITICAL,
    FindingClass.EXTERNAL_CONNECTION: Severity.CRITICAL,
    FindingClass.NAMED_RANGE_INVALID: Severity.CRITICAL,
    FindingClass.CHART_REFERENCE_INVALID: Severity.CRITICAL,
    FindingClass.PIVOT_SOURCE_INVALID: Severity.CRITICAL,
    FindingClass.PPT_DRAFT_TOKEN: Severity.CRITICAL,
    FindingClass.PPT_REQUIRED_SLIDE_MISSING: Severity.CRITICAL,
    FindingClass.REQUIRED_VALUE_MISSING: Severity.CRITICAL,
    FindingClass.DUPLICATE_KEY: Severity.CRITICAL,
    FindingClass.NUMERIC_BOUND_VIOLATION: Severity.CRITICAL,
    FindingClass.TIE_OUT_MISMATCH: Severity.CRITICAL,
    FindingClass.ROW_DELETED: Severity.CRITICAL,
    FindingClass.COLUMN_DELETED: Severity.CRITICAL,
    FindingClass.SHEET_REMOVED: Severity.CRITICAL,
    FindingClass.WORKBOOK_REMOVED: Severity.CRITICAL,
    #: An in-place key change on a constant key rewrites history identity.
    FindingClass.ROW_KEY_CHANGED: Severity.CRITICAL,
    FindingClass.COLUMN_KEY_CHANGED: Severity.CRITICAL,
    FindingClass.CROSSCHECK_MISMATCH: Severity.CRITICAL,
    FindingClass.PACKAGE_PERIOD_MISMATCH: Severity.CRITICAL,
    # Formula and structure drift: warning.
    FindingClass.FORMULA_LOGIC_CHANGED: Severity.WARNING,
    FindingClass.FORMULA_NOT_EXTENDED: Severity.WARNING,
    FindingClass.FORMULA_INCONSISTENT: Severity.WARNING,
    FindingClass.FORMULA_CACHE_MISSING: Severity.WARNING,
    FindingClass.PERIOD_DUPLICATE: Severity.WARNING,
    FindingClass.PERIOD_OUT_OF_ORDER: Severity.WARNING,
    FindingClass.PERIOD_GAP: Severity.WARNING,
    FindingClass.CALCULATION_MODE: Severity.WARNING,
    FindingClass.CHART_LENGTH_MISMATCH: Severity.WARNING,
    FindingClass.PPT_EMPTY_SLIDE: Severity.WARNING,
    FindingClass.PPT_DUPLICATE_TITLE: Severity.WARNING,
    FindingClass.PPT_PERIOD_INCONSISTENT: Severity.WARNING,
    FindingClass.PPT_TABLE_BLANK: Severity.WARNING,
    FindingClass.PPT_CHART_LENGTH_MISMATCH: Severity.WARNING,
    FindingClass.PPT_CHART_VALUE_MISSING: Severity.WARNING,
    FindingClass.PPT_REPEATED_CLAIM_MISMATCH: Severity.WARNING,
    FindingClass.CONTROL_INVALID: Severity.WARNING,
    FindingClass.WAIVER_EXPIRED: Severity.WARNING,
    FindingClass.NUMBER_FORMAT_CHANGED: Severity.WARNING,
    FindingClass.NAMED_RANGE_CHANGED: Severity.WARNING,
    FindingClass.TABLE_STRUCTURE_CHANGED: Severity.WARNING,
    FindingClass.DATA_VALIDATION_CHANGED: Severity.WARNING,
    FindingClass.CONDITIONAL_FORMAT_CHANGED: Severity.WARNING,
    FindingClass.CHART_STRUCTURE_CHANGED: Severity.WARNING,
    FindingClass.CHART_PLOT_CHANGED: Severity.WARNING,
    FindingClass.CHART_SERIES_CHANGED: Severity.WARNING,
    FindingClass.CHART_AXIS_CHANGED: Severity.WARNING,
    FindingClass.CHART_LEGEND_CHANGED: Severity.WARNING,
    FindingClass.CHART_LABELS_CHANGED: Severity.WARNING,
    FindingClass.PIVOT_SOURCE_CHANGED: Severity.WARNING,
    FindingClass.HIDDEN_CHANGED: Severity.WARNING,
    FindingClass.ROW_INSERTED: Severity.WARNING,
    FindingClass.COLUMN_INSERTED: Severity.WARNING,
    FindingClass.REGION_UNPAIRED: Severity.WARNING,
    FindingClass.ALIGNMENT_LOW_CONFIDENCE: Severity.WARNING,
    FindingClass.FINDINGS_CAPPED: Severity.WARNING,
    FindingClass.SLIDE_REMOVED: Severity.WARNING,
    FindingClass.SLIDE_TEXT_CHANGED: Severity.WARNING,
    FindingClass.TABLE_VALUE_CHANGED: Severity.WARNING,
    FindingClass.CHART_VALUE_CHANGED: Severity.WARNING,
    FindingClass.PPT_TABLE_STRUCTURE_CHANGED: Severity.WARNING,
    FindingClass.PPT_CHART_STRUCTURE_CHANGED: Severity.WARNING,
    FindingClass.PPT_CHART_PLOT_CHANGED: Severity.WARNING,
    FindingClass.PPT_CHART_SERIES_CHANGED: Severity.WARNING,
    FindingClass.PPT_CHART_AXIS_CHANGED: Severity.WARNING,
    FindingClass.PPT_CHART_LEGEND_CHANGED: Severity.WARNING,
    FindingClass.PPT_CHART_LABELS_CHANGED: Severity.WARNING,
    FindingClass.PPT_MEDIA_CHANGED: Severity.WARNING,
    # Presentation-only and additive events: info.
    FindingClass.CELL_COMMENT_CHANGED: Severity.INFO,
    FindingClass.STYLE_CHANGED: Severity.INFO,
    FindingClass.CHART_GEOMETRY_CHANGED: Severity.INFO,
    FindingClass.PPT_SHAPE_GEOMETRY_CHANGED: Severity.INFO,
    FindingClass.SHEET_ADDED: Severity.INFO,
    #: A confirmed, analyst-recognized logical rename -- not a surprise.
    FindingClass.SHEET_RENAMED: Severity.INFO,
    FindingClass.WORKBOOK_ADDED: Severity.INFO,
    FindingClass.SLIDE_ADDED: Severity.INFO,
    FindingClass.SLIDE_REORDERED: Severity.INFO,
    FindingClass.CROSSCHECK_UNRESOLVED: Severity.INFO,
    FindingClass.HIDDEN_CONTENT: Severity.INFO,
    # Geometry alone does not prove cadence growth.
    FindingClass.ROW_GROWTH: Severity.WARNING,
    FindingClass.COLUMN_GROWTH: Severity.WARNING,
}

_SEVERITY_RANK = {
    Severity.CRITICAL: 0,
    Severity.WARNING: 1,
    Severity.INFO: 2,
    Severity.EXPECTED: 3,
}

#: Default severity for non-material numeric tiers. Material tiers fall
#: through to the per-class path so history integrity keeps its default.
MATERIALITY_SEVERITIES: dict[Materiality, Severity] = {
    Materiality.NOISE: Severity.INFO,
    Materiality.WITHIN_TOLERANCE: Severity.INFO,
    # Legacy history only; new findings carry temporal_context separately.
    Materiality.RECENT_RESTATEMENT: Severity.WARNING,
}

_RECENT_CONTEXTS = frozenset(
    {
        FindingTemporalContext.CURRENT_PERIOD,
        FindingTemporalContext.RECENT_WINDOW,
    }
)

#: Pre-existing conditions proven identical in the baseline are not this
#: cycle's regressions; they inform rather than alarm.
_INHERITED_DEMOTED_CLASSES = frozenset({FindingClass.FORMULA_INCONSISTENT})

#: Subtype-specific defaults: a formula-derived label re-resolving is driver
#: churn, not a historical-row deletion; the oldest row leaving a detected
#: rolling window is that window's design, not a historical deletion.
_SUBTYPE_SEVERITIES: dict[tuple[FindingClass, FindingSubtype], Severity] = {
    (
        FindingClass.ROW_KEY_CHANGED,
        FindingSubtype.AXIS_KEY_DERIVED_LABEL,
    ): Severity.WARNING,
    (
        FindingClass.COLUMN_KEY_CHANGED,
        FindingSubtype.AXIS_KEY_DERIVED_LABEL,
    ): Severity.WARNING,
    (
        FindingClass.ROW_DELETED,
        FindingSubtype.AXIS_ROLLING_TURNOVER,
    ): Severity.WARNING,
    (
        FindingClass.COLUMN_DELETED,
        FindingSubtype.AXIS_ROLLING_TURNOVER,
    ): Severity.WARNING,
}

_REGRESSION_ERROR_PROVENANCE = frozenset(
    {FindingProvenance.NEW, FindingProvenance.CHANGED}
)
_SYSTEMATIC_ERROR_EVIDENCE = frozenset(
    {
        FindingEvidenceTag.CONCENTRATED_POPULATION,
        FindingEvidenceTag.CONTIGUOUS_POPULATION,
    }
)


def _formula_error_severity(finding: Finding) -> Severity | None:
    if finding.finding_class is not FindingClass.FORMULA_ERROR:
        return None
    evidence = finding.evidence_tags
    if FindingEvidenceTag.STRUCTURAL_ERROR in evidence:
        return Severity.CRITICAL
    if finding.provenance in _REGRESSION_ERROR_PROVENANCE:
        return Severity.CRITICAL
    if (
        finding.provenance is FindingProvenance.INHERITED
        and FindingEvidenceTag.EXPLICIT_NA in evidence
        and FindingEvidenceTag.FORMULA_TEXT in evidence
    ):
        return Severity.INFO
    if (
        FindingEvidenceTag.FORMULA_PRESENCE in evidence
        and not _SYSTEMATIC_ERROR_EVIDENCE.isdisjoint(evidence)
    ):
        return Severity.WARNING
    return Severity.CRITICAL


def assign_severity(finding: Finding, profile: DeliverableProfile | None = None) -> Severity:
    """Pure function of typed attributes only -- the shared severity
    projection producer-time candidates apply before grouping (both this
    and `matches_waiver` read plain attributes, never other findings).
    """
    if finding.expected_reason is not None:
        return Severity.EXPECTED
    if finding.expected_growth:  # legacy evidence without a typed reason
        return Severity.EXPECTED
    error_severity = _formula_error_severity(finding)
    if error_severity is not None:
        return error_severity
    tier = finding.materiality
    if tier is not None:
        tier_overrides = profile.materiality_severity if profile is not None else {}
        mapped = tier_overrides.get(tier, MATERIALITY_SEVERITIES.get(tier))
        if mapped is not None:
            return mapped
    if finding.materiality is Materiality.MATERIAL and (
        finding.temporal_context in _RECENT_CONTEXTS
    ):
        return Severity.WARNING
    if (
        finding.provenance is FindingProvenance.INHERITED
        and finding.finding_class in _INHERITED_DEMOTED_CLASSES
    ):
        return Severity.INFO
    overrides = profile.severity if profile is not None else {}
    if finding.finding_class in overrides:
        return overrides[finding.finding_class]
    if finding.subtype is not None:
        subtype_default = _SUBTYPE_SEVERITIES.get(
            (finding.finding_class, finding.subtype)
        )
        if subtype_default is not None:
            return subtype_default
    return DEFAULT_SEVERITIES.get(finding.finding_class, Severity.WARNING)


def triage(
    findings: list[Finding], profile: DeliverableProfile | None = None
) -> list[Finding]:
    """Assign severities, order deterministically, and number the findings."""
    today = dt.date.today()
    findings.extend(expired_waiver_findings(profile, today))
    assign_severities(findings, profile, today=today)

    candidates = [
        (finding, key)
        for finding in findings
        if (key := root_cause_candidate_key(finding)) is not None
    ]
    root_counts = Counter(key for _, key in candidates)
    for finding, key in candidates:
        if root_counts[key] > 1:
            finding.root_cause_key = key
    ordered = sorted(findings, key=triage_sort_key)
    for index, finding in enumerate(ordered, start=1):
        finding.finding_id = f"F{index:04d}"
    return ordered


def expired_waiver_findings(
    profile: DeliverableProfile | None, today: dt.date
) -> list[Finding]:
    """One finding per profile waiver whose expiry date has passed."""
    waivers = profile.waivers if profile is not None else []
    return [
        Finding(
            artifact="profile",
            artifact_member=waiver.member,
            finding_class=FindingClass.WAIVER_EXPIRED,
            sheet=waiver.sheet,
            slide=waiver.slide,
            location=waiver.location,
            element=waiver.element or waiver.finding_class.value,
            current_value=waiver.expires.isoformat(),
            message=(
                f"waiver for {waiver.finding_class.value} expired on "
                f"{waiver.expires.isoformat()}: {waiver.reason}"
            ),
        )
        for waiver in waivers
        if waiver.expires < today
    ]


def matches_waiver(finding: Finding, waiver) -> bool:
    """Public so producer-time candidates can share this exact check."""
    return (
        finding.finding_class is waiver.finding_class
        and finding.artifact_member == waiver.member
        and (waiver.sheet is None or finding.sheet == waiver.sheet)
        and (waiver.slide is None or finding.slide == waiver.slide)
        and (waiver.location is None or finding.location == waiver.location)
        and (waiver.element is None or finding.element == waiver.element)
    )


def assign_severities(
    findings: list[Finding],
    profile: DeliverableProfile | None,
    *,
    today: dt.date,
) -> None:
    """Apply active waivers and severity rules to each finding in place."""
    waivers = profile.waivers if profile is not None else []
    for finding in findings:
        active = next(
            (
                waiver
                for waiver in waivers
                if waiver.expires >= today and matches_waiver(finding, waiver)
            ),
            None,
        )
        if active is not None:
            finding.mark_expected(FindingExpectedReason.WAIVER)
            finding.severity = Severity.EXPECTED
            finding.waiver_reason = active.reason
            finding.waiver_expires = active.expires.isoformat()
        else:
            finding.severity = assign_severity(finding, profile)


_ROOT_CAUSE_CLASSES = frozenset(
    {
        FindingClass.FORMULA_ERROR,
        FindingClass.FORMULA_HARDCODED,
        FindingClass.FORMULA_REMOVED,
        FindingClass.FORMULA_MISSING,
        FindingClass.FORMULA_NOT_EXTENDED,
        FindingClass.FORMULA_LOGIC_CHANGED,
        FindingClass.FORMULA_INCONSISTENT,
    }
)


def root_cause_candidate_key(finding: Finding) -> str | None:
    """Formula co-location key; shared keys mark a common root cause."""
    if finding.finding_class not in _ROOT_CAUSE_CLASSES:
        return None
    if not (finding.location or finding.element):
        return None
    return (
        f"{finding.artifact}"
        + (
            f":{finding.artifact_member}"
            if finding.artifact_member != "primary"
            else ""
        )
        + ":"
        f"{finding.sheet or finding.slide}:"
        f"{finding.location or finding.element}"
    )


def triage_sort_key(finding: Finding) -> tuple[int, str, str, str, str, str, str, str]:
    """The global deterministic ordering every report and store relies on."""
    return (
        _SEVERITY_RANK[finding.severity or Severity.WARNING],
        finding.artifact,
        finding.artifact_member,
        finding.sheet or finding.slide or "",
        finding.location or "",
        finding.element or "",
        finding.finding_class.value,
        finding.message,
    )
