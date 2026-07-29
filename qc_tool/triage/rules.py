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
from qc_tool.findings import Finding, FindingClass, Severity

DEFAULT_SEVERITIES: dict[FindingClass, Severity] = {
    # Historical data integrity: critical.
    FindingClass.VALUE_CHANGED: Severity.CRITICAL,
    FindingClass.FORMULA_ERROR: Severity.CRITICAL,
    FindingClass.FORMULA_HARDCODED: Severity.CRITICAL,
    FindingClass.FORMULA_REMOVED: Severity.CRITICAL,
    FindingClass.FORMULA_MISSING: Severity.CRITICAL,
    FindingClass.EXTERNAL_LINK: Severity.CRITICAL,
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
    # Presentation-only and additive events: info.
    FindingClass.STYLE_CHANGED: Severity.INFO,
    FindingClass.CHART_GEOMETRY_CHANGED: Severity.INFO,
    FindingClass.PPT_SHAPE_GEOMETRY_CHANGED: Severity.INFO,
    FindingClass.SHEET_ADDED: Severity.INFO,
    FindingClass.SLIDE_ADDED: Severity.INFO,
    FindingClass.SLIDE_REORDERED: Severity.INFO,
    FindingClass.CROSSCHECK_UNRESOLVED: Severity.INFO,
    FindingClass.HIDDEN_CONTENT: Severity.INFO,
    # Growth classes only ever occur with expected_growth=True.
    FindingClass.ROW_GROWTH: Severity.EXPECTED,
    FindingClass.COLUMN_GROWTH: Severity.EXPECTED,
}

_SEVERITY_RANK = {
    Severity.CRITICAL: 0,
    Severity.WARNING: 1,
    Severity.INFO: 2,
    Severity.EXPECTED: 3,
}


def assign_severity(finding: Finding, profile: DeliverableProfile | None = None) -> Severity:
    if finding.expected_growth:
        return Severity.EXPECTED
    overrides = profile.severity if profile is not None else {}
    return overrides.get(
        finding.finding_class,
        DEFAULT_SEVERITIES.get(finding.finding_class, Severity.WARNING),
    )


def triage(
    findings: list[Finding], profile: DeliverableProfile | None = None
) -> list[Finding]:
    """Assign severities, order deterministically, and number the findings."""
    today = dt.date.today()
    waivers = profile.waivers if profile is not None else []
    for waiver in waivers:
        if waiver.expires >= today:
            continue
        findings.append(
            Finding(
                artifact="profile",
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
        )

    def matches_waiver(finding: Finding, waiver) -> bool:
        return (
            finding.finding_class is waiver.finding_class
            and (waiver.sheet is None or finding.sheet == waiver.sheet)
            and (waiver.slide is None or finding.slide == waiver.slide)
            and (waiver.location is None or finding.location == waiver.location)
            and (waiver.element is None or finding.element == waiver.element)
        )

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
            finding.severity = Severity.EXPECTED
            finding.waiver_reason = active.reason
            finding.waiver_expires = active.expires.isoformat()
        else:
            finding.severity = assign_severity(finding, profile)

    formula_classes = {
        FindingClass.FORMULA_ERROR,
        FindingClass.FORMULA_HARDCODED,
        FindingClass.FORMULA_REMOVED,
        FindingClass.FORMULA_MISSING,
        FindingClass.FORMULA_NOT_EXTENDED,
        FindingClass.FORMULA_LOGIC_CHANGED,
        FindingClass.FORMULA_INCONSISTENT,
    }
    candidates = [
        (
            finding,
            f"{finding.artifact}:{finding.sheet or finding.slide}:"
            f"{finding.location or finding.element}",
        )
        for finding in findings
        if finding.finding_class in formula_classes
        and (finding.location or finding.element)
    ]
    root_counts = Counter(key for _, key in candidates)
    for finding, key in candidates:
        if root_counts[key] > 1:
            finding.root_cause_key = key
    ordered = sorted(
        findings,
        key=lambda f: (
            _SEVERITY_RANK[f.severity or Severity.WARNING],
            f.artifact,
            f.sheet or f.slide or "",
            f.location or "",
            f.element or "",
            f.finding_class.value,
            f.message,
        ),
    )
    for index, finding in enumerate(ordered, start=1):
        finding.finding_id = f"F{index:04d}"
    return ordered
