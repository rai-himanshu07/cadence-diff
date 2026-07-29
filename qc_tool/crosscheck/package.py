"""Current Excel plus current PowerPoint final-package reconciliation."""

from dataclasses import dataclass, field

from qc_tool.config.profile import CrosscheckProfile
from qc_tool.coverage import CoverageItem, CoverageState, MappingCoverage
from qc_tool.crosscheck.trace import (
    MappingSuggestion,
    SuggestedSource,
    extract_deck_figures,
    mapping_identity,
    occurrence_identity,
    suggest_sources,
    verify_mappings,
)
from qc_tool.excel.periods import Period, is_period_after, parse_period
from qc_tool.findings import Finding, FindingClass
from qc_tool.io.model import WorkbookSnapshot
from qc_tool.ppt.extract import DeckSnapshot
from qc_tool.ppt.preflight import contextual_periods


@dataclass(slots=True)
class PackageQCResult:
    findings: list[Finding] = field(default_factory=list)
    coverage: list[CoverageItem] = field(default_factory=list)
    mapping_coverage: MappingCoverage = field(default_factory=MappingCoverage)
    suggestions: list[MappingSuggestion] = field(default_factory=list)


def _latest_workbook_periods(workbook: WorkbookSnapshot) -> dict[str, Period]:
    latest: dict[str, Period] = {}
    for sheet in workbook.sheets:
        for cell in sheet.cells.values():
            period = parse_period(cell.value)
            if period is None:
                continue
            previous = latest.get(period.kind)
            if previous is None or is_period_after(period, previous):
                latest[period.kind] = period
    return latest


def _period_reconciliation(
    workbook: WorkbookSnapshot, deck: DeckSnapshot
) -> list[Finding]:
    findings: list[Finding] = []
    latest = _latest_workbook_periods(workbook)
    for (kind, sort_key), sources in contextual_periods(deck).items():
        workbook_period = latest.get(kind)
        if workbook_period is None or workbook_period.sort_key == sort_key:
            continue
        findings.append(
            Finding(
                artifact="crosscheck",
                finding_class=FindingClass.PACKAGE_PERIOD_MISMATCH,
                element="reporting period",
                baseline_value=workbook_period.label,
                current_value="; ".join(sorted(sources)),
                message=(
                    "PowerPoint contextual reporting period does not match the "
                    f"latest {kind} period in Excel ({workbook_period.label})"
                ),
            )
        )
    return findings


def reconcile_package(
    workbook: WorkbookSnapshot, deck: DeckSnapshot, profile: CrosscheckProfile
) -> PackageQCResult:
    result = PackageQCResult()
    occurrences = extract_deck_figures(deck)
    mapped_identities = {mapping_identity(mapping) for mapping in profile.mappings}
    mapped_occurrences = [
        occurrence
        for occurrence in occurrences
        if occurrence_identity(occurrence) in mapped_identities
    ]
    unmapped_occurrences = [
        occurrence
        for occurrence in occurrences
        if occurrence_identity(occurrence) not in mapped_identities
    ]

    verified = verify_mappings(deck, workbook, profile)
    result.findings.extend(verified.findings)
    mismatched = sum(
        finding.finding_class is FindingClass.CROSSCHECK_MISMATCH
        for finding in verified.findings
    )
    unresolved = sum(
        finding.finding_class is FindingClass.CROSSCHECK_UNRESOLVED
        for finding in verified.findings
    )
    result.mapping_coverage = MappingCoverage(
        eligible=len(occurrences),
        mapped=len(mapped_occurrences),
        verified=len(verified.verified),
        mismatched=mismatched,
        unresolved=unresolved,
        unmapped=len(unmapped_occurrences),
    )
    for occurrence in unmapped_occurrences:
        candidates = suggest_sources(
            occurrence, workbook, limit=profile.max_candidates
        )
        result.suggestions.append(
            MappingSuggestion(
                slide=occurrence.slide,
                line=occurrence.line,
                line_skeleton=occurrence.line_skeleton,
                figure_index=occurrence.figure_index,
                figure_raw=occurrence.figure.raw,
                candidates=[SuggestedSource.from_candidate(candidate) for candidate in candidates],
            )
        )

    period_findings = _period_reconciliation(workbook, deck)
    result.findings.extend(period_findings)
    mapping = result.mapping_coverage
    result.coverage.extend(
        [
            CoverageItem(
                check_id="excel-ppt-crosscheck",
                label="Excel to PowerPoint figure mappings",
                artifact="package",
                state=(
                    CoverageState.CHECKED
                    if deck.charts_available
                    else CoverageState.DEGRADED
                ),
                findings=len(verified.findings),
                detail=(
                    f"{mapping.eligible} eligible; {mapping.mapped} mapped; "
                    f"{mapping.verified} verified; {mapping.mismatched} mismatched; "
                    f"{mapping.unresolved} unresolved; {mapping.unmapped} unmapped"
                    + (
                        ""
                        if deck.charts_available
                        else "; visible native chart labels unavailable"
                    )
                ),
            ),
            CoverageItem(
                check_id="package-period-reconciliation",
                label="Excel and PowerPoint reporting-period reconciliation",
                artifact="package",
                state=CoverageState.CHECKED,
                findings=len(period_findings),
            ),
        ]
    )
    return result
