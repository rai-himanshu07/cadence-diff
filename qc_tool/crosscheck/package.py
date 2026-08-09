"""Current Excel plus current PowerPoint final-package reconciliation."""

from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field

from qc_tool.config.profile import CrosscheckProfile
from qc_tool.coverage import CoverageItem, CoverageState, MappingCoverage
from qc_tool.crosscheck.claims import claim_population
from qc_tool.crosscheck.trace import (
    MappingSuggestion,
    SourceCandidate,
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


def _latest_workbook_periods(
    workbooks: Iterable[WorkbookSnapshot],
) -> dict[str, Period]:
    latest: dict[str, Period] = {}
    for workbook in workbooks:
        _update_latest_workbook_periods(latest, workbook)
    return latest


def _update_latest_workbook_periods(
    latest: dict[str, Period],
    workbook: WorkbookSnapshot,
) -> None:
    for sheet in workbook.sheets:
        for cell in sheet.cells.values():
            period = parse_period(cell.value)
            if period is None:
                continue
            previous = latest.get(period.kind)
            if previous is None or is_period_after(period, previous):
                latest[period.kind] = period


def _period_reconciliation_from_latest(
    latest: Mapping[str, Period],
    deck: DeckSnapshot,
) -> list[Finding]:
    findings: list[Finding] = []
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


def _period_reconciliation(
    workbooks: Iterable[WorkbookSnapshot],
    deck: DeckSnapshot,
) -> list[Finding]:
    latest = _latest_workbook_periods(workbooks)
    return _period_reconciliation_from_latest(latest, deck)


def reconcile_package(
    workbook: WorkbookSnapshot, deck: DeckSnapshot, profile: CrosscheckProfile
) -> PackageQCResult:
    result = PackageQCResult()
    occurrences = extract_deck_figures(deck)
    population = claim_population(deck, occurrences)
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
        unavailable=population.unreadable_shapes,
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

    period_findings = _period_reconciliation((workbook,), deck)
    result.findings.extend(period_findings)
    mapping = result.mapping_coverage
    result.coverage.extend(
        [
            CoverageItem(
                check_id="excel-ppt-crosscheck",
                label="Excel to PowerPoint figure mappings",
                artifact="package",
                state=(
                    CoverageState.DEGRADED
                    if not deck.charts_available or not population.complete
                    else CoverageState.CHECKED
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
                    + (
                        ""
                        if population.complete
                        else (
                            f"; {mapping.unavailable} rasterized or embedded "
                            "surface(s) carry claims that cannot be read on "
                            "slide(s) "
                            + ", ".join(
                                str(index) for index in population.unreadable_slides
                            )
                        )
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


def reconcile_multi_package(
    workbooks: Mapping[str, WorkbookSnapshot],
    deck: DeckSnapshot,
    profile: CrosscheckProfile,
) -> PackageQCResult:
    """Reconcile one deck once against stable-ID current workbook members."""
    reconciler = MultiPackageReconciler(deck, profile, workbooks)
    for member_id, workbook in sorted(workbooks.items()):
        reconciler.observe_member(member_id, workbook)
    return reconciler.finish()


def _candidate_sort_key(candidate: SourceCandidate) -> tuple[object, ...]:
    return (
        not candidate.display_match,
        -candidate.label_score,
        candidate.rel_diff,
        candidate.source_member,
        candidate.sheet,
        candidate.cell,
    )


class MultiPackageReconciler:
    """Incremental whole-package reconciliation over one workbook at a time."""

    def __init__(
        self,
        deck: DeckSnapshot,
        profile: CrosscheckProfile,
        member_ids: Iterable[str],
    ) -> None:
        self._deck = deck
        self._profile = profile
        self._member_ids = frozenset(member_ids)
        self._observed: set[str] = set()
        self._occurrences = extract_deck_figures(deck)
        self._population = claim_population(deck, self._occurrences)
        mapped_identities = {
            mapping_identity(mapping) for mapping in profile.mappings
        }
        self._mapped_occurrences = [
            occurrence
            for occurrence in self._occurrences
            if occurrence_identity(occurrence) in mapped_identities
        ]
        self._unmapped_occurrences = [
            occurrence
            for occurrence in self._occurrences
            if occurrence_identity(occurrence) not in mapped_identities
        ]
        self._candidates: list[list[SourceCandidate]] = [
            [] for _occurrence in self._unmapped_occurrences
        ]
        self._latest_periods: dict[str, Period] = {}
        self._verified_count = 0
        self._result = PackageQCResult()
        for member_id in sorted(
            {mapping.source_member for mapping in profile.mappings}
            - self._member_ids
        ):
            for mapping in profile.mappings:
                if mapping.source_member != member_id:
                    continue
                self._result.findings.append(
                    Finding(
                        artifact="crosscheck",
                        artifact_member=member_id,
                        finding_class=FindingClass.CROSSCHECK_UNRESOLVED,
                        slide=mapping.slide,
                        element=mapping.label or mapping.line_skeleton,
                        message=(
                            f"Mapped source workbook member {member_id!r} "
                            "is unavailable in the current package"
                        ),
                    )
                )

    def observe_member(
        self,
        member_id: str,
        workbook: WorkbookSnapshot,
    ) -> None:
        if member_id not in self._member_ids:
            raise ValueError(f"unexpected workbook member {member_id!r}")
        if member_id in self._observed:
            raise ValueError(f"workbook member {member_id!r} was observed twice")
        self._observed.add(member_id)
        mappings = [
            mapping
            for mapping in self._profile.mappings
            if mapping.source_member == member_id
        ]
        if mappings:
            member_profile = CrosscheckProfile(
                mappings=mappings,
                max_candidates=self._profile.max_candidates,
            )
            verified = verify_mappings(self._deck, workbook, member_profile)
            self._result.findings.extend(verified.findings)
            self._verified_count += len(verified.verified)
        _update_latest_workbook_periods(self._latest_periods, workbook)
        for index, occurrence in enumerate(self._unmapped_occurrences):
            candidates = [
                *self._candidates[index],
                *suggest_sources(
                    occurrence,
                    workbook,
                    limit=self._profile.max_candidates,
                    source_member=member_id,
                ),
            ]
            candidates.sort(key=_candidate_sort_key)
            self._candidates[index] = candidates[: self._profile.max_candidates]

    def finish(self) -> PackageQCResult:
        missing = sorted(self._member_ids - self._observed)
        if missing:
            raise ValueError(
                "package reconciliation missed workbook member(s): "
                + ", ".join(missing)
            )
        mismatched = sum(
            finding.finding_class is FindingClass.CROSSCHECK_MISMATCH
            for finding in self._result.findings
        )
        unresolved = sum(
            finding.finding_class is FindingClass.CROSSCHECK_UNRESOLVED
            for finding in self._result.findings
        )
        self._result.mapping_coverage = MappingCoverage(
            eligible=len(self._occurrences),
            mapped=len(self._mapped_occurrences),
            verified=self._verified_count,
            mismatched=mismatched,
            unresolved=unresolved,
            unmapped=len(self._unmapped_occurrences),
            unavailable=self._population.unreadable_shapes,
        )
        for occurrence, candidates in zip(
            self._unmapped_occurrences,
            self._candidates,
            strict=True,
        ):
            self._result.suggestions.append(
                MappingSuggestion(
                    slide=occurrence.slide,
                    line=occurrence.line,
                    line_skeleton=occurrence.line_skeleton,
                    figure_index=occurrence.figure_index,
                    figure_raw=occurrence.figure.raw,
                    candidates=[
                        SuggestedSource.from_candidate(candidate)
                        for candidate in candidates
                    ],
                )
            )
        period_findings = _period_reconciliation_from_latest(
            self._latest_periods,
            self._deck,
        )
        self._result.findings.extend(period_findings)
        mapping = self._result.mapping_coverage
        self._result.coverage.extend(
            [
                CoverageItem(
                    check_id="excel-ppt-crosscheck",
                    label="Excel to PowerPoint figure mappings",
                    artifact="package",
                    state=(
                        CoverageState.DEGRADED
                        if not self._deck.charts_available
                        or not self._population.complete
                        else CoverageState.CHECKED
                    ),
                    findings=mismatched + unresolved,
                    detail=(
                        f"{mapping.eligible} eligible; {mapping.mapped} mapped; "
                        f"{mapping.verified} verified; "
                        f"{mapping.mismatched} mismatched; "
                        f"{mapping.unresolved} unresolved; "
                        f"{mapping.unmapped} unmapped across "
                        f"{len(self._member_ids)} workbook members"
                        + (
                            ""
                            if self._population.complete
                            else (
                                f"; {mapping.unavailable} opaque surfaces unavailable"
                            )
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
        return self._result
