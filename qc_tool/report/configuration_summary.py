"""Value-free summary of the exact resolved input configuration."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass

from qc_tool.config.resolved_input import ResolvedInputConfigurationV1


@dataclass(frozen=True, slots=True)
class ResolvedConfigurationSummary:
    digest: str
    member_count: int
    sheet_count: int
    region_count: int
    selector_count: int
    coverage_counts: tuple[tuple[str, int], ...]
    excluded_regions: int
    degraded_regions: int
    acknowledgement_count: int
    override_count: int

    @property
    def coverage_text(self) -> str:
        return ", ".join(
            f"{coverage}: {count}" for coverage, count in self.coverage_counts
        ) or "none"


def summarize_resolved_configuration(
    resolved: ResolvedInputConfigurationV1 | None,
    *,
    digest: str = "",
) -> ResolvedConfigurationSummary | None:
    """Summarize structural decisions without values, formulas, or paths."""
    if resolved is None:
        return None
    sheets = [sheet for member in resolved.members for sheet in member.sheets]
    regions = [region for sheet in sheets for region in sheet.regions]
    selectors = [selector for sheet in sheets for selector in sheet.selectors]
    coverage = Counter(region.coverage for region in regions)
    return ResolvedConfigurationSummary(
        digest=digest or resolved.canonical_sha256(),
        member_count=len(resolved.members),
        sheet_count=len(sheets),
        region_count=len(regions),
        selector_count=len(selectors),
        coverage_counts=tuple(sorted(coverage.items())),
        excluded_regions=sum(region.coverage == "excluded" for region in regions),
        degraded_regions=(
            sum(region.coverage == "degraded_acknowledged" for region in regions)
            + len(resolved.degradations_accepted)
        ),
        acknowledgement_count=len(resolved.warnings_acknowledged),
        override_count=len(resolved.override_reasons),
    )
