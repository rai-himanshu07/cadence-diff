"""Group-first candidate spill, grouping, and emission (plan-20260906, A2).

Exercises `qc_tool.excel.population` directly: candidates are built the way
a producer would (`Finding.model_construct` plain kwargs, no validation),
spilled, grouped by population key, and either emitted as one population
`Finding` or returned for atomic replay. See `tests/test_population.py` for
the A1 model/codec/contract tests this builds on. The trailing section runs
the full `run_qc()` pipeline end to end on a real synthetic workbook pair,
once with the policy enabled and once disabled, and checks the population's
reconstructed pairs equal the atomic run's pairs exactly.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

from openpyxl import Workbook

from qc_tool.config.profile import (
    DeliverableProfile,
    PopulationPolicy,
    ReviewPolicy,
    default_profile,
)
from qc_tool.engine import run_qc
from qc_tool.excel.population import CandidateSpill, finalize_populations
from qc_tool.findings import Finding, FindingClass
from qc_tool.scope import ComparisonScope

_TODAY = dt.date(2026, 9, 7)
_NO_SCOPE = ComparisonScope()


def _profile() -> DeliverableProfile:
    return default_profile()


def _candidate(
    *,
    sheet: str = "Sheet1",
    location: str,
    baseline_location: str,
    baseline_value: str = "=A1",
    current_value: str = "=A2",
    finding_class: FindingClass = FindingClass.FORMULA_LOGIC_CHANGED,
) -> Finding:
    return Finding.model_construct(
        artifact="excel",
        finding_class=finding_class,
        sheet=sheet,
        location=location,
        baseline_location=baseline_location,
        baseline_value=baseline_value,
        current_value=current_value,
        message=f"{sheet}!{location}: formula logic changed",
    )


def test_shift_mode_group_at_or_above_threshold_becomes_one_population() -> None:
    spill = CandidateSpill(_profile(), _TODAY)
    for row in range(2, 17):  # 15 members, baseline is one row above current (dr=-1)
        candidate = _candidate(location=f"B{row}", baseline_location=f"B{row - 1}")
        spill.add(candidate, shape_before="digest-before", shape_after="digest-after")

    policy = PopulationPolicy(enabled=True, threshold=10)
    outcome = finalize_populations(spill, policy, _NO_SCOPE)

    assert outcome.replay_findings == []
    assert len(outcome.population_findings) == 1
    population = outcome.population_findings[0]
    assert population.population is not None
    assert population.population.member_count == 15
    assert population.population.membership.baseline_mode == "shift"
    assert population.population.membership.shift == (-1, 0)
    assert population.element == "population"
    assert population.location == "B2:B16"
    assert population.baseline_location == "B1:B15"
    assert population.root_cause_key.startswith("population:")
    stats = outcome.stats[FindingClass.FORMULA_LOGIC_CHANGED]
    assert stats.candidates == 15
    assert stats.populations == 1
    assert stats.replayed == 0


def test_below_threshold_group_replays_every_member_as_atomic() -> None:
    spill = CandidateSpill(_profile(), _TODAY)
    for row in range(2, 7):  # 5 members
        candidate = _candidate(location=f"C{row}", baseline_location=f"C{row - 1}")
        spill.add(candidate, shape_before="digest-before", shape_after="digest-after")

    policy = PopulationPolicy(enabled=True, threshold=10)
    outcome = finalize_populations(spill, policy, _NO_SCOPE)

    assert outcome.population_findings == []
    assert len(outcome.replay_findings) == 5
    assert all(finding.population is None for finding in outcome.replay_findings)
    stats = outcome.stats[FindingClass.FORMULA_LOGIC_CHANGED]
    assert stats.below_threshold == 5
    assert stats.replayed == 5
    assert stats.populations == 0


def test_pairs_mode_used_when_offsets_are_not_uniform() -> None:
    spill = CandidateSpill(_profile(), _TODAY)
    locations = [(f"D{row}", f"D{row - 1}") for row in range(2, 12)]
    # Break the uniform shift for one member -> forces explicit pairs mode.
    locations[3] = (locations[3][0], "D1")
    for current, baseline in locations:
        candidate = _candidate(location=current, baseline_location=baseline)
        spill.add(candidate, shape_before="digest-before", shape_after="digest-after")

    policy = PopulationPolicy(enabled=True, threshold=10)
    outcome = finalize_populations(spill, policy, _NO_SCOPE)

    assert len(outcome.population_findings) == 1
    membership = outcome.population_findings[0].population.membership  # type: ignore[union-attr]
    assert membership.baseline_mode == "pairs"
    assert membership.pairs is not None
    assert len(membership.pairs) == 10
    assert set(membership.pairs) == set(locations)


def test_over_rectangle_cap_replays_instead_of_emitting_a_population() -> None:
    spill = CandidateSpill(_profile(), _TODAY)
    # One member per row, rows spaced apart so each forms its own separate
    # rectangle (adjacent rows would merge into one run via `_rectangles`).
    for index in range(11):
        row = 5 + index * 3
        candidate = _candidate(location=f"B{row}", baseline_location=f"B{row - 1}")
        spill.add(candidate, shape_before="digest-before", shape_after="digest-after")

    policy = PopulationPolicy(enabled=True, threshold=10, max_rectangles=5)
    outcome = finalize_populations(spill, policy, _NO_SCOPE)

    assert outcome.population_findings == []
    assert len(outcome.replay_findings) == 11
    stats = outcome.stats[FindingClass.FORMULA_LOGIC_CHANGED]
    assert stats.over_cap == 11


def test_different_sheets_never_share_a_population() -> None:
    spill = CandidateSpill(_profile(), _TODAY)
    for row in range(2, 13):
        spill.add(
            _candidate(sheet="Sheet1", location=f"B{row}", baseline_location=f"B{row - 1}"),
            shape_before="digest-before",
            shape_after="digest-after",
        )
    for row in range(2, 13):
        spill.add(
            _candidate(sheet="Sheet2", location=f"B{row}", baseline_location=f"B{row - 1}"),
            shape_before="digest-before",
            shape_after="digest-after",
        )

    policy = PopulationPolicy(enabled=True, threshold=10)
    outcome = finalize_populations(spill, policy, _NO_SCOPE)

    assert len(outcome.population_findings) == 2
    assert {p.sheet for p in outcome.population_findings} == {"Sheet1", "Sheet2"}


def test_scope_excludes_one_groups_sheet_while_another_group_still_populates() -> None:
    """Scope filters per finding's sheet, and every group's members share one
    sheet by construction (sheet is part of the population key) -- so scope
    can never partially trim a single group's member count. What it CAN do
    is drop one group's sheet entirely while a different, scope-included
    group on another sheet is unaffected and still becomes a population.
    """
    spill = CandidateSpill(_profile(), _TODAY)
    for row in range(2, 17):
        spill.add(
            _candidate(sheet="Sheet1", location=f"B{row}", baseline_location=f"B{row - 1}"),
            shape_before="digest-before",
            shape_after="digest-after",
        )
    for row in range(2, 17):
        spill.add(
            _candidate(sheet="Sheet2", location=f"B{row}", baseline_location=f"B{row - 1}"),
            shape_before="digest-before",
            shape_after="digest-after",
        )

    policy = PopulationPolicy(enabled=True, threshold=10)
    scope = ComparisonScope(excel_sheets=("Sheet2",))  # excludes Sheet1 entirely
    outcome = finalize_populations(spill, policy, scope)

    assert len(outcome.population_findings) == 1
    assert outcome.population_findings[0].sheet == "Sheet2"
    assert outcome.replay_findings == []
    # Sheet1's excluded group contributes no stats at all -- out of scope
    # means it never existed in this comparison, same as an ordinary finding.
    assert outcome.stats[FindingClass.FORMULA_LOGIC_CHANGED].candidates == 15


def test_scope_dropping_every_candidate_leaves_the_outcome_empty() -> None:
    spill = CandidateSpill(_profile(), _TODAY)
    for row in range(2, 17):
        spill.add(
            _candidate(sheet="Sheet1", location=f"B{row}", baseline_location=f"B{row - 1}"),
            shape_before="digest-before",
            shape_after="digest-after",
        )

    policy = PopulationPolicy(enabled=True, threshold=10)
    scope = ComparisonScope(excel_sheets=("Sheet2",))  # excludes every candidate
    outcome = finalize_populations(spill, policy, scope)

    assert outcome.population_findings == []
    assert outcome.replay_findings == []
    assert outcome.stats == {}


# --- end-to-end: real run_qc() pipeline, policy on vs. off -------------------


def _write_formula_pair(dest: Path, *, rows: int) -> tuple[Path, Path]:
    """A uniform formula-change pair: column B's formula changes every row."""

    def build(path: Path, *, multiplier: int) -> None:
        workbook = Workbook()
        sheet = workbook.active
        assert sheet is not None
        sheet.title = "Data"
        sheet.append(["Input", "Output"])
        for row in range(2, 2 + rows):
            sheet.append([100, f"=A{row}*{multiplier}"])
        workbook.save(path)

    base, curr = dest / "formula_base.xlsx", dest / "formula_curr.xlsx"
    build(base, multiplier=2)
    build(curr, multiplier=3)
    return base, curr


def _expand_rectangle(range_ref: str) -> set[str]:
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import range_boundaries

    min_col, min_row, max_col, max_row = range_boundaries(range_ref)
    assert min_col is not None
    assert min_row is not None
    assert max_col is not None
    assert max_row is not None
    return {
        f"{get_column_letter(col)}{row}"
        for row in range(min_row, max_row + 1)
        for col in range(min_col, max_col + 1)
    }


def test_population_pipeline_end_to_end_matches_atomic_pairs(tmp_path: Path) -> None:
    base, curr = _write_formula_pair(tmp_path, rows=19)

    atomic_result = run_qc(baseline_excel=base, current_excel=curr)
    atomic_pairs = {
        (finding.location, finding.baseline_location)
        for finding in atomic_result.findings
        if finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
    }
    assert len(atomic_pairs) == 19

    profile = DeliverableProfile(
        name="population-e2e",
        review_policy=ReviewPolicy(
            populations=PopulationPolicy(enabled=True, threshold=10)
        ),
    )
    population_result = run_qc(baseline_excel=base, current_excel=curr, profile=profile)
    population_findings = [
        finding
        for finding in population_result.findings
        if finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
    ]
    assert len(population_findings) == 1
    population = population_findings[0]
    assert population.population is not None
    assert population.population.member_count == 19
    assert population.element == "population"

    membership = population.population.membership
    assert membership.baseline_mode == "shift"
    assert membership.shift == (0, 0)  # formula changed in place, never moved
    # Criterion 6: a population's generic `impacts` field must stay empty --
    # any sample-derived impacts live only on the typed `sampled_impacts`
    # field, never mistakable for the population's exhaustive downstream set.
    assert population.impacts == []

    current_coordinates: set[str] = set()
    for rectangle in membership.current_rectangles:
        current_coordinates |= _expand_rectangle(rectangle)
    reconstructed_pairs = {(coord, coord) for coord in current_coordinates}
    assert reconstructed_pairs == atomic_pairs

    non_population_atomics = [
        finding
        for finding in population_result.findings
        if finding.finding_class is not FindingClass.FORMULA_LOGIC_CHANGED
    ]
    assert non_population_atomics == [
        finding
        for finding in atomic_result.findings
        if finding.finding_class is not FindingClass.FORMULA_LOGIC_CHANGED
    ]


def test_population_pipeline_below_threshold_replays_atomically(tmp_path: Path) -> None:
    base, curr = _write_formula_pair(tmp_path, rows=5)

    atomic_result = run_qc(baseline_excel=base, current_excel=curr)
    atomic_findings = [
        finding
        for finding in atomic_result.findings
        if finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
    ]
    assert len(atomic_findings) == 5

    profile = DeliverableProfile(
        name="population-e2e-below-threshold",
        review_policy=ReviewPolicy(
            populations=PopulationPolicy(enabled=True, threshold=10)
        ),
    )
    population_result = run_qc(baseline_excel=base, current_excel=curr, profile=profile)
    replayed = [
        finding
        for finding in population_result.findings
        if finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
    ]
    assert len(replayed) == 5
    assert all(finding.population is None for finding in replayed)

    def digest(result_findings: list[Finding]) -> list[tuple[object, ...]]:
        return [
            (
                finding.severity,
                finding.finding_class,
                finding.sheet,
                finding.location,
                finding.baseline_location,
                finding.baseline_value,
                finding.current_value,
                finding.message,
            )
            for finding in sorted(result_findings, key=lambda f: f.location or "")
        ]

    assert digest(replayed) == digest(atomic_findings)

