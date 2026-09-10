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
import shutil
import tempfile
import time
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


# --- plan-20260910 Step 6: delta-encoded candidate spill --------------------


def test_candidate_spill_stores_only_the_first_row_of_a_group_in_full() -> None:
    """White-box check of the internal spill row shape: the first `add()`
    call for a population key writes a full row; every later call for the
    SAME key writes a smaller delta row, since `population_key` already
    guarantees roughly a dozen fields are identical across the group.
    """
    spill = CandidateSpill(_profile(), _TODAY)
    for row in range(2, 17):  # 15 members, one uniform shift
        candidate = _candidate(location=f"B{row}", baseline_location=f"B{row - 1}")
        spill.add(candidate, shape_before="digest-before", shape_after="digest-after")

    groups = list(spill.groups())
    assert len(groups) == 1
    group = groups[0]
    assert len(group) == 15
    kinds = [row["row"]["kind"] for row in group]
    assert kinds.count("full") == 1
    assert kinds.count("delta") == 14

    full_row = next(row for row in group if row["row"]["kind"] == "full")
    delta_row = next(row for row in group if row["row"]["kind"] == "delta")
    # The delta omits every field population_key already guarantees shared
    # (artifact/sheet/finding_class/severity/... are absent), keeping only
    # what genuinely varies member to member.
    assert len(delta_row["row"]["finding"]) < len(full_row["row"]["finding"])
    for shared_field in ("artifact", "sheet", "finding_class", "severity"):
        assert shared_field not in delta_row["row"]["finding"]
    assert "location" in delta_row["row"]["finding"]
    assert "baseline_location" in delta_row["row"]["finding"]


def test_reconstructed_candidates_are_field_identical_to_the_originals() -> None:
    """Parity check: every candidate reconstructed from the spill (whether
    stored full or delta-encoded) is field-for-field identical to what
    `finding_payload()` would have produced directly from the original
    in-memory candidate, before `assign_severities` mutates severity/waiver
    state and before spilling (both sides must reflect that mutation).
    """
    spill = CandidateSpill(_profile(), _TODAY)
    originals: list[Finding] = []
    for row in range(2, 21):  # 19 members, one uniform shift
        candidate = _candidate(location=f"B{row}", baseline_location=f"B{row - 1}")
        spill.add(candidate, shape_before="digest-before", shape_after="digest-after")
        originals.append(candidate)  # add() mutates severity in place

    policy = PopulationPolicy(enabled=True, threshold=100)  # force atomic replay
    outcome = finalize_populations(spill, policy, _NO_SCOPE)

    assert outcome.population_findings == []
    assert len(outcome.replay_findings) == 19
    reconstructed_by_location = {
        finding.location: finding for finding in outcome.replay_findings
    }
    for original in originals:
        reconstructed = reconstructed_by_location[original.location]
        assert reconstructed.model_dump(mode="json") == original.model_dump(mode="json")


def test_candidate_spill_handles_heterogeneous_evidence_with_delta_encoding() -> None:
    """The delta encoding must not hide a genuine per-member difference in a
    field `population_key` does NOT guarantee homogeneous (`evidence_tags`)
    -- reruns the heterogeneous-evidence scenario through the new codec.
    """
    spill = CandidateSpill(_profile(), _TODAY)
    for row in range(2, 8):  # 6 members, no tags
        spill.add(
            _candidate(location=f"C{row}", baseline_location=f"C{row - 1}"),
            shape_before="digest-before",
            shape_after="digest-after",
        )
    for row in range(8, 14):  # 6 members, a distinct event_key
        candidate = _candidate(location=f"C{row}", baseline_location=f"C{row - 1}")
        candidate.event_key = "distinct-event"
        spill.add(candidate, shape_before="digest-before", shape_after="digest-after")

    policy = PopulationPolicy(enabled=True, threshold=10)
    outcome = finalize_populations(spill, policy, _NO_SCOPE)

    assert outcome.population_findings == [], (
        "a genuine event_key difference within one population_key group must "
        "still force atomic replay, delta encoding or not"
    )
    assert len(outcome.replay_findings) == 12
    stats = outcome.stats[FindingClass.FORMULA_LOGIC_CHANGED]
    assert stats.heterogeneous_evidence == 12


def test_candidate_spill_delta_encoding_reduces_the_full_pipeline_wall_time() -> None:
    """Step 6's required benchmark: for a large, homogeneous population, the
    current add()+finalize_populations() cycle is measurably faster than a
    self-contained reference that mimics the pre-Step-6 approach exactly
    (always-full spill rows, always-full per-member reconstruction) on the
    identical fixture. Candidates here are built with the REAL `Finding(...)`
    constructor, not `model_construct` -- profiling this benchmark first
    with `model_construct` (matching this file's other, small-scale,
    correctness-only tests) showed pydantic's `model_construct` spending
    ~85% of total wall time in `inspect.signature()` introspection for
    unfilled defaults, a pure test-fixture artifact with zero relationship
    to production cost (real producers always call the normal `Finding(...)`
    constructor) -- that artifact was large enough to hide Step 6's real
    effect entirely at this row count. Uses `time.process_time()` +
    `min()`-over-repeats (Step 3/5's proven robust methodology).
    """
    from qc_tool.findings_store import SpillWriter, finding_payload, merge_spill
    from qc_tool.review import _coordinate, _rectangles
    from qc_tool.review import population_key as _population_key
    from qc_tool.triage.rules import assign_severities

    def _real_candidate(location: str, baseline_location: str) -> Finding:
        return Finding(
            artifact="excel",
            finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
            sheet="Sheet1",
            location=location,
            baseline_location=baseline_location,
            baseline_value="=A1",
            current_value="=A2",
            message=f"Sheet1!{location}: formula logic changed",
        )

    profile = _profile()
    rows = 20_000
    policy = PopulationPolicy(enabled=True, threshold=10)

    def run_reference() -> float:
        """Pre-Step-6 behavior: every spill row stores a full payload, and
        finalize reconstructs a full `Finding` for every member up front.
        """
        directory = Path(tempfile.mkdtemp(prefix="qc-population-reference-"))
        try:
            writer = SpillWriter(directory / "candidates.qcfb")
            writer.__enter__()
            started = time.process_time()
            for line_number in range(2, 2 + rows):
                candidate = _real_candidate(f"B{line_number}", f"B{line_number - 1}")
                assign_severities([candidate], profile, today=_TODAY)
                key = _population_key(candidate, ("digest-before", "digest-after"))
                coordinate = _coordinate(candidate.location) or (0, 0)
                sort_key = (*key, coordinate[0], coordinate[1])
                writer.append(
                    sort_key,
                    {
                        "finding": finding_payload(candidate),
                        "shape_before": "digest-before",
                        "shape_after": "digest-after",
                    },
                )
            writer.__exit__(None, None, None)
            current_key = None
            bucket: list[dict] = []
            groups: list[list[dict]] = []
            for raw_row in merge_spill(writer.path):
                assert isinstance(raw_row, dict)
                row: dict = raw_row
                finding = Finding.from_trusted_payload(row["finding"])
                key = _population_key(finding, (row["shape_before"], row["shape_after"]))
                if current_key is not None and key != current_key:
                    groups.append(bucket)
                    bucket = []
                current_key = key
                bucket.append(row)
            if bucket:
                groups.append(bucket)
            for group in groups:
                findings = [Finding.from_trusted_payload(row["finding"]) for row in group]
                members = [
                    (finding, current, _coordinate(finding.baseline_location))
                    for finding in findings
                    if (current := _coordinate(finding.location)) is not None
                ]
                if len(members) < policy.threshold:
                    continue
                current_coords = {current for _, current, _ in members}
                if len(_rectangles(current_coords)) > policy.max_rectangles:
                    continue
            return time.process_time() - started
        finally:
            shutil.rmtree(directory, ignore_errors=True)

    def run_current() -> float:
        spill = CandidateSpill(profile, _TODAY)
        started = time.process_time()
        for row in range(2, 2 + rows):
            candidate = _real_candidate(f"B{row}", f"B{row - 1}")
            spill.add(candidate, shape_before="digest-before", shape_after="digest-after")
        finalize_populations(spill, policy, _NO_SCOPE)
        return time.process_time() - started

    repeats = 5
    run_reference()  # untimed warm-up each side
    run_current()
    reference_times = [run_reference() for _ in range(repeats)]
    current_times = [run_current() for _ in range(repeats)]

    reference_best = min(reference_times)
    current_best = min(current_times)
    assert reference_best > 0.05, (
        "fixture too small to measure reliably "
        f"(reference best {reference_best:.4f}s CPU) -- widen it, don't "
        "loosen the improvement check"
    )
    improvement = (reference_best - current_best) / reference_best

    assert improvement >= 0.30, (
        f"Step 6's candidate pipeline improved only {improvement:.4%} vs the "
        f"pre-Step-6 reference (reference best {reference_best:.4f}s CPU, "
        f"current best {current_best:.4f}s CPU) -- expected a meaningful, "
        "not marginal, shrink"
    )


def test_diff_against_template_falls_back_to_full_when_template_has_extra_keys() -> None:
    from qc_tool.excel.population import _diff_against_template

    template = {"a": 1, "b": 2, "c": 3}
    payload_missing_c = {"a": 1, "b": 5}
    assert _diff_against_template(template, payload_missing_c) is None

    payload_same_keys = {"a": 1, "b": 5, "c": 3}
    assert _diff_against_template(template, payload_same_keys) == {"b": 5}

    payload_extra_key = {"a": 1, "b": 2, "c": 3, "d": 9}
    assert _diff_against_template(template, payload_extra_key) == {"d": 9}


def test_candidate_spill_delta_encoded_file_is_smaller_than_a_full_row_equivalent(
    tmp_path: Path,
) -> None:
    from qc_tool.findings_store import SpillWriter, finding_payload

    profile = _profile()
    rows = 20_000

    delta_spill = CandidateSpill(profile, _TODAY)
    for row in range(2, 2 + rows):
        candidate = _candidate(location=f"B{row}", baseline_location=f"B{row - 1}")
        delta_spill.add(candidate, shape_before="digest-before", shape_after="digest-after")
    delta_path = delta_spill._spill.path
    delta_spill._spill.__exit__(None, None, None)  # force-flush without reading back
    delta_bytes = delta_path.stat().st_size
    delta_spill.abort()  # cleans up the tempdir; safe on an already-exited writer

    full_path = tmp_path / "full_reference.qcfb"
    full_writer = SpillWriter(full_path)
    full_writer.__enter__()
    for row in range(2, 2 + rows):
        reference_candidate = _candidate(location=f"B{row}", baseline_location=f"B{row - 1}")
        full_writer.append(
            [row, 0],
            {
                "finding": finding_payload(reference_candidate),
                "shape_before": "digest-before",
                "shape_after": "digest-after",
            },
        )
    full_writer.__exit__(None, None, None)
    full_bytes = full_path.stat().st_size

    reduction = (full_bytes - delta_bytes) / full_bytes
    assert reduction >= 0.10, (
        f"delta-encoded spill ({delta_bytes} bytes) did not shrink by at "
        f"least 10% vs a full-row-per-candidate equivalent ({full_bytes} "
        f"bytes) -- actual reduction {reduction:.2%}"
    )

