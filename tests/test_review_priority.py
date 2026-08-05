"""Guided review prioritization: determinism, losslessness, and cited evidence."""

from __future__ import annotations

from qc_tool.engine import QCRunResult
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingProvenance,
    FindingTemporalContext,
    Materiality,
    Severity,
)
from qc_tool.review import (
    ReviewPriority,
    build_review_groups,
    prioritize_review,
)
from qc_tool.story import ChangeStory, StoryKind, build_stories


def _finding(
    index: int,
    finding_class: FindingClass,
    severity: Severity,
    **extra: object,
) -> Finding:
    return Finding(
        finding_id=f"F{index}",
        artifact="excel",
        finding_class=finding_class,
        severity=severity,
        sheet="Data",
        location=f"B{index}",
        message=f"finding {index}",
        **extra,  # type: ignore[arg-type]
    )


def _groups(findings: list[Finding]) -> list:
    return build_review_groups(findings)


# --- losslessness --------------------------------------------------------


def test_prioritization_is_a_permutation_of_its_input(qc_result: QCRunResult) -> None:
    """Ordering may de-emphasize, but it must never hide a decision."""
    groups = build_review_groups(qc_result.findings)

    ordered = prioritize_review(groups, build_stories(qc_result.findings))

    assert len(ordered) == len(groups)
    assert sorted(item.group.group_id for item in ordered) == sorted(
        group.group_id for group in groups
    )
    assert [item.rank for item in ordered] == list(range(1, len(groups) + 1))


def test_every_atomic_finding_stays_reachable(qc_result: QCRunResult) -> None:
    groups = build_review_groups(qc_result.findings)

    ordered = prioritize_review(groups)

    reachable = {
        member.finding_id
        for item in ordered
        for member in item.group.members
    }
    assert reachable == {finding.finding_id for finding in qc_result.findings}


# --- determinism ---------------------------------------------------------


def test_prioritization_is_deterministic(qc_result: QCRunResult) -> None:
    groups = build_review_groups(qc_result.findings)
    stories = build_stories(qc_result.findings)

    first = prioritize_review(groups, stories)
    second = prioritize_review(groups, stories)

    assert [(item.rank, item.group.group_id, item.score) for item in first] == [
        (item.rank, item.group.group_id, item.score) for item in second
    ]


def test_input_order_does_not_change_the_result(qc_result: QCRunResult) -> None:
    groups = build_review_groups(qc_result.findings)

    forward = prioritize_review(list(groups))
    reversed_input = prioritize_review(list(reversed(groups)))

    assert [item.group.group_id for item in forward] == [
        item.group.group_id for item in reversed_input
    ]


# --- ordering ------------------------------------------------------------


def test_material_critical_outranks_expected_growth() -> None:
    material = _finding(
        1,
        FindingClass.VALUE_CHANGED,
        Severity.CRITICAL,
        materiality=Materiality.MATERIAL,
        provenance=FindingProvenance.NEW,
    )
    expected = _finding(
        2, FindingClass.ROW_GROWTH, Severity.EXPECTED, expected_growth=True
    )

    ordered = prioritize_review(_groups([expected, material]))

    assert ordered[0].group.members[0].finding_id == "F1"
    assert ordered[-1].group.members[0].finding_id == "F2"


def test_noise_only_group_sinks_below_an_equal_severity_material_group() -> None:
    noise = _finding(
        1, FindingClass.VALUE_CHANGED, Severity.CRITICAL, materiality=Materiality.NOISE
    )
    material = _finding(
        5,
        FindingClass.VALUE_CHANGED,
        Severity.CRITICAL,
        materiality=Materiality.MATERIAL,
    )

    ordered = prioritize_review(_groups([noise, material]))

    assert ordered[0].group.members[0].finding_id == "F5"
    assert "noise_only" in ordered[-1].signals


def test_historical_change_outranks_current_period_change() -> None:
    historical = _finding(
        1,
        FindingClass.VALUE_CHANGED,
        Severity.CRITICAL,
        temporal_context=FindingTemporalContext.HISTORICAL,
    )
    current = _finding(
        5,
        FindingClass.VALUE_CHANGED,
        Severity.CRITICAL,
        temporal_context=FindingTemporalContext.CURRENT_PERIOD,
    )

    ordered = prioritize_review(_groups([current, historical]))

    assert ordered[0].group.members[0].finding_id == "F1"


def test_waived_group_sinks_to_the_bottom() -> None:
    waived = _finding(
        1, FindingClass.VALUE_CHANGED, Severity.CRITICAL, waiver_reason="approved"
    )
    plain = _finding(5, FindingClass.STYLE_CHANGED, Severity.INFO)

    ordered = prioritize_review(_groups([waived, plain]))

    assert ordered[-1].group.members[0].finding_id == "F1"
    assert "waived" in ordered[-1].signals


def test_story_explained_group_sinks_below_an_unexplained_one() -> None:
    explained = _finding(1, FindingClass.VALUE_CHANGED, Severity.WARNING)
    unexplained = _finding(5, FindingClass.VALUE_CHANGED, Severity.WARNING)
    story = ChangeStory(
        story_id="S1",
        kind=StoryKind.DATA_REFRESH,
        title="refresh",
        description="",
        evidence=(),
        finding_ids=("F1",),
        member_count=1,
        severity_counts={},
    )

    ordered = prioritize_review(_groups([explained, unexplained]), [story])

    assert ordered[0].group.members[0].finding_id == "F5"
    assert "unexplained_residual" in ordered[0].signals
    assert "story_explained" in ordered[-1].signals


def test_residual_story_membership_does_not_count_as_explained() -> None:
    finding = _finding(1, FindingClass.VALUE_CHANGED, Severity.WARNING)
    residual = ChangeStory(
        story_id="S9",
        kind=StoryKind.RESIDUAL,
        title="residual",
        description="",
        evidence=(),
        finding_ids=("F1",),
        member_count=1,
        severity_counts={},
    )

    ordered = prioritize_review(_groups([finding]), [residual])

    assert "unexplained_residual" in ordered[0].signals


# --- explanation ---------------------------------------------------------


def test_every_rationale_cites_observed_evidence(qc_result: QCRunResult) -> None:
    ordered = prioritize_review(
        build_review_groups(qc_result.findings), build_stories(qc_result.findings)
    )

    assert ordered
    for item in ordered:
        assert item.signals, f"{item.group.group_id} has no signal"
        assert item.rationale
        assert item.rationale != "no prioritization signal applies"


def test_rationale_reports_the_counts_it_scored_on() -> None:
    members = [
        _finding(
            index,
            FindingClass.VALUE_CHANGED,
            Severity.CRITICAL,
            materiality=Materiality.MATERIAL,
            provenance=FindingProvenance.NEW,
        )
        for index in range(1, 4)
    ]

    ordered = prioritize_review(_groups(members))

    rationale = ordered[0].rationale
    assert "critical severity" in rationale
    assert "3 member(s) exceed the materiality threshold" in rationale
    assert "3 member(s) have no baseline counterpart" in rationale


def test_priority_carries_its_group_unchanged() -> None:
    groups = _groups([_finding(1, FindingClass.VALUE_CHANGED, Severity.CRITICAL)])

    ordered = prioritize_review(groups)

    assert isinstance(ordered[0], ReviewPriority)
    assert ordered[0].group is groups[0]


# --- measured effect -----------------------------------------------------


def test_prioritization_front_loads_the_fixture_defects(
    qc_result: QCRunResult,
) -> None:
    """Measured against the generated corpus, whose defects are ground truth.

    A defect-bearing decision is one whose members are not expected growth. The
    guided order must reach the last of them no later than the existing order.
    """
    groups = build_review_groups(qc_result.findings)
    stories = build_stories(qc_result.findings)

    def last_defect_position(ordered: list) -> int:
        positions = [
            index
            for index, group in enumerate(ordered, start=1)
            if group.severity is not Severity.EXPECTED
        ]
        return max(positions) if positions else 0

    guided = [item.group for item in prioritize_review(groups, stories)]
    baseline_position = last_defect_position(groups)
    guided_position = last_defect_position(guided)

    assert guided_position <= baseline_position
    assert guided_position == sum(
        1 for group in groups if group.severity is not Severity.EXPECTED
    )


def test_no_expected_group_outranks_a_defect_group(qc_result: QCRunResult) -> None:
    """A false review decision here means noise displacing a real defect."""
    ordered = prioritize_review(
        build_review_groups(qc_result.findings), build_stories(qc_result.findings)
    )

    seen_expected = False
    misordered = 0
    for item in ordered:
        if item.group.severity is Severity.EXPECTED:
            seen_expected = True
        elif seen_expected:
            misordered += 1
    assert misordered == 0


# --- deeper stories ------------------------------------------------------


def test_residual_story_explains_what_it_could_not_explain(
    qc_result: QCRunResult,
) -> None:
    residual = next(
        story
        for story in build_stories(qc_result.findings)
        if story.kind is StoryKind.RESIDUAL
    )

    assert any(line.startswith("unexplained by class: ") for line in residual.evidence)
