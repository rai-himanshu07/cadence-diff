"""Streaming builders must equal the list-based builders byte-for-byte."""

from __future__ import annotations

import copy
from collections.abc import Sequence

from qc_tool.engine import QCRunResult
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingEvidenceTag,
    Severity,
)
from qc_tool.review import (
    build_pattern_groups,
    build_review_groups,
    count_pattern_groups,
    review_counts,
)
from qc_tool.review_stream import (
    counts_from_summaries,
    stream_stories,
    summarize_pattern_groups,
    summarize_review_groups,
    summary_of,
)
from qc_tool.story import (
    StoryEvidenceContext,
    annotate_story_evidence,
    build_stories,
)


def _multi_member(findings: Sequence[Finding]) -> list[Finding]:
    cloned: list[Finding] = []
    for index, finding in enumerate(findings):
        clone = finding.model_copy(deep=True)
        clone.artifact_member = "member-b" if index % 3 == 0 else "primary"
        cloned.append(clone)
    return cloned


def test_review_group_summaries_match_list_builder(qc_result: QCRunResult) -> None:
    findings = qc_result.findings
    expected = [summary_of(group) for group in build_review_groups(findings)]
    streamed = summarize_review_groups(iter(findings))
    assert streamed == expected


def test_pattern_group_summaries_match_list_builder(qc_result: QCRunResult) -> None:
    findings = qc_result.findings
    expected = [summary_of(group) for group in build_pattern_groups(findings)]
    streamed = summarize_pattern_groups(iter(findings))
    assert streamed == expected


def test_counts_match_list_builders(qc_result: QCRunResult) -> None:
    findings = qc_result.findings
    assert counts_from_summaries(
        summarize_review_groups(iter(findings))
    ) == review_counts(build_review_groups(findings))
    assert counts_from_summaries(
        summarize_pattern_groups(iter(findings))
    ) == count_pattern_groups(build_pattern_groups(findings))


def test_summaries_match_for_multi_member_findings(qc_result: QCRunResult) -> None:
    findings = _multi_member(qc_result.findings)
    assert summarize_review_groups(iter(findings)) == [
        summary_of(group) for group in build_review_groups(findings)
    ]
    assert summarize_pattern_groups(iter(findings)) == [
        summary_of(group) for group in build_pattern_groups(findings)
    ]


def test_duplicate_singletons_keep_unique_ids_and_order() -> None:
    def clone(index: int) -> Finding:
        return Finding(
            finding_id=f"F{index:04d}",
            artifact="excel",
            finding_class=FindingClass.SHEET_REMOVED,
            severity=Severity.CRITICAL,
            sheet="Data",
            element="Data",
            message="sheet 'Data' removed in current cycle",
        )

    findings = [clone(1), clone(2)]
    expected = [summary_of(group) for group in build_review_groups(findings)]
    streamed = summarize_review_groups(iter(findings))
    assert streamed == expected
    assert streamed[0].group_id != streamed[1].group_id
    assert streamed[1].group_id.endswith("-2")


def test_stream_stories_match_batch_builder(qc_result: QCRunResult) -> None:
    findings = [finding.model_copy(deep=True) for finding in qc_result.findings]
    expected = build_stories(findings)  # annotates in place first
    streamed = stream_stories([iter(findings), iter(findings), iter(findings)])
    assert streamed == expected


def test_stream_stories_match_for_multi_member(qc_result: QCRunResult) -> None:
    findings = _multi_member(qc_result.findings)
    expected = build_stories(findings)
    streamed = stream_stories([iter(findings), iter(findings), iter(findings)])
    assert streamed == expected


def test_two_phase_annotation_equals_recorded_tags(qc_result: QCRunResult) -> None:
    batch = [finding.model_copy(deep=True) for finding in qc_result.findings]
    annotate_story_evidence(batch)

    streamed = [finding.model_copy(deep=True) for finding in qc_result.findings]
    context = StoryEvidenceContext.collect(iter(streamed))
    for finding in streamed:
        context.apply(finding)

    assert [sorted(t.value for t in f.evidence_tags) for f in streamed] == [
        sorted(t.value for t in f.evidence_tags) for f in batch
    ]


def test_two_phase_annotation_is_idempotent() -> None:
    driver = Finding(
        finding_id="F0001",
        artifact="excel",
        finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
        severity=Severity.WARNING,
        sheet="Data",
        element="tbl_Revenue",
        current_value="tbl_Revenue[NewCol] added",
        message="Excel table 'tbl_Revenue' columns added",
    )
    formula = Finding(
        finding_id="F0002",
        artifact="excel",
        finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
        severity=Severity.WARNING,
        sheet="Data",
        location="B2",
        baseline_value="=SUM(A1:A9)",
        current_value="=SUM(tbl_Revenue[NewCol])",
        message="formula changed",
    )
    style = Finding(
        finding_id="F0003",
        artifact="excel",
        finding_class=FindingClass.STYLE_CHANGED,
        severity=Severity.INFO,
        sheet="Data",
        location="B2",
        message="style changed",
    )
    findings = [driver, formula, style]
    annotate_story_evidence(findings)
    first = copy.deepcopy([sorted(t.value for t in f.evidence_tags) for f in findings])
    assert FindingEvidenceTag.RESOLVED_DRIVER in formula.evidence_tags
    assert FindingEvidenceTag.EXACT_COLOCATION in formula.evidence_tags
    assert FindingEvidenceTag.EXACT_COLOCATION in style.evidence_tags

    annotate_story_evidence(findings)
    assert [sorted(t.value for t in f.evidence_tags) for f in findings] == first


def test_priority_aggregates_match_signal_counts(qc_result: QCRunResult) -> None:
    """The stored rollup equals what _priority_signals reads from live members."""
    from qc_tool.findings import (
        FindingProvenance,
        FindingTemporalContext,
        Materiality,
    )
    from qc_tool.review_stream import summarize_pattern_groups_with_priority

    findings = qc_result.findings
    summaries, aggregates = summarize_pattern_groups_with_priority(iter(findings))
    assert summaries == summarize_pattern_groups(iter(findings))
    groups = {group.group_id: group for group in build_pattern_groups(findings)}
    assert set(aggregates) == set(groups)
    for group_id, aggregate in aggregates.items():
        members = groups[group_id].members
        assert aggregate.material == sum(
            1 for item in members if item.materiality is Materiality.MATERIAL
        )
        assert aggregate.historical == sum(
            1
            for item in members
            if item.temporal_context is FindingTemporalContext.HISTORICAL
        )
        assert aggregate.new == sum(
            1 for item in members if item.provenance is FindingProvenance.NEW
        )
        assert aggregate.impacts == sum(len(item.impacts) for item in members)
        assert set(aggregate.tiers) == {
            item.materiality.value
            for item in members
            if item.materiality is not None
        }
        assert aggregate.all_inherited == all(
            item.provenance is FindingProvenance.INHERITED for item in members
        )
        assert aggregate.any_waiver == any(item.waiver_reason for item in members)
        assert aggregate.any_reviewed == any(
            item.severity_overridden or item.analyst_comment for item in members
        )


def test_summary_prioritization_matches_materialized(qc_result: QCRunResult) -> None:
    """The summary prioritizer is an exact twin of prioritize_review."""
    from qc_tool.review import prioritize_review
    from qc_tool.review_stream import (
        prioritize_review_summaries,
        summarize_pattern_groups_with_priority,
    )

    findings = [finding.model_copy(deep=True) for finding in qc_result.findings]
    stories = build_stories(findings)
    materialized = prioritize_review(build_pattern_groups(findings), stories)
    summaries, aggregates = summarize_pattern_groups_with_priority(iter(findings))
    streamed = prioritize_review_summaries(summaries, aggregates, stories)
    assert [item.summary.group_id for item in streamed] == [
        item.group.group_id for item in materialized
    ]
    for summary_item, group_item in zip(streamed, materialized, strict=True):
        assert summary_item.rank == group_item.rank
        assert summary_item.score == group_item.score
        assert summary_item.signals == group_item.signals
        assert summary_item.rationale == group_item.rationale


def test_summary_prioritization_honors_live_annotations(
    qc_result: QCRunResult,
) -> None:
    """Live decisions defer a group exactly like decorated members do."""
    from qc_tool.review import prioritize_review
    from qc_tool.review_stream import (
        prioritize_review_summaries,
        summarize_pattern_groups_with_priority,
    )

    findings = [finding.model_copy(deep=True) for finding in qc_result.findings]
    summaries, aggregates = summarize_pattern_groups_with_priority(iter(findings))
    stories = build_stories(findings)
    # annotate the first member of the top group, as a persisted decision would
    target = prioritize_review_summaries(summaries, aggregates, stories)[0]
    annotated_id = target.summary.member_finding_ids[0]
    for finding in findings:
        if finding.finding_id == annotated_id:
            finding.analyst_comment = "reviewed"
    materialized = prioritize_review(build_pattern_groups(findings), stories)
    streamed = prioritize_review_summaries(
        summaries, aggregates, stories, reviewed_ids={annotated_id}
    )
    assert [item.summary.group_id for item in streamed] == [
        item.group.group_id for item in materialized
    ]
    assert "already_reviewed" in streamed[-1].signals or any(
        "already_reviewed" in item.signals for item in streamed
    )


def test_fused_record_accumulators_match_separate_passes(
    qc_result: QCRunResult,
) -> None:
    """record_run's single fused pass equals the two separate summarizers."""
    from qc_tool.review_stream import (
        RecordSummaryAccumulators,
        summarize_pattern_groups_with_priority,
    )

    findings = qc_result.findings
    fused = RecordSummaryAccumulators()
    for finding in findings:
        fused.observe(finding)
    assert fused.review.finish() == summarize_review_groups(iter(findings))
    pattern_summaries, aggregates = fused.pattern.finish()
    expected_summaries, expected_aggregates = (
        summarize_pattern_groups_with_priority(iter(findings))
    )
    assert pattern_summaries == expected_summaries
    assert aggregates == expected_aggregates
