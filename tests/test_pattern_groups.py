"""Step 8: semantically pure pattern review groups and the additive migration."""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Any

from qc_tool.engine import QCRunResult
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingProvenance,
    FindingSubtype,
    Materiality,
    Severity,
)
from qc_tool.history.store import RunHistory
from qc_tool.review import (
    build_pattern_groups,
    build_review_groups,
    count_pattern_groups,
    finding_identity_key,
    review_counts,
)
from qc_tool.triage.rules import triage

_ORACLE_DIR = Path(__file__).parent / "oracles"


def _oracle(name: str) -> dict[str, Any]:
    return json.loads((_ORACLE_DIR / name).read_text(encoding="utf-8"))


def _logic(location: str, baseline: str, current: str) -> Finding:
    return Finding(
        artifact="excel",
        finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
        sheet="Data",
        location=location,
        baseline_location=location,
        baseline_value=baseline,
        current_value=current,
        message=f"Data!{location}: formula logic changed",
    )


def _value(location: str, subtype: FindingSubtype) -> Finding:
    return Finding(
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        subtype=subtype,
        sheet="Data",
        location=location,
        baseline_location=location,
        baseline_value="" if subtype is FindingSubtype.VALUE_ADDED_POPULATION else "1",
        current_value="2",
        message=f"Data!{location}: value",
    )


def test_identical_shapes_merge_across_disconnected_ranges() -> None:
    findings = triage(
        [
            _logic("B2", "=A2", "=IFERROR(A2,0)"),
            _logic("B3", "=A3", "=IFERROR(A3,0)"),
            _logic("B40", "=A40", "=IFERROR(A40,0)"),
            _logic("D9", "=C9", "=IFERROR(C9,0)"),
        ]
    )

    groups = build_pattern_groups(findings)

    assert len(groups) == 1
    group = groups[0]
    assert group.member_count == 4
    assert not group.spatial
    assert group.ranges == ("B2:B3", "D9", "B40")
    assert group.bounding_range == "B2:D40"


def test_distinct_token_shapes_never_share_a_pattern_group() -> None:
    findings = triage(
        [
            _logic("B2", "=A2", "=IFERROR(A2,0)"),
            _logic("B3", "=A3", "=A3*2"),
        ]
    )

    spatial = build_review_groups(findings)
    patterns = build_pattern_groups(findings)

    assert len(spatial) == 1  # the spatial layer still merges adjacent cells
    assert len(patterns) == 2
    assert all(
        len({(member.baseline_value, member.current_value) for member in group.members})
        == 1
        for group in patterns
    )


def test_value_groups_never_mix_materiality_tiers() -> None:
    """A noise member and a material member never share a group, even when
    the rest of their partition key matches."""

    def tiered(location: str, tier: Materiality) -> Finding:
        finding = _value(location, FindingSubtype.VALUE_REPLACEMENT)
        finding.materiality = tier
        return finding

    findings = triage(
        [
            tiered("B2", Materiality.NOISE),
            tiered("B3", Materiality.NOISE),
            tiered("B4", Materiality.MATERIAL),
            tiered("B5", Materiality.RECENT_RESTATEMENT),
            tiered("B6", Materiality.WITHIN_TOLERANCE),
        ]
    )

    groups = [
        group
        for group in build_pattern_groups(findings)
        if group.finding_class is FindingClass.VALUE_CHANGED
    ]

    tiers_by_group = [
        {member.materiality for member in group.members} for group in groups
    ]
    assert all(len(tiers) == 1 for tiers in tiers_by_group)
    assert len(groups) == 4
    noise_group = next(
        group
        for group in groups
        if all(m.materiality is Materiality.NOISE for m in group.members)
    )
    assert noise_group.member_count == 2

    # Tier-less legacy findings still form their own stable group.
    legacy = triage([_value("C2", FindingSubtype.VALUE_REPLACEMENT)])
    legacy_groups = build_pattern_groups(legacy)
    assert len(legacy_groups) == 1


def test_consistency_groups_never_mix_dominant_patterns() -> None:
    findings = triage(
        [
            Finding(
                artifact="excel",
                finding_class=FindingClass.FORMULA_INCONSISTENT,
                provenance=FindingProvenance.INHERITED,
                sheet="Data",
                location=location,
                baseline_value=dominant,
                current_value=deviation,
                message="deviates",
            )
            for location, dominant, deviation in (
                ("B2", "=RC[-1]", "=RC[-1]*2"),
                ("B3", "=RC[-1]", "=RC[-1]*3"),
                ("B4", "=SUM(RC[-1])", "=RC[-1]*2"),
            )
        ]
    )

    groups = build_pattern_groups(findings)

    assert len(groups) == 2
    assert all(
        len({member.baseline_value for member in group.members}) == 1
        for group in groups
    )
    assert sorted(group.member_count for group in groups) == [1, 2]


def test_value_groups_never_mix_addition_with_replacement() -> None:
    findings = triage(
        [
            _value("A1", FindingSubtype.VALUE_ADDED_POPULATION),
            _value("A2", FindingSubtype.VALUE_REPLACEMENT),
            _value("A3", FindingSubtype.VALUE_REPLACEMENT),
        ]
    )

    groups = build_pattern_groups(findings)

    assert len(groups) == 2
    assert all(
        len({member.subtype for member in group.members}) == 1 for group in groups
    )


def test_pattern_groups_are_deterministic_and_lossless() -> None:
    findings = triage(
        [
            _logic("B2", "=A2", "=IFERROR(A2,0)"),
            _logic("B3", "=A3", "=A3*2"),
            _value("A1", FindingSubtype.VALUE_ADDED_POPULATION),
            Finding(
                artifact="excel",
                finding_class=FindingClass.CONDITIONAL_FORMAT_CHANGED,
                subtype=FindingSubtype.OBJECT_ADDED,
                sheet="Data",
                element="expression rule 1",
                message="conditional-format expression rule 1 added",
            ),
        ]
    )

    first = build_pattern_groups(findings)
    second = build_pattern_groups(list(reversed(findings)))

    assert [group.group_id for group in first] == [group.group_id for group in second]
    assert sum(group.member_count for group in first) == len(findings)
    assert {
        finding_identity_key(member) for group in first for member in group.members
    } == {finding_identity_key(finding) for finding in findings}
    counts = count_pattern_groups(first)
    assert sum(counts.review_items.values()) == len(first)
    assert sum(counts.atomic_findings.values()) == len(findings)


def test_pattern_groups_never_mix_severity_provenance_or_expected_state() -> None:
    findings = triage(
        [
            _logic("B2", "=A2", "=IFERROR(A2,0)"),
            _logic("B3", "=A3", "=IFERROR(A3,0)").model_copy(
                update={"expected_growth": True}
            ),
        ]
    )

    groups = build_pattern_groups(findings)

    assert len(groups) == 2
    assert {group.expected_growth for group in groups} == {False, True}
    assert {group.severity for group in groups} == {
        Severity.WARNING,
        Severity.EXPECTED,
    }


def test_history_stores_spatial_and_pattern_counts_separately(tmp_path: Path) -> None:
    findings = triage(
        [
            _logic("B2", "=A2", "=IFERROR(A2,0)"),
            _logic("B3", "=A3", "=A3*2"),
        ]
    )
    history = RunHistory(tmp_path / "history.sqlite3")

    run_id = history.record_run(
        QCRunResult(profile_name="patterns", findings=findings),
        file_hashes={},
        report_paths={},
    )
    record = history.get_run(run_id)

    spatial = review_counts(build_review_groups(findings))
    patterns = count_pattern_groups(build_pattern_groups(findings))
    assert record.review_counts == {
        severity.value: count for severity, count in spatial.review_items.items()
    }
    assert record.pattern_review_counts == {
        severity.value: count for severity, count in patterns.review_items.items()
    }
    assert sum(record.pattern_review_counts.values()) == 2
    assert sum(record.review_counts.values()) == 1
    assert record.story_counts, "stories must persist alongside pattern counts"
    assert sum(bucket["members"] for bucket in record.story_counts.values()) == len(
        findings
    )


def test_legacy_rows_keep_empty_pattern_counts_and_original_spatial_counts(
    tmp_path: Path,
) -> None:
    database = tmp_path / "legacy.sqlite3"
    with sqlite3.connect(database) as connection:
        connection.executescript(
            """
            CREATE TABLE runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                started_at TEXT NOT NULL,
                profile TEXT NOT NULL,
                files TEXT NOT NULL,
                file_hashes TEXT NOT NULL,
                counts TEXT NOT NULL,
                disclosures TEXT NOT NULL,
                verified_crosschecks INTEGER NOT NULL,
                findings TEXT NOT NULL,
                report_paths TEXT NOT NULL,
                review_counts TEXT NOT NULL DEFAULT '{}'
            );
            """
        )
        connection.execute(
            "INSERT INTO runs (started_at, profile, files, file_hashes, counts,"
            " disclosures, verified_crosschecks, findings, report_paths, review_counts)"
            " VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                "2026-01-01T00:00:00+00:00",
                "legacy",
                "{}",
                "{}",
                '{"critical": 4}',
                "[]",
                0,
                "[]",
                "{}",
                '{"critical": 2}',
            ),
        )

    record = RunHistory(database).list_runs()[0]

    assert record.review_counts == {"critical": 2}
    assert record.pattern_review_counts == {}
    assert record.story_counts == {}
    assert record.counts == {"critical": 4}


def test_step8_checkpoint_reconciles_with_the_delta_ledger() -> None:
    step7 = _oracle("real_workload_step7.json")
    step8 = _oracle("real_workload_step8.json")
    ledger = _oracle("step_delta_ledger.json")["steps"]["8"]

    assert step8["atomic_findings"] - step7["atomic_findings"] == (
        ledger["expected_atomic_delta"]
    )
    assert step8["severity"] == step7["severity"]
    assert step8["review_counts"] == step7["review_counts"]
    assert step8["coverage_states"] == step7["coverage_states"]
    assert sum(step8["pattern_review_counts"].values()) == (
        ledger["expected_pattern_review_items"]
    )
    assert sum(step8["pattern_review_counts"].values()) < sum(
        step7["review_counts"].values()
    )
    assert step8["mixed_pattern_groups"] == 0
    assert step8["source_hashes_unchanged"] is True
