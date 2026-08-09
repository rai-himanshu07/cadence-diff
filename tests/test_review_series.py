"""Frozen contracts for the logical-series review lens.

The run-18-equivalent oracle is generated here, never loaded from private data:
three canonical decisions of 8 + 9 + 2 members that the lens must promote into
exactly four structural parents.
"""

from __future__ import annotations

from typing import Literal

import pytest
from openpyxl.utils.cell import coordinate_to_tuple

from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingSubtype,
    FindingTemporalContext,
    Materiality,
    SeriesAnchor,
    SeriesAnchorV1,
    SeriesAnchorV2,
    Severity,
)
from qc_tool.review import ReviewGroup, build_pattern_groups
from qc_tool.review_series import (
    anchor_matches_finding,
    anchor_segment,
    build_series_review_lens,
    canonical_series_aggregate_digest,
    canonical_series_anchor_digest,
    cluster_confirmation_updates,
    series_anchor_eligible,
    series_anchor_segment,
    series_cluster_key,
    series_cluster_label,
)

SHEET = "Ops"
#: Two proved regions on one sheet: the C/D block and the L/M block.
REGION_HI = "Ops!B215:H240"
REGION_LO = "Ops!B130:P140"

_Band = Literal["historical", "recent", "current", "new_period", "cleared"]

HISTORICAL = ("C215", "C229", "D215", "D229", "L132", "L136", "M132", "M136")
RECENT = ("C234", "C237", "C239", "D234", "D236", "D237", "D239", "L137", "M137")
CURRENT = ("L138", "M138")
#: Periods populated for the first time this cycle; no baseline number exists,
#: so these carry no materiality tier and no temporal context.
NEW_PERIOD = ("L139", "M139")
#: Historical cells cleared to blank while sibling measures keep the period
#: alive; no current number exists, so no materiality tier or temporal context.
CLEARED = ("L140", "M140")

_TEMPORAL = {
    "historical": (Severity.CRITICAL, FindingTemporalContext.HISTORICAL),
    "recent": (Severity.WARNING, FindingTemporalContext.RECENT_WINDOW),
    "current": (Severity.WARNING, FindingTemporalContext.CURRENT_PERIOD),
}


def _region_for(location: str) -> str:
    row, _col = coordinate_to_tuple(location)
    return REGION_HI if row >= 200 else REGION_LO


def series_oracle_finding(
    finding_id: str,
    location: str,
    band: str,
    *,
    artifact_member: str = "primary",
) -> Finding:
    if band == "new_period":
        return Finding(
            finding_id=finding_id,
            artifact="excel",
            artifact_member=artifact_member,
            finding_class=FindingClass.VALUE_CHANGED,
            severity=Severity.CRITICAL,
            subtype=FindingSubtype.VALUE_ADDED_POPULATION,
            sheet=SHEET,
            location=location,
            baseline_location=location,
            current_value="120",
            message=f"{SHEET}!{location}: value added to a previously blank cell",
        )
    if band == "cleared":
        return Finding(
            finding_id=finding_id,
            artifact="excel",
            artifact_member=artifact_member,
            finding_class=FindingClass.VALUE_CHANGED,
            severity=Severity.CRITICAL,
            subtype=FindingSubtype.VALUE_CLEARED_POPULATION,
            sheet=SHEET,
            location=location,
            baseline_location=location,
            baseline_value="130",
            message=f"{SHEET}!{location}: historical value cleared",
        )
    severity, temporal = _TEMPORAL[band]
    return Finding(
        finding_id=finding_id,
        artifact="excel",
        artifact_member=artifact_member,
        finding_class=FindingClass.VALUE_CHANGED,
        severity=severity,
        subtype=FindingSubtype.VALUE_REPLACEMENT,
        materiality=Materiality.MATERIAL,
        temporal_context=temporal,
        sheet=SHEET,
        location=location,
        baseline_location=location,
        baseline_value="100",
        current_value="110",
        message=f"{SHEET}!{location}: historical value changed",
    )


def series_oracle_anchor(
    location: str,
    *,
    segment: Literal["restatement", "new_period", "cleared_period"] = "restatement",
) -> SeriesAnchorV2:
    row, col = coordinate_to_tuple(location)
    return SeriesAnchorV2(
        sheet=SHEET,
        current_region_id=_region_for(location),
        period_axis="rows",
        series_index=col,
        period_index=row,
        segment=segment,
    )


def series_oracle_anchor_v1(location: str) -> SeriesAnchorV1:
    row, col = coordinate_to_tuple(location)
    return SeriesAnchorV1(
        sheet=SHEET,
        current_region_id=_region_for(location),
        period_axis="rows",
        series_index=col,
        period_index=row,
    )


def series_oracle(
    *, with_new_period: bool = False, with_cleared: bool = False
) -> tuple[list[Finding], dict[str, SeriesAnchor]]:
    """The exact run-18-equivalent trio and its producer-authored anchors."""
    findings: list[Finding] = []
    anchors: dict[str, SeriesAnchor] = {}
    bands: list[tuple[_Band, tuple[str, ...]]] = [
        ("historical", HISTORICAL),
        ("recent", RECENT),
        ("current", CURRENT),
    ]
    if with_new_period:
        bands.append(("new_period", NEW_PERIOD))
    if with_cleared:
        bands.append(("cleared", CLEARED))
    segments: dict[str, Literal["restatement", "new_period", "cleared_period"]] = {
        "new_period": "new_period",
        "cleared": "cleared_period",
    }
    index = 0
    for band, locations in bands:
        for location in locations:
            index += 1
            finding_id = f"F{index}"
            findings.append(series_oracle_finding(finding_id, location, band))
            anchors[finding_id] = series_oracle_anchor(
                location, segment=segments.get(band, "restatement")
            )
    return findings, anchors


def _bindings(
    findings: list[Finding], anchors: dict[str, SeriesAnchor]
) -> dict[str, tuple[str, str, SeriesAnchor]]:
    by_id = {finding.finding_id: finding for finding in findings}
    return {
        finding_id: (
            by_id[finding_id].artifact_member,
            by_id[finding_id].location or "",
            anchor,
        )
        for finding_id, anchor in anchors.items()
    }


def _groups(findings: list[Finding]) -> list[ReviewGroup]:
    return build_pattern_groups(findings)


# --- canonical oracle ---------------------------------------------------------


def test_oracle_is_exactly_eight_nine_and_two_canonical_members() -> None:
    findings, anchors = series_oracle()

    assert len(findings) == 19
    assert len(anchors) == 19

    groups = _groups(findings)
    assert len(groups) == 3
    by_size = {group.member_count: group for group in groups}
    assert sorted(by_size) == [2, 8, 9]
    assert by_size[8].severity is Severity.CRITICAL
    assert by_size[9].severity is Severity.WARNING
    assert by_size[2].severity is Severity.WARNING
    assert {member.temporal_context for member in by_size[8].members} == {
        FindingTemporalContext.HISTORICAL
    }
    assert {member.temporal_context for member in by_size[9].members} == {
        FindingTemporalContext.RECENT_WINDOW
    }
    assert {member.temporal_context for member in by_size[2].members} == {
        FindingTemporalContext.CURRENT_PERIOD
    }


def test_every_oracle_finding_is_anchor_eligible_and_locator_agrees() -> None:
    findings, anchors = series_oracle()

    for finding in findings:
        assert series_anchor_eligible(finding)
        assert anchor_matches_finding(finding, anchors[finding.finding_id])


# --- digests ------------------------------------------------------------------


def test_entry_digest_binds_finding_member_locator_and_anchor() -> None:
    anchor = series_oracle_anchor("M137")
    digest = canonical_series_anchor_digest("F1", "primary", "M137", anchor)

    assert digest != canonical_series_anchor_digest("F2", "primary", "M137", anchor)
    assert digest != canonical_series_anchor_digest("F1", "ops", "M137", anchor)
    assert digest != canonical_series_anchor_digest("F1", "primary", "M138", anchor)
    assert digest != canonical_series_anchor_digest(
        "F1",
        "primary",
        "M137",
        anchor.model_copy(update={"period_index": 138}),
    )
    assert digest == canonical_series_anchor_digest("F1", "primary", "M137", anchor)


def test_aggregate_digest_binds_the_complete_sorted_population() -> None:
    findings, anchors = series_oracle()
    bindings = _bindings(findings, anchors)
    digest = canonical_series_aggregate_digest(bindings)

    assert digest == canonical_series_aggregate_digest(dict(reversed(list(bindings.items()))))

    dropped = dict(bindings)
    dropped.pop("F1")
    assert canonical_series_aggregate_digest(dropped) != digest

    tampered = dict(bindings)
    member, location, anchor = tampered["F1"]
    tampered["F1"] = (member, location, anchor.model_copy(update={"series_index": 26}))
    assert canonical_series_aggregate_digest(tampered) != digest

    assert canonical_series_aggregate_digest({}) != ""
    assert len(canonical_series_aggregate_digest({})) == 64


# --- cluster identity ---------------------------------------------------------


def test_cluster_key_is_member_qualified() -> None:
    anchor = series_oracle_anchor("M137")
    core = series_oracle_finding("F1", "M137", "recent", artifact_member="core")
    ops = series_oracle_finding("F2", "M137", "recent", artifact_member="ops")

    assert series_cluster_key(core, anchor) != series_cluster_key(ops, anchor)
    assert series_cluster_key(core, anchor)[0] == "core"
    assert series_cluster_key(ops, anchor)[0] == "ops"


def test_cluster_key_separates_columns_regions_sheets_and_axes() -> None:
    findings, anchors = series_oracle()
    by_id = {finding.finding_id: finding for finding in findings}
    keys = {
        series_cluster_key(by_id[finding_id], anchor)
        for finding_id, anchor in anchors.items()
    }

    assert len(keys) == 4
    assert keys == {
        ("primary", SHEET, REGION_HI, "rows", 3),
        ("primary", SHEET, REGION_HI, "rows", 4),
        ("primary", SHEET, REGION_LO, "rows", 12),
        ("primary", SHEET, REGION_LO, "rows", 13),
    }


@pytest.mark.parametrize(
    "update",
    [
        {"sheet": "Other"},
        {"period_axis": "columns"},
        {"series_index": 99},
        {"period_index": 99},
    ],
)
def test_anchor_must_agree_with_the_current_a1_locator(update: dict[str, object]) -> None:
    finding = series_oracle_finding("F1", "M137", "recent")
    anchor = series_oracle_anchor("M137")

    assert anchor_matches_finding(finding, anchor)
    assert not anchor_matches_finding(finding, anchor.model_copy(update=update))


@pytest.mark.parametrize(
    "update",
    [
        {"subtype": FindingSubtype.VALUE_ADDED_POPULATION},
        {"subtype": FindingSubtype.VALUE_CLEARED_POPULATION},
        {"materiality": None},
        {"temporal_context": None},
        {"finding_class": FindingClass.FORMULA_LOGIC_CHANGED},
        {"location": "M137:M138"},
    ],
)
def test_ineligible_findings_never_match_an_anchor(update: dict[str, object]) -> None:
    finding = series_oracle_finding("F1", "M137", "recent").model_copy(update=update)

    assert not anchor_matches_finding(finding, series_oracle_anchor("M137"))


# --- structural labels --------------------------------------------------------


def test_parent_label_is_structural_only() -> None:
    label = series_cluster_label(("primary", SHEET, REGION_LO, "rows", 13), 4)

    assert label == "Ops · column M · 4 changed periods"

    findings, _anchors = series_oracle()
    values = {finding.current_value for finding in findings} | {
        finding.baseline_value for finding in findings
    }
    for value in values:
        assert value is not None
        assert value not in label

    assert series_cluster_label(("primary", SHEET, REGION_LO, "columns", 7), 1) == (
        "Ops · row 7 · 1 changed period"
    )


# --- lens promotion oracle ----------------------------------------------------


def test_full_sheet_lens_promotes_exactly_four_structural_parents() -> None:
    findings, anchors = series_oracle()
    lens = build_series_review_lens(_groups(findings), anchors)

    assert len(lens.clusters) == 4
    assert lens.fallbacks == ()
    assert len(lens.rows) == 4

    by_column = {
        cluster.key[4]: cluster for cluster in lens.clusters
    }
    assert sorted(by_column) == [3, 4, 12, 13]
    assert [len(by_column[column].slices) for column in (3, 4, 12, 13)] == [2, 2, 3, 3]
    assert [by_column[column].member_count for column in (3, 4, 12, 13)] == [5, 6, 4, 4]
    assert sum(cluster.member_count for cluster in lens.clusters) == 19

    assert by_column[13].label == "Ops · column M · 4 changed periods"
    assert by_column[3].label == "Ops · column C · 5 changed periods"


def test_m_series_parent_holds_two_one_and_one_child_slices() -> None:
    findings, anchors = series_oracle()
    lens = build_series_review_lens(_groups(findings), anchors)

    cluster = next(item for item in lens.clusters if item.key[4] == 13)
    assert [slice_.member_count for slice_ in cluster.slices] == [2, 1, 1]
    assert [
        tuple(member.location for member in slice_.members)
        for slice_ in cluster.slices
    ] == [("M132", "M136"), ("M137",), ("M138",)]
    assert [slice_.members[0].severity for slice_ in cluster.slices] == [
        Severity.CRITICAL,
        Severity.WARNING,
        Severity.WARNING,
    ]
    assert [slice_.members[0].temporal_context for slice_ in cluster.slices] == [
        FindingTemporalContext.HISTORICAL,
        FindingTemporalContext.RECENT_WINDOW,
        FindingTemporalContext.CURRENT_PERIOD,
    ]
    canonical = {group.group_id for group in _groups(findings)}
    assert {slice_.group_id for slice_ in cluster.slices} == canonical
    assert all(not slice_.whole_group for slice_ in cluster.slices)


def test_lens_rows_are_an_exact_permutation_of_every_atomic_finding() -> None:
    findings, anchors = series_oracle()
    lens = build_series_review_lens(_groups(findings), anchors)

    seen = [
        member.finding_id
        for cluster in lens.clusters
        for slice_ in cluster.slices
        for member in slice_.members
    ] + [
        member.finding_id
        for slice_ in lens.fallbacks
        for member in slice_.members
    ]

    assert sorted(seen) == sorted(finding.finding_id for finding in findings)
    assert len(seen) == len(set(seen))


def test_single_group_series_is_never_promoted() -> None:
    findings = [
        series_oracle_finding("F1", "M137", "recent"),
        series_oracle_finding("F2", "M138", "recent"),
    ]
    anchors = {
        "F1": series_oracle_anchor("M137"),
        "F2": series_oracle_anchor("M138"),
    }
    lens = build_series_review_lens(_groups(findings), anchors)

    assert lens.clusters == ()
    assert len(lens.fallbacks) == 1
    assert lens.fallbacks[0].whole_group
    assert lens.fallbacks[0].slice_id == lens.fallbacks[0].group_id


def test_unanchored_members_stay_in_one_fallback_remainder() -> None:
    findings, anchors = series_oracle()
    findings.append(
        Finding(
            finding_id="F20",
            artifact="excel",
            finding_class=FindingClass.VALUE_CHANGED,
            severity=Severity.CRITICAL,
            subtype=FindingSubtype.VALUE_REPLACEMENT,
            materiality=Materiality.MATERIAL,
            temporal_context=FindingTemporalContext.HISTORICAL,
            sheet=SHEET,
            location="Z9",
            baseline_location="Z9",
            baseline_value="1",
            current_value="2",
            message="Ops!Z9: historical value changed",
        )
    )
    lens = build_series_review_lens(_groups(findings), anchors)

    assert len(lens.clusters) == 4
    assert len(lens.fallbacks) == 1
    remainder = lens.fallbacks[0]
    assert not remainder.whole_group
    assert [member.finding_id for member in remainder.members] == ["F20"]
    assert remainder.slice_id != remainder.group_id


def test_package_members_never_share_a_parent() -> None:
    findings = [
        series_oracle_finding("F1", "M132", "historical", artifact_member="core"),
        series_oracle_finding("F2", "M137", "recent", artifact_member="core"),
        series_oracle_finding("F3", "M132", "historical", artifact_member="ops"),
        series_oracle_finding("F4", "M137", "recent", artifact_member="ops"),
    ]
    anchors = {
        "F1": series_oracle_anchor("M132"),
        "F2": series_oracle_anchor("M137"),
        "F3": series_oracle_anchor("M132"),
        "F4": series_oracle_anchor("M137"),
    }
    lens = build_series_review_lens(_groups(findings), anchors)

    assert len(lens.clusters) == 2
    assert {cluster.artifact_member for cluster in lens.clusters} == {"core", "ops"}
    for cluster in lens.clusters:
        assert {member.artifact_member for member in cluster.members} == {
            cluster.artifact_member
        }


def test_tampered_anchor_is_ignored_by_the_lens() -> None:
    findings, anchors = series_oracle()
    anchors["F1"] = anchors["F1"].model_copy(update={"period_index": 999})
    lens = build_series_review_lens(_groups(findings), anchors)

    promoted = {
        member.finding_id
        for cluster in lens.clusters
        for member in cluster.members
    }
    assert "F1" not in promoted
    assert sum(slice_.member_count for slice_ in lens.fallbacks) == 1


# --- same-severity confirmation ----------------------------------------------


def test_confirmation_preserves_each_findings_current_severity() -> None:
    findings, anchors = series_oracle()
    lens = build_series_review_lens(_groups(findings), anchors)
    cluster = next(item for item in lens.clusters if item.key[4] == 13)

    updates = cluster_confirmation_updates(cluster.slices, "")

    assert len(updates) == 4
    assert {update.severity for update in updates} == {"critical", "warning"}
    by_id = {update.finding_id: update for update in updates}
    for member in cluster.members:
        assert member.severity is not None
        assert by_id[member.finding_id].severity == member.severity.value
        assert by_id[member.finding_id].comment == ""
        # generation is pure: no finding is mutated
        assert not member.severity_overridden


def test_confirmation_skips_already_reviewed_and_adds_one_shared_comment() -> None:
    findings, anchors = series_oracle()
    lens = build_series_review_lens(_groups(findings), anchors)
    cluster = next(item for item in lens.clusters if item.key[4] == 13)
    cluster.members[0].analyst_comment = "already looked at"

    updates = cluster_confirmation_updates(cluster.slices, "  confirmed as shown  ")

    assert [update.finding_id for update in updates] == [
        member.finding_id for member in cluster.members[1:]
    ]
    assert {update.comment for update in updates} == {"confirmed as shown"}


def test_confirmation_only_touches_visible_slices() -> None:
    findings, anchors = series_oracle()
    lens = build_series_review_lens(_groups(findings), anchors)
    cluster = next(item for item in lens.clusters if item.key[4] == 13)
    visible = tuple(
        slice_
        for slice_ in cluster.slices
        if slice_.members[0].severity is Severity.WARNING
    )

    updates = cluster_confirmation_updates(visible, "")

    assert [update.finding_id for update in updates] == [
        member.finding_id for slice_ in visible for member in slice_.members
    ]
    assert {update.severity for update in updates} == {"warning"}


# --- new-period segments ------------------------------------------------------


def test_v1_anchors_read_as_restatements_and_keep_their_digest() -> None:
    legacy = series_oracle_anchor_v1("M137")

    assert legacy.version == 1
    assert not hasattr(legacy, "segment")
    assert anchor_segment(legacy) == "restatement"
    # a stored V1 payload must keep the digest it was written with
    assert canonical_series_anchor_digest(
        "F1", "primary", "M137", legacy
    ) == canonical_series_anchor_digest("F1", "primary", "M137", legacy)
    assert canonical_series_anchor_digest(
        "F1", "primary", "M137", legacy
    ) != canonical_series_anchor_digest(
        "F1", "primary", "M137", series_oracle_anchor("M137")
    )


def test_v2_carries_an_explicit_segment() -> None:
    restatement = series_oracle_anchor("M137")
    new_period = series_oracle_anchor("M139", segment="new_period")

    assert restatement.version == 2
    assert anchor_segment(restatement) == "restatement"
    assert anchor_segment(new_period) == "new_period"
    assert canonical_series_anchor_digest(
        "F1", "primary", "M137", restatement
    ) != canonical_series_anchor_digest(
        "F1", "primary", "M137", restatement.model_copy(update={"segment": "new_period"})
    )


def test_added_population_is_eligible_without_materiality_or_temporal() -> None:
    finding = series_oracle_finding("F1", "M139", "new_period")

    assert finding.materiality is None
    assert finding.temporal_context is None
    assert series_anchor_segment(finding) == "new_period"
    assert series_anchor_eligible(finding)
    assert anchor_matches_finding(
        finding, series_oracle_anchor("M139", segment="new_period")
    )


def test_segment_must_agree_with_the_finding_subtype() -> None:
    replacement = series_oracle_finding("F1", "M137", "recent")
    addition = series_oracle_finding("F2", "M139", "new_period")

    assert not anchor_matches_finding(
        replacement, series_oracle_anchor("M137", segment="new_period")
    )
    assert not anchor_matches_finding(
        addition, series_oracle_anchor("M139", segment="restatement")
    )
    assert not anchor_matches_finding(addition, series_oracle_anchor_v1("M139"))


def test_only_value_changed_subtypes_may_claim_a_segment() -> None:
    # 2026-08-08: cleared_population joined the lens by user decision, so the
    # remaining exclusions are non-value classes and axis events.
    formula = series_oracle_finding("F1", "M137", "recent").model_copy(
        update={"finding_class": FindingClass.FORMULA_LOGIC_CHANGED}
    )
    axis = series_oracle_finding("F2", "M137", "recent").model_copy(
        update={"subtype": FindingSubtype.AXIS_KEY_REPLACEMENT}
    )

    assert series_anchor_segment(formula) is None
    assert not series_anchor_eligible(formula)
    assert series_anchor_segment(axis) is None
    assert not series_anchor_eligible(axis)


def test_new_period_joins_its_column_parent() -> None:
    findings, anchors = series_oracle(with_new_period=True)
    lens = build_series_review_lens(_groups(findings), anchors)

    assert len(findings) == 21
    assert len(_groups(findings)) == 4
    assert len(lens.clusters) == 4
    assert lens.fallbacks == ()

    by_column = {cluster.key[4]: cluster for cluster in lens.clusters}
    assert [len(by_column[c].slices) for c in (3, 4, 12, 13)] == [2, 2, 4, 4]
    assert [by_column[c].member_count for c in (3, 4, 12, 13)] == [5, 6, 5, 5]
    assert sum(cluster.member_count for cluster in lens.clusters) == 21

    m_series = by_column[13]
    assert [tuple(m.location for m in s.members) for s in m_series.slices] == [
        ("M132", "M136"),
        ("M137",),
        ("M138",),
        ("M139",),
    ]
    assert m_series.slices[-1].members[0].subtype is (
        FindingSubtype.VALUE_ADDED_POPULATION
    )


def test_child_slices_read_chronologically() -> None:
    findings, anchors = series_oracle(with_new_period=True)
    lens = build_series_review_lens(_groups(findings), anchors)

    for cluster in lens.clusters:
        positions = [
            min(anchors[m.finding_id].period_index for m in slice_.members)
            for slice_ in cluster.slices
        ]
        assert positions == sorted(positions)


def test_shipped_restatement_ordering_is_unchanged() -> None:
    findings, anchors = series_oracle()
    lens = build_series_review_lens(_groups(findings), anchors)
    cluster = next(item for item in lens.clusters if item.key[4] == 13)

    assert [tuple(m.location for m in s.members) for s in cluster.slices] == [
        ("M132", "M136"),
        ("M137",),
        ("M138",),
    ]


def test_label_reports_affected_periods_only_when_a_new_period_is_present() -> None:
    findings, anchors = series_oracle(with_new_period=True)
    lens = build_series_review_lens(_groups(findings), anchors)
    by_column = {cluster.key[4]: cluster for cluster in lens.clusters}

    assert by_column[13].label == "Ops · column M · 5 periods affected"
    assert by_column[12].label == "Ops · column L · 5 periods affected"
    # a column with no new period keeps today's exact wording
    assert by_column[3].label == "Ops · column C · 5 changed periods"
    assert by_column[4].label == "Ops · column D · 6 changed periods"


def test_a_lone_new_period_is_not_promoted() -> None:
    findings = [
        series_oracle_finding("F1", "M139", "new_period"),
        series_oracle_finding("F2", "L139", "new_period"),
    ]
    anchors: dict[str, SeriesAnchor] = {
        "F1": series_oracle_anchor("M139", segment="new_period"),
        "F2": series_oracle_anchor("L139", segment="new_period"),
    }
    lens = build_series_review_lens(_groups(findings), anchors)

    assert lens.clusters == ()
    assert len(lens.fallbacks) == 1
    assert lens.fallbacks[0].whole_group


def test_new_period_confirmation_keeps_its_own_severity() -> None:
    findings, anchors = series_oracle(with_new_period=True)
    lens = build_series_review_lens(_groups(findings), anchors)
    cluster = next(item for item in lens.clusters if item.key[4] == 13)

    updates = cluster_confirmation_updates(cluster.slices, "")

    by_id = {update.finding_id: update for update in updates}
    for member in cluster.members:
        assert member.severity is not None
        assert by_id[member.finding_id].severity == member.severity.value
    assert {update.severity for update in updates} == {"critical", "warning"}


# --- cleared-period segments ----------------------------------------------------


def test_cleared_population_maps_to_the_cleared_segment() -> None:
    cleared = series_oracle_finding("F1", "M140", "cleared")

    assert cleared.materiality is None
    assert cleared.temporal_context is None
    assert series_anchor_segment(cleared) == "cleared_period"
    assert series_anchor_eligible(cleared)
    assert anchor_matches_finding(
        cleared, series_oracle_anchor("M140", segment="cleared_period")
    )


def test_cleared_segment_must_agree_with_the_subtype() -> None:
    cleared = series_oracle_finding("F1", "M140", "cleared")
    replacement = series_oracle_finding("F2", "M137", "recent")
    addition = series_oracle_finding("F3", "M139", "new_period")

    assert not anchor_matches_finding(
        cleared, series_oracle_anchor("M140", segment="restatement")
    )
    assert not anchor_matches_finding(
        cleared, series_oracle_anchor("M140", segment="new_period")
    )
    assert not anchor_matches_finding(cleared, series_oracle_anchor_v1("M140"))
    assert not anchor_matches_finding(
        replacement, series_oracle_anchor("M137", segment="cleared_period")
    )
    assert not anchor_matches_finding(
        addition, series_oracle_anchor("M139", segment="cleared_period")
    )


def test_cleared_period_joins_its_column_parent_chronologically() -> None:
    findings, anchors = series_oracle(with_new_period=True, with_cleared=True)
    lens = build_series_review_lens(_groups(findings), anchors)

    assert len(findings) == 23
    m_series = next(item for item in lens.clusters if item.key[4] == 13)
    assert [tuple(m.location for m in s.members) for s in m_series.slices] == [
        ("M132", "M136"),
        ("M137",),
        ("M138",),
        ("M139",),
        ("M140",),
    ]
    assert m_series.slices[-1].members[0].subtype is (
        FindingSubtype.VALUE_CLEARED_POPULATION
    )
    assert m_series.label == "Ops · column M · 6 periods affected"

    seen = [
        member.finding_id
        for cluster in lens.clusters
        for child in cluster.slices
        for member in child.members
    ] + [m.finding_id for item in lens.fallbacks for m in item.members]
    assert sorted(seen) == sorted(f.finding_id for f in findings)


def test_cleared_only_cluster_label_reads_periods_affected() -> None:
    findings = [
        series_oracle_finding("F1", "M132", "historical"),
        series_oracle_finding("F2", "M140", "cleared"),
    ]
    anchors: dict[str, SeriesAnchor] = {
        "F1": series_oracle_anchor("M132"),
        "F2": series_oracle_anchor("M140", segment="cleared_period"),
    }
    lens = build_series_review_lens(_groups(findings), anchors)

    assert len(lens.clusters) == 1
    assert lens.clusters[0].label == "Ops · column M · 2 periods affected"


def test_row_wipe_without_anchors_stays_one_cross_column_row() -> None:
    # A wiped period row is one event: its cleared cells earn no anchor and the
    # canonical cross-column decision must render exactly as today.
    findings, anchors = series_oracle()
    wiped = [
        series_oracle_finding("F90", "L141", "cleared"),
        series_oracle_finding("F91", "M141", "cleared"),
    ]
    findings.extend(wiped)
    lens = build_series_review_lens(_groups(findings), anchors)

    assert len(lens.clusters) == 4
    remainder = [
        item
        for item in lens.fallbacks
        if {m.finding_id for m in item.members} == {"F90", "F91"}
    ]
    assert len(remainder) == 1
    assert remainder[0].whole_group


def test_cleared_confirmation_keeps_its_own_severity() -> None:
    findings, anchors = series_oracle(with_cleared=True)
    lens = build_series_review_lens(_groups(findings), anchors)
    cluster = next(item for item in lens.clusters if item.key[4] == 13)

    updates = cluster_confirmation_updates(cluster.slices, "")

    by_id = {update.finding_id: update for update in updates}
    cleared_member = next(
        m
        for m in cluster.members
        if m.subtype is FindingSubtype.VALUE_CLEARED_POPULATION
    )
    assert by_id[cleared_member.finding_id].severity == "critical"
    for member in cluster.members:
        assert member.severity is not None
        assert by_id[member.finding_id].severity == member.severity.value
