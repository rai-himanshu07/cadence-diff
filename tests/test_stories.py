"""Step 7: change-story engine — evidence-linked, severity-neutral partition."""

from __future__ import annotations

from collections import Counter

from qc_tool.config.profile import DeliverableProfile
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingEvidenceTag,
    FindingProvenance,
    FindingSubtype,
    Materiality,
    Severity,
)
from qc_tool.story import (
    ChangeStory,
    StoryKind,
    annotate_story_evidence,
    build_stories,
)
from qc_tool.triage.rules import triage


def _story(stories: list[ChangeStory], kind: StoryKind) -> ChangeStory:
    matches = [story for story in stories if story.kind is kind]
    assert len(matches) == 1, f"expected one {kind} story, got {len(matches)}"
    return matches[0]


def _table_driver() -> Finding:
    return Finding(
        artifact="excel",
        finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
        sheet="Setup",
        element="tbl_Setup",
        subtype=FindingSubtype.OBJECT_COLUMNS_CHANGED,
        message="Excel table 'tbl_Setup' columns added, removed, reordered, or renamed",
    )


def _named_range_driver() -> Finding:
    return Finding(
        artifact="excel",
        finding_class=FindingClass.NAMED_RANGE_CHANGED,
        element="SetupEnabledList",
        current_value="tbl_Setup[Enabled]",
        message="named range 'SetupEnabledList' added -> tbl_Setup[Enabled]",
    )


def _dv_driver() -> Finding:
    return Finding(
        artifact="excel",
        finding_class=FindingClass.DATA_VALIDATION_CHANGED,
        sheet="Setup",
        subtype=FindingSubtype.OBJECT_CONDITION_CHANGED,
        current_value="tbl_Setup[Enabled]",
        message="data-validation list rule 3 criteria changed",
    )


def _wrapped_formula(location: str, *, sheet: str = "Backend") -> Finding:
    guard = "=IF(XLOOKUP($H$1,tbl_Setup[Brand],tbl_Setup[Enabled],FALSE),IF(B1,C1,NA()),NA())"
    return Finding(
        artifact="excel",
        finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
        subtype=FindingSubtype.FORMULA_WRAPPED,
        event_key="formula-wrapper:wrapped:abc123def456",
        sheet=sheet,
        location=location,
        baseline_location=location,
        baseline_value="=IF(B1,C1,NA())",
        current_value=guard,
        message=(
            f"{sheet}!{location}: formula logic changed "
            "(existing logic preserved inside a new wrapper)"
        ),
    )


def _plain_logic_change(location: str) -> Finding:
    return Finding(
        artifact="excel",
        finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
        sheet="Other",
        location=location,
        baseline_location=location,
        baseline_value="=A1*2",
        current_value="=A1*3",
        message=f"Other!{location}: formula logic changed",
    )


def _value(location: str, tier: Materiality | None, **kwargs: object) -> Finding:
    finding = Finding.model_validate(
        {
            "artifact": "excel",
            "finding_class": FindingClass.VALUE_CHANGED,
            "subtype": FindingSubtype.VALUE_REPLACEMENT,
            "sheet": "Data",
            "location": location,
            "baseline_location": location,
            "baseline_value": "1",
            "current_value": "2",
            "message": f"Data!{location}: historical value changed",
            **kwargs,
        }
    )
    finding.materiality = tier
    return finding


class TestStoryPartition:
    def _build(self) -> tuple[list[Finding], list[ChangeStory]]:
        findings = triage(
            [
                _table_driver(),
                _named_range_driver(),
                _dv_driver(),
                _wrapped_formula("M10"),
                _wrapped_formula("M11"),
                _wrapped_formula("N10"),
                Finding(  # co-located format change joins the rollout story
                    artifact="excel",
                    finding_class=FindingClass.NUMBER_FORMAT_CHANGED,
                    sheet="Backend",
                    location="M10",
                    baseline_location="M10",
                    message="Backend!M10: number format changed",
                ),
                _plain_logic_change("Z9"),
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.ROW_KEY_CHANGED,
                    subtype=FindingSubtype.AXIS_KEY_DERIVED_LABEL,
                    sheet="Summary",
                    location="row 9",
                    baseline_location="row 9",
                    baseline_value="Old",
                    current_value="New",
                    message=(
                        "Summary: historical key at row 9 replaced in place "
                        "[label is formula-derived (=INDEX(x)); the upstream "
                        "driver changed, not this position]"
                    ),
                ),
                _value("B2", Materiality.NOISE),
                _value("B3", Materiality.NOISE),
                _value("B4", Materiality.RECENT_RESTATEMENT),
                _value("B5", Materiality.WITHIN_TOLERANCE),
                _value("B6", Materiality.MATERIAL),
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.FORMULA_ERROR,
                    provenance=FindingProvenance.INHERITED,
                    sheet="Backend",
                    location="I26",
                    current_value="#N/A",
                    message="Backend!I26: error value #N/A (inherited from the baseline)",
                ),
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.FORMULA_INCONSISTENT,
                    provenance=FindingProvenance.HISTORICAL_PATTERN,
                    sheet="Data",
                    location="C9",
                    baseline_value="=RC[-1]",
                    current_value="=RC[-1]*2",
                    message="Data!C9: formula deviates",
                ),
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.ROW_GROWTH,
                    expected_growth=True,
                    sheet="Data",
                    location="row 30",
                    message="Data: new-cycle row 30 appended",
                ),
            ]
        )
        return findings, build_stories(findings)

    def test_partition_is_complete_and_disjoint(self) -> None:
        findings, stories = self._build()
        all_ids = [fid for story in stories for fid in story.finding_ids]
        assert len(all_ids) == len(findings)
        assert len(set(all_ids)) == len(all_ids)
        assert sum(story.member_count for story in stories) == len(findings)

    def test_structure_story_links_drivers_formulas_and_colocated_format(self) -> None:
        findings, stories = self._build()
        structure = _story(stories, StoryKind.STRUCTURE_DRIVER)

        by_id = {finding.finding_id: finding for finding in findings}
        members = [by_id[fid] for fid in structure.finding_ids]
        classes = Counter(member.finding_class for member in members)
        assert classes[FindingClass.TABLE_STRUCTURE_CHANGED] == 1
        assert classes[FindingClass.NAMED_RANGE_CHANGED] == 1
        assert classes[FindingClass.DATA_VALIDATION_CHANGED] == 1
        assert classes[FindingClass.FORMULA_LOGIC_CHANGED] == 3
        assert classes[FindingClass.NUMBER_FORMAT_CHANGED] == 1
        assert any("driver:" in line for line in structure.evidence)
        assert "tbl_Setup" in structure.title or "SetupEnabledList" in structure.title
        formula_members = [
            member
            for member in members
            if member.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
        ]
        assert all(
            {
                FindingEvidenceTag.ADDED_REFERENCE,
                FindingEvidenceTag.RESOLVED_DRIVER,
            }
            <= member.evidence_tags
            for member in formula_members
        )
        colocated = next(
            member
            for member in members
            if member.finding_class is FindingClass.NUMBER_FORMAT_CHANGED
        )
        assert FindingEvidenceTag.EXACT_COLOCATION in colocated.evidence_tags
        formula_at_same_cell = next(
            member
            for member in formula_members
            if member.sheet == colocated.sheet and member.location == colocated.location
        )
        assert FindingEvidenceTag.EXACT_COLOCATION in formula_at_same_cell.evidence_tags

    def test_unlinked_findings_land_in_residual(self) -> None:
        findings, stories = self._build()
        residual = _story(stories, StoryKind.RESIDUAL)
        by_id = {finding.finding_id: finding for finding in findings}
        members = [by_id[fid] for fid in residual.finding_ids]
        assert {member.location for member in members} == {"Z9", "B6"}

    def test_kind_membership(self) -> None:
        findings, stories = self._build()
        by_id = {finding.finding_id: finding for finding in findings}

        noise = _story(stories, StoryKind.REPRESENTATION_NOISE)
        assert {by_id[fid].location for fid in noise.finding_ids} == {"B2", "B3"}

        refresh = _story(stories, StoryKind.DATA_REFRESH)
        assert {by_id[fid].location for fid in refresh.finding_ids} == {
            "B4",
            "row 30",
        }

        accepted = _story(stories, StoryKind.ACCEPTED_DIFFERENCES)
        assert {by_id[fid].location for fid in accepted.finding_ids} == {"B5"}

        inherited = _story(stories, StoryKind.INHERITED)
        assert {by_id[fid].location for fid in inherited.finding_ids} == {"I26", "C9"}

        derived = _story(stories, StoryKind.DERIVED_LABELS)
        assert {by_id[fid].location for fid in derived.finding_ids} == {"row 9"}

    def test_stories_are_severity_neutral(self) -> None:
        findings, _ = self._build()
        before = [(finding.finding_id, finding.severity) for finding in findings]
        build_stories(findings)
        after = [(finding.finding_id, finding.severity) for finding in findings]
        assert before == after

    def test_severity_counts_reflect_members(self) -> None:
        _, stories = self._build()
        noise = _story(stories, StoryKind.REPRESENTATION_NOISE)
        assert noise.severity_counts[Severity.INFO.value] == 2
        assert noise.severity_counts[Severity.CRITICAL.value] == 0

    def test_deterministic_ids_and_order(self) -> None:
        _, first = self._build()
        _, second = self._build()
        assert [story.story_id for story in first] == [story.story_id for story in second]
        assert [story.kind for story in first] == [story.kind for story in second]
        kinds = [story.kind for story in first]
        assert kinds.index(StoryKind.STRUCTURE_DRIVER) == 0
        assert kinds[-1] is StoryKind.RESIDUAL


class TestWrapperWithoutDriver:
    def test_wrapper_rollout_forms_its_own_structure_story(self) -> None:
        findings = triage(
            [
                _wrapped_formula("M10"),
                _wrapped_formula("M11"),
                _value("B6", Materiality.MATERIAL),
            ]
        )
        stories = build_stories(findings)
        structure = _story(stories, StoryKind.STRUCTURE_DRIVER)
        assert structure.member_count == 2
        assert "no structural driver was identified" in structure.description

    def test_two_skeletons_form_two_stories(self) -> None:
        second = _wrapped_formula("Q4")
        second.event_key = "formula-wrapper:wrapped:fff000fff000"
        findings = triage([_wrapped_formula("M10"), second])
        stories = build_stories(findings)
        structure_stories = [
            story for story in stories if story.kind is StoryKind.STRUCTURE_DRIVER
        ]
        assert len(structure_stories) == 2


class TestExplicitEvidenceStories:
    def test_shared_structural_event_merges_without_proximity(self) -> None:
        findings = triage(
            [
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
                    event_key="excel:Data:table:T",
                    sheet="Data",
                    element="T",
                    message="table range changed",
                ),
                Finding(
                    artifact="excel",
                    finding_class=FindingClass.TABLE_STRUCTURE_CHANGED,
                    event_key="excel:Data:table:T",
                    sheet="Other",
                    element="unrelated display label",
                    message="table columns changed",
                ),
            ]
        )

        stories = build_stories(findings)

        assert len(stories) == 1
        assert stories[0].kind is StoryKind.STRUCTURE_DRIVER
        assert stories[0].member_count == 2

    def test_accepted_difference_is_separate_and_severity_neutral(self) -> None:
        findings = triage(
            [
                _value("B2", Materiality.WITHIN_TOLERANCE),
                _value("B3", Materiality.RECENT_RESTATEMENT),
            ]
        )
        before = [(finding.finding_id, finding.severity) for finding in findings]

        stories = build_stories(findings)

        accepted = _story(stories, StoryKind.ACCEPTED_DIFFERENCES)
        refresh = _story(stories, StoryKind.DATA_REFRESH)
        accepted_finding = next(
            finding
            for finding in findings
            if finding.materiality is Materiality.WITHIN_TOLERANCE
        )
        refresh_finding = next(
            finding
            for finding in findings
            if finding.materiality is Materiality.RECENT_RESTATEMENT
        )
        assert accepted.finding_ids == (accepted_finding.finding_id,)
        assert refresh.finding_ids == (refresh_finding.finding_id,)
        assert [(finding.finding_id, finding.severity) for finding in findings] == before

    def test_waived_structural_driver_does_not_stamp_resolved_driver(self) -> None:
        driver = _table_driver()
        formula = _wrapped_formula("M10")
        profile = DeliverableProfile.model_validate(
            {
                "name": "waived-driver",
                "waivers": [
                    {
                        "finding_class": "table_structure_changed",
                        "sheet": "Setup",
                        "element": "tbl_Setup",
                        "reason": "approved setup change",
                        "expires": "2999-12-31",
                    }
                ],
            }
        )

        findings = triage([driver, formula], profile)
        annotate_story_evidence(findings)

        assert driver.expected_growth
        assert FindingEvidenceTag.RESOLVED_DRIVER not in formula.evidence_tags

        unwaived_driver = _table_driver()
        unwaived_formula = _wrapped_formula("M10")
        unwaived = triage([unwaived_driver, unwaived_formula])
        annotate_story_evidence(unwaived)
        assert FindingEvidenceTag.RESOLVED_DRIVER in unwaived_formula.evidence_tags


class TestErrorPopulationStories:
    @staticmethod
    def _population(location: str, event_key: str) -> Finding:
        return Finding(
            artifact="excel",
            finding_class=FindingClass.FORMULA_ERROR,
            subtype=FindingSubtype.COLUMNAR_ERROR_POPULATION,
            event_key=event_key,
            evidence_tags={
                FindingEvidenceTag.CACHED_VALUE_ONLY,
                FindingEvidenceTag.SPARSE_MASS_POPULATION,
            },
            sheet="Data",
            location=location,
            current_value="#N/A",
            message=f"Data!{location}: sparse cached error population",
        )

    def test_population_event_keys_partition_deterministically(self) -> None:
        findings = triage(
            [
                self._population("B2", "error-population:aaaaaaaaaaaa"),
                self._population("B8", "error-population:aaaaaaaaaaaa"),
                self._population("C2", "error-population:bbbbbbbbbbbb"),
            ]
        )
        before = [(finding.finding_id, finding.severity) for finding in findings]

        first = build_stories(findings)
        second = build_stories(list(reversed(findings)))
        populations = [
            story for story in first if story.kind is StoryKind.ERROR_POPULATION
        ]

        assert sorted(story.member_count for story in populations) == [1, 2]
        assert sum(story.member_count for story in first) == len(findings)
        assert len({fid for story in first for fid in story.finding_ids}) == len(findings)
        assert sorted(story.story_id for story in first) == sorted(
            story.story_id for story in second
        )
        assert [(finding.finding_id, finding.severity) for finding in findings] == before
        assert all("does not imply reduced risk" in story.description for story in populations)


def test_empty_input_yields_no_stories() -> None:
    assert build_stories([]) == []
