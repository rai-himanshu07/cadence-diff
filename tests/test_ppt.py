"""PPT extraction, fuzzy matching, and diff tests (criteria 7, 8)."""

from pathlib import Path

import pytest

from qc_tool.config.profile import PptProfile
from qc_tool.findings import Finding, FindingClass
from qc_tool.ppt.diff import diff_decks
from qc_tool.ppt.extract import DeckSnapshot, load_deck_snapshot
from qc_tool.ppt.match import SlideMatching, match_slides
from tests.fixtures.manifest_schema import FixtureManifest


@pytest.fixture(scope="module")
def base_deck(fixture_dir: Path) -> DeckSnapshot:
    return load_deck_snapshot(fixture_dir / "baseline.pptx")


@pytest.fixture(scope="module")
def curr_deck(fixture_dir: Path) -> DeckSnapshot:
    return load_deck_snapshot(fixture_dir / "current.pptx")


@pytest.fixture(scope="module")
def matching(base_deck: DeckSnapshot, curr_deck: DeckSnapshot) -> SlideMatching:
    return match_slides(base_deck, curr_deck)


@pytest.fixture(scope="module")
def findings(matching: SlideMatching) -> list[Finding]:
    return diff_decks(matching)


def _by_class(findings: list[Finding], cls: FindingClass) -> list[Finding]:
    return [f for f in findings if f.finding_class is cls]


def test_extraction_shapes(base_deck: DeckSnapshot, curr_deck: DeckSnapshot) -> None:
    assert [s.title for s in base_deck.slides] == [
        "Executive Summary",
        "Revenue by Region",
        "Revenue Trend",
        "Weekly Ops",
        "Notes & Definitions",
        "Deep Dive Archive",
    ]
    trend = next(s for s in curr_deck.slides if s.title == "Revenue Trend")
    assert trend.charts[0].categories[-1] == "Jun-26"
    region = next(s for s in curr_deck.slides if s.title == "Revenue by Region")
    assert region.tables[0].rows[0][0] == "Region"


def test_fuzzy_matching_survives_reorder_add_delete(matching: SlideMatching) -> None:
    pair_titles = {(b.title, c.title) for b, c in matching.pairs}
    assert pair_titles == {
        ("Executive Summary", "Executive Summary"),
        ("Revenue by Region", "Revenue by Region"),
        ("Revenue Trend", "Revenue Trend"),
        ("Weekly Ops", "Weekly Ops"),
        ("Notes & Definitions", "Notes & Definitions"),
    }
    assert [s.title for s in matching.added] == ["New Initiatives"]  # P01
    assert [s.title for s in matching.removed] == ["Deep Dive Archive"]  # P02
    # PX01: only the deliberately moved slide counts as reordered — pure
    # index shifts caused by the insertion do not.
    assert [(b.title, c.index) for b, c in matching.reordered] == [
        ("Notes & Definitions", 1)
    ]


def test_slide_pins_override(base_deck: DeckSnapshot, curr_deck: DeckSnapshot) -> None:
    profile = PptProfile(
        slide_pins={"Deep Dive Archive": "New Initiatives"}, match_threshold=55.0
    )
    pinned = match_slides(base_deck, curr_deck, profile)
    assert ("Deep Dive Archive", "New Initiatives") in {
        (b.title, c.title) for b, c in pinned.pairs
    }
    assert pinned.added == [] and pinned.removed == []


def test_slide_add_remove_reorder_findings(findings: list[Finding]) -> None:
    assert [f.slide for f in _by_class(findings, FindingClass.SLIDE_ADDED)] == [
        "New Initiatives"
    ]
    assert [f.slide for f in _by_class(findings, FindingClass.SLIDE_REMOVED)] == [
        "Deep Dive Archive"
    ]
    reordered = _by_class(findings, FindingClass.SLIDE_REORDERED)
    assert [(f.slide, f.expected_growth) for f in reordered] == [
        ("Notes & Definitions", True)
    ]


def test_text_changes_split_wording_from_figures(
    findings: list[Finding], manifest: FixtureManifest
) -> None:
    texts = _by_class(findings, FindingClass.SLIDE_TEXT_CHANGED)
    unexpected = [f for f in texts if not f.expected_growth]
    p03 = manifest.defect("P03")
    assert [(f.slide, f.baseline_value, f.current_value) for f in unexpected] == [
        ("Executive Summary", p03.baseline, p03.current)
    ]
    # Figure lines (revenue/margin bullets) refresh each cycle: expected.
    expected = [f for f in texts if f.expected_growth]
    assert len(expected) == 2
    assert all(f.slide == "Executive Summary" for f in expected)


def test_table_diff(findings: list[Finding], manifest: FixtureManifest) -> None:
    tables = _by_class(findings, FindingClass.TABLE_VALUE_CHANGED)
    p04 = manifest.defect("P04")
    unexpected = [f for f in tables if not f.expected_growth]
    assert [(f.slide, f.element, f.baseline_value, f.current_value) for f in unexpected] == [
        ("Revenue by Region", p04.element, p04.baseline, p04.current)
    ]
    expected = [f for f in tables if f.expected_growth]
    assert [(f.slide, f.element) for f in expected] == [
        ("Revenue by Region", "Jun-26")
    ]  # PX04


def test_chart_diff_full_history(findings: list[Finding], manifest: FixtureManifest) -> None:
    charts = [
        f
        for f in _by_class(findings, FindingClass.CHART_VALUE_CHANGED)
        if f.slide == "Revenue Trend"
    ]
    p05 = manifest.defect("P05")
    unexpected = [f for f in charts if not f.expected_growth]
    assert len(unexpected) == 1
    finding = unexpected[0]
    assert finding.element == p05.element  # Mar-26
    assert float(finding.baseline_value or "") == float(p05.baseline or "")
    assert float(finding.current_value or "") == float(p05.current or "")
    expected = [f for f in charts if f.expected_growth]
    assert [f.element for f in expected] == ["Jun-26"]  # PX03


def test_chart_diff_rolling_window(findings: list[Finding]) -> None:
    weekly = [
        f
        for f in _by_class(findings, FindingClass.CHART_VALUE_CHANGED)
        if f.slide == "Weekly Ops"
    ]
    # PX02: the whole rolling advance is expected — no unexpected findings.
    assert weekly and all(f.expected_growth for f in weekly)
    window = next(f for f in weekly if f.element == "window")
    assert window.baseline_value == "W17..W20"
    assert window.current_value == "W18..W21"


def test_chart_window_profile_override(matching: SlideMatching) -> None:
    profile = PptProfile(chart_windows={"Weekly Ops": "full"})
    findings = diff_decks(matching, profile)
    weekly = [
        f
        for f in findings
        if f.finding_class is FindingClass.CHART_VALUE_CHANGED and f.slide == "Weekly Ops"
    ]
    removed = [f for f in weekly if not f.expected_growth and "removed" in f.message]
    assert [f.element for f in removed] == ["W17"]  # forced full-history: drop-off flags
