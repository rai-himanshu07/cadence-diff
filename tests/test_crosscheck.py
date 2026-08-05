"""Cross-check tests (acceptance criterion 9): parser, tracing, verification."""

from pathlib import Path

import pytest

from qc_tool.config.profile import CrosscheckProfile
from qc_tool.crosscheck.numbers import (
    display_matches,
    extract_figures,
    parse_figure,
)
from qc_tool.crosscheck.trace import (
    FigureOccurrence,
    build_mapping,
    extract_deck_figures,
    suggest_sources,
    verify_mappings,
)
from qc_tool.findings import FindingClass
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.io.model import WorkbookSnapshot
from qc_tool.ppt.extract import DeckSnapshot, load_deck_snapshot
from tests.fixtures.manifest_schema import FixtureManifest


@pytest.mark.parametrize(
    ("text", "value", "is_percent", "decimals"),
    [
        ("$1.2M", 1_200_000.0, False, 1),
        ("2.60M", 2_600_000.0, False, 2),
        ("12%", 0.12, True, 0),
        ("37.9%", 0.379, True, 1),
        ("(123)", -123.0, False, 0),
        ("(1,234.5)", -1234.5, False, 1),
        ("1,234,567", 1_234_567.0, False, 0),
        ("3.4bn", 3_400_000_000.0, False, 1),
        ("250k", 250_000.0, False, 0),
        ("-42", -42.0, False, 0),
    ],
)
def test_parse_figure(text: str, value: float, is_percent: bool, decimals: int) -> None:
    figure = parse_figure(text)
    assert figure is not None
    assert figure.value == pytest.approx(value)
    assert figure.is_percent is is_percent
    assert figure.decimals == decimals


@pytest.mark.parametrize("text", ["Jan-26", "W05", "Q1 FY26", "North", ""])
def test_period_labels_are_not_figures(text: str) -> None:
    assert extract_figures(text) == []


@pytest.mark.parametrize(
    ("text", "cell_value", "expected"),
    [
        ("$1.2M", 1_234_567.89, True),  # documented equivalence case
        ("$1.2M", 1_260_000.0, False),  # displays as 1.3M
        ("12%", 0.1204, True),  # documented equivalence case
        ("12%", 0.1304, False),  # displays as 13%
        ("(123)", -123.4, True),  # parenthesized negative, 0 decimals
        ("(123)", 123.4, False),  # sign matters
        ("100,000", 100_000.0, True),
        ("100,000", 100_001.0, False),
        ("37.9%", 0.3721, False),  # the X03 drift
    ],
)
def test_display_matches(text: str, cell_value: float, expected: bool) -> None:
    figure = parse_figure(text)
    assert figure is not None
    assert display_matches(figure, cell_value) is expected


@pytest.fixture(scope="module")
def workbook(fixture_dir: Path) -> WorkbookSnapshot:
    return load_workbook_snapshot(fixture_dir / "current.xlsx")


@pytest.fixture(scope="module")
def deck(fixture_dir: Path) -> DeckSnapshot:
    return load_deck_snapshot(fixture_dir / "current.pptx")


@pytest.fixture(scope="module")
def occurrences(deck: DeckSnapshot) -> list[FigureOccurrence]:
    return extract_deck_figures(deck)


def _occurrence(
    occurrences: list[FigureOccurrence], slide: str, skeleton_part: str
) -> FigureOccurrence:
    matches = [
        o for o in occurrences if o.slide == slide and skeleton_part in o.line_skeleton
    ]
    assert matches, f"no figure occurrence for {slide!r} / {skeleton_part!r}"
    return matches[0]


def test_deck_figures_extracted(occurrences: list[FigureOccurrence]) -> None:
    revenue = _occurrence(occurrences, "Executive Summary", "Total revenue")
    assert revenue.figure.scale == 1e6
    margin = _occurrence(occurrences, "Executive Summary", "Margin")
    assert margin.figure.is_percent
    table = _occurrence(occurrences, "Revenue by Region", "table:North/Jan-26")
    assert table.figure.value == pytest.approx(100_000.0)
    assert table.shape_id is not None and table.shape_id > 0
    assert revenue.shape_id is None


def test_auto_suggest_finds_kpi_source(
    occurrences: list[FigureOccurrence], workbook: WorkbookSnapshot
) -> None:
    revenue = _occurrence(occurrences, "Executive Summary", "Total revenue")
    candidates = suggest_sources(revenue, workbook)
    assert candidates, "expected at least one candidate"
    top = candidates[0]
    assert (top.sheet, top.cell) == ("Dashboard", "B2")
    assert top.display_match


def test_auto_suggest_near_miss_for_drifted_figure(
    occurrences: list[FigureOccurrence], workbook: WorkbookSnapshot
) -> None:
    margin = _occurrence(occurrences, "Executive Summary", "Margin")
    candidates = suggest_sources(margin, workbook)
    assert candidates, "the drifted figure should still get near-miss suggestions"
    top = candidates[0]
    assert (top.sheet, top.cell) == ("Dashboard", "B4")
    assert not top.display_match  # X03: it drifts from the workbook


def test_auto_suggest_table_cell_source(
    occurrences: list[FigureOccurrence], workbook: WorkbookSnapshot
) -> None:
    table = _occurrence(occurrences, "Revenue by Region", "table:North/Jan-26")
    candidates = suggest_sources(table, workbook)
    assert (candidates[0].sheet, candidates[0].cell) == ("Long_Monthly", "C2")


def test_verify_mappings_end_to_end(
    occurrences: list[FigureOccurrence],
    deck: DeckSnapshot,
    workbook: WorkbookSnapshot,
    manifest: FixtureManifest,
) -> None:
    revenue = _occurrence(occurrences, "Executive Summary", "Total revenue")
    margin = _occurrence(occurrences, "Executive Summary", "Margin")
    table = _occurrence(occurrences, "Revenue by Region", "table:North/Jan-26")

    profile = CrosscheckProfile(
        mappings=[
            build_mapping(revenue, suggest_sources(revenue, workbook)[0], label="Total revenue"),
            build_mapping(margin, suggest_sources(margin, workbook)[0], label="Margin"),
            build_mapping(table, suggest_sources(table, workbook)[0], label="North / Jan-26"),
        ]
    )
    result = verify_mappings(deck, workbook, profile)

    # Two of three fixture crosscheck entries hold; the margin drifts (X03).
    assert {m.label for m in result.verified} == {"Total revenue", "North / Jan-26"}
    assert len(result.findings) == 1
    mismatch = result.findings[0]
    assert mismatch.finding_class is FindingClass.CROSSCHECK_MISMATCH
    assert mismatch.element == "Margin"
    assert mismatch.location == "Dashboard!B4"
    x03 = manifest.defect("X03")
    assert mismatch.current_value == x03.current.split()[-1] if x03.current else True


def test_verify_unresolved_when_wording_changes(
    deck: DeckSnapshot, workbook: WorkbookSnapshot
) -> None:
    from qc_tool.config.profile import CrosscheckMapping

    profile = CrosscheckProfile(
        mappings=[
            CrosscheckMapping(
                slide="Executive Summary",
                line_skeleton="A line that no longer exists #",
                figure_index=0,
                label="Ghost figure",
                source_sheet="Dashboard",
                source_cell="B2",
            ),
            CrosscheckMapping(
                slide="Executive Summary",
                line_skeleton="Total revenue #",
                figure_index=0,
                label="Formula source",
                source_sheet="Summary",
                source_cell="B2",  # formula cell without cached value
            ),
        ]
    )
    result = verify_mappings(deck, workbook, profile)
    assert [f.finding_class for f in result.findings] == [
        FindingClass.CROSSCHECK_UNRESOLVED,
        FindingClass.CROSSCHECK_UNRESOLVED,
    ]
    assert result.verified == []
