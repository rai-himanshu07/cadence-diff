"""Standalone current-deck preflight behavior."""

from pathlib import Path

from qc_tool.config.profile import PptProfile
from qc_tool.coverage import CoverageState, QCRunMode
from qc_tool.engine import run_qc
from qc_tool.findings import FindingClass
from qc_tool.history.store import sha256_file
from qc_tool.ppt.extract import (
    ChartContent,
    DeckSnapshot,
    SlideContent,
    TableContent,
)
from qc_tool.ppt.preflight import preflight_deck


def test_current_ppt_preflight_runs_without_baseline(fixture_dir: Path) -> None:
    source = fixture_dir / "current.pptx"
    before = sha256_file(source)

    result = run_qc(current_ppt=source, mode=QCRunMode.CURRENT_FILE_PREFLIGHT)

    assert result.mode is QCRunMode.CURRENT_FILE_PREFLIGHT
    assert result.files == {"current_ppt": "current.pptx"}
    intrinsic = next(
        item for item in result.coverage if item.check_id == "ppt-slide-inventory"
    )
    assert intrinsic.state is CoverageState.CHECKED
    comparison = next(
        item for item in result.coverage if item.check_id == "ppt-cycle-comparison"
    )
    assert comparison.state is CoverageState.UNAVAILABLE
    assert sha256_file(source) == before


def test_ppt_preflight_detects_intrinsic_defects() -> None:
    deck = DeckSnapshot(
        source_name="current.pptx",
        slides=[
            SlideContent(
                index=0,
                title="Executive Summary",
                texts=["Reporting cycle Jun-26", "DRAFT — replace before issue"],
                tables=[
                    TableContent(
                        rows=[["Region", "Jun-26"], ["North", ""]],
                        shape_id=101,
                    )
                ],
                charts=[
                    ChartContent(
                        chart_type="LINE",
                        categories=["May-26", "Jun-26"],
                        series=[("Revenue", [100.0])],
                        shape_id=102,
                    )
                ],
                shape_count=4,
            ),
            SlideContent(
                index=1,
                title="Executive Summary",
                texts=["Reporting cycle May-26"],
                shape_count=2,
            ),
            SlideContent(index=2, title=None, texts=[], shape_count=0),
        ],
    )
    profile = PptProfile(required_slides=["Appendix"])

    result = preflight_deck(deck, profile)
    classes = {finding.finding_class for finding in result.findings}

    assert {
        FindingClass.PPT_DRAFT_TOKEN,
        FindingClass.PPT_EMPTY_SLIDE,
        FindingClass.PPT_DUPLICATE_TITLE,
        FindingClass.PPT_REQUIRED_SLIDE_MISSING,
        FindingClass.PPT_PERIOD_INCONSISTENT,
        FindingClass.PPT_TABLE_BLANK,
        FindingClass.PPT_CHART_LENGTH_MISMATCH,
    } <= classes
    table = next(
        finding
        for finding in result.findings
        if finding.finding_class is FindingClass.PPT_TABLE_BLANK
    )
    chart = next(
        finding
        for finding in result.findings
        if finding.finding_class is FindingClass.PPT_CHART_LENGTH_MISMATCH
    )
    assert table.focus_shape_id == 101
    assert chart.focus_shape_id == 102
