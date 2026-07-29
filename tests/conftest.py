"""Shared fixtures: generate the QC fixture set once per test session."""

from pathlib import Path

import pytest

from qc_tool.config.profile import DeliverableProfile
from qc_tool.engine import QCRunResult, run_qc
from tests.fixtures.generate import generate
from tests.fixtures.manifest_schema import FixtureManifest


@pytest.fixture(scope="session")
def fixture_dir(tmp_path_factory: pytest.TempPathFactory) -> Path:
    dest = tmp_path_factory.mktemp("qc_fixtures")
    generate(dest)
    return dest


@pytest.fixture(scope="session")
def manifest(fixture_dir: Path) -> FixtureManifest:
    return FixtureManifest.model_validate_json(
        (fixture_dir / "manifest.json").read_text(encoding="utf-8")
    )


def fixture_profile() -> DeliverableProfile:
    """The profile that fully configures the synthetic deliverable."""
    return DeliverableProfile.model_validate(
        {
            "name": "fixture",
            "excel": {"sheets": {"Dashboard": {"refresh_ranges": ["B2:B5"]}}},
            "crosscheck": {
                "mappings": [
                    {
                        "slide": "Executive Summary",
                        "line_skeleton": "Total revenue $#M",
                        "figure_index": 0,
                        "label": "Total revenue",
                        "source_sheet": "Dashboard",
                        "source_cell": "B2",
                    },
                    {
                        "slide": "Executive Summary",
                        "line_skeleton": "Margin #",
                        "figure_index": 0,
                        "label": "Margin",
                        "source_sheet": "Dashboard",
                        "source_cell": "B4",
                    },
                ]
            },
        }
    )


@pytest.fixture(scope="session")
def qc_result(fixture_dir: Path) -> QCRunResult:
    """One full-pipeline run over the fixture pair, shared across test modules."""
    return run_qc(
        baseline_excel=fixture_dir / "baseline.xlsx",
        current_excel=fixture_dir / "current.xlsx",
        baseline_ppt=fixture_dir / "baseline.pptx",
        current_ppt=fixture_dir / "current.pptx",
        profile=fixture_profile(),
    )
