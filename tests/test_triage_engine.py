"""Severity triage and engine orchestration tests (criterion 10)."""

from pathlib import Path

import pytest

import qc_tool.engine as engine_module
from qc_tool.config.profile import DeliverableProfile
from qc_tool.coverage import CoverageItem, CoverageState
from qc_tool.engine import QCRunResult, _apply_findings_budget, run_qc
from qc_tool.findings import Finding, FindingClass, Severity, limit_findings
from qc_tool.triage.rules import triage

FIXTURE_PROFILE = DeliverableProfile.model_validate(
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


@pytest.fixture(scope="module")
def result(fixture_dir: Path) -> QCRunResult:
    return run_qc(
        baseline_excel=fixture_dir / "baseline.xlsx",
        current_excel=fixture_dir / "current.xlsx",
        baseline_ppt=fixture_dir / "baseline.pptx",
        current_ppt=fixture_dir / "current.pptx",
        profile=FIXTURE_PROFILE,
    )


def _one(result: QCRunResult, cls: FindingClass, **attrs: str) -> Finding:
    matches = [
        f
        for f in result.findings
        if f.finding_class is cls
        and all(getattr(f, key) == value for key, value in attrs.items())
    ]
    assert len(matches) == 1, f"expected exactly one {cls} with {attrs}, got {len(matches)}"
    return matches[0]


def test_default_severities(result: QCRunResult) -> None:
    assert _one(
        result, FindingClass.VALUE_CHANGED, sheet="Long_Monthly", location="C7"
    ).severity is (Severity.CRITICAL)
    assert _one(result, FindingClass.FORMULA_HARDCODED).severity is Severity.CRITICAL
    assert _one(result, FindingClass.CROSSCHECK_MISMATCH).severity is Severity.CRITICAL
    assert _one(result, FindingClass.NUMBER_FORMAT_CHANGED).severity is Severity.WARNING
    assert _one(result, FindingClass.STYLE_CHANGED).severity is Severity.INFO
    assert _one(result, FindingClass.SLIDE_REORDERED).severity is Severity.EXPECTED


def test_ppt_media_change_defaults_to_warning() -> None:
    findings = triage(
        [
            Finding(
                artifact="ppt",
                finding_class=FindingClass.PPT_MEDIA_CHANGED,
                slide="Executive Summary",
                message="embedded media bytes changed",
            )
        ]
    )

    assert findings[0].severity is Severity.WARNING


def test_expected_growth_always_expected(result: QCRunResult) -> None:
    growth = [f for f in result.findings if f.expected_growth]
    assert growth
    assert all(f.severity is Severity.EXPECTED for f in growth)


def test_ordering_and_ids(result: QCRunResult) -> None:
    ids = [f.finding_id for f in result.findings]
    assert ids == [f"F{i:04d}" for i in range(1, len(ids) + 1)]
    ranks = [
        [Severity.CRITICAL, Severity.WARNING, Severity.INFO, Severity.EXPECTED].index(
            f.severity or Severity.WARNING
        )
        for f in result.findings
    ]
    assert ranks == sorted(ranks)  # criticals first, expected last


def test_crosscheck_integrated(result: QCRunResult) -> None:
    assert result.verified_crosschecks == 1  # Total revenue verifies
    mismatch = _one(result, FindingClass.CROSSCHECK_MISMATCH)
    assert mismatch.element == "Margin"


def test_profile_severity_override(result: QCRunResult) -> None:
    override_profile = DeliverableProfile.model_validate(
        {"name": "strict", "severity": {"style_changed": "critical"}}
    )
    findings = [f.model_copy(deep=True) for f in result.findings]
    retriaged = triage(findings, override_profile)
    style = next(f for f in retriaged if f.finding_class is FindingClass.STYLE_CHANGED)
    assert style.severity is Severity.CRITICAL
    # Expected findings never escalate through class overrides.
    growth = next(f for f in retriaged if f.expected_growth)
    assert growth.severity is Severity.EXPECTED


def test_findings_budget_discloses_omissions_and_degrades_coverage(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    findings = [
        Finding(
            artifact="excel",
            finding_class=FindingClass.VALUE_CHANGED,
            sheet="Data",
            location=f"A{index}",
            message=f"changed {index}",
        )
        for index in range(1, 6)
    ]
    coverage = [
        CoverageItem(
            check_id="excel-values",
            label="Excel values",
            artifact="excel",
            state=CoverageState.CHECKED,
        )
    ]
    monkeypatch.setattr(
        engine_module,
        "limit_findings",
        lambda items: limit_findings(
            items,
            max_per_class_scope=2,
            max_total=10,
        ),
    )

    bounded = _apply_findings_budget(findings, coverage)
    finalized = triage(bounded)

    assert len(bounded) == 3
    summary = next(
        item
        for item in finalized
        if item.finding_class is FindingClass.FINDINGS_CAPPED
    )
    assert summary.severity is Severity.WARNING
    assert summary.current_value == "2 retained; 3 omitted"
    assert coverage[0].state is CoverageState.DEGRADED
    assert "omitted 3 findings" in coverage[0].detail


def test_disclosure_for_xlsb_pair(fixture_dir: Path) -> None:
    result = run_qc(
        baseline_excel=fixture_dir / "baseline.xlsb",
        current_excel=fixture_dir / "current.xlsb",
    )
    assert result.disclosures and "degraded" in result.disclosures[0]
    value_changed = [
        f for f in result.findings if f.finding_class is FindingClass.VALUE_CHANGED
    ]
    # XB01's revenue edit plus its stored derived margin (xlsb keeps values
    # only, so the margin column changes as data, not as a formula result).
    assert {f.location for f in value_changed} == {"C7", "E7"}


def test_engine_input_validation(fixture_dir: Path) -> None:
    with pytest.raises(ValueError, match="both a baseline and a current"):
        run_qc(baseline_excel=fixture_dir / "baseline.xlsx")
    with pytest.raises(ValueError, match="nothing to compare"):
        run_qc()
