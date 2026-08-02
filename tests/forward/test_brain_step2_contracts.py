"""Step 2 contracts for orthogonal finding evidence axes."""

from __future__ import annotations

import ast
from pathlib import Path

import pytest
from pydantic import ValidationError

from qc_tool.engine import QCRunResult
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingEvidenceTag,
    FindingExpectedReason,
    FindingTemporalContext,
    Materiality,
    Severity,
)
from qc_tool.triage.rules import DEFAULT_SEVERITIES, assign_severity

_ROOT = Path(__file__).resolve().parents[2]


def _finding(**updates: object) -> Finding:
    return Finding.model_validate(
        {
            "artifact": "excel",
            "finding_class": FindingClass.VALUE_CHANGED,
            "sheet": "Data",
            "location": "B2",
            "message": "value changed",
            **updates,
        }
    )


def test_step2_enum_inventories_are_closed() -> None:
    assert {item.value for item in FindingTemporalContext} == {
        "current_period",
        "recent_window",
        "historical",
    }
    assert {item.value for item in FindingExpectedReason} == {
        "period_progression",
        "cadence_extension",
        "rolling_window",
        "profile_refresh",
        "figure_refresh",
        "presentation_reorder",
        "waiver",
    }
    assert {item.value for item in FindingEvidenceTag} == {
        "display_equivalent",
        "ulp_scale",
        "explicit_na",
        "formula_text",
        "formula_presence",
        "cached_value_only",
        "concentrated_population",
        "contiguous_population",
        "sparse_mass_population",
        "structural_error",
        "exact_wrapper",
        "shape_wrapper",
        "added_reference",
        "resolved_driver",
        "exact_colocation",
    }


def test_expected_reason_derives_compatibility_bool_and_round_trips() -> None:
    finding = _finding(
        expected_reason=FindingExpectedReason.PERIOD_PROGRESSION,
        temporal_context=FindingTemporalContext.CURRENT_PERIOD,
        evidence_tags={
            FindingEvidenceTag.ULP_SCALE,
            FindingEvidenceTag.DISPLAY_EQUIVALENT,
        },
    )

    assert finding.expected_growth is True
    payload = finding.model_dump(mode="json")
    assert payload["expected_reason"] == "period_progression"
    assert payload["temporal_context"] == "current_period"
    assert payload["evidence_tags"] == ["display_equivalent", "ulp_scale"]
    assert Finding.model_validate(payload) == finding

    with pytest.raises(ValidationError):
        _finding(evidence_tags=["unbounded_free_text"])


def test_legacy_expected_and_recent_materiality_rehydrate() -> None:
    finding = _finding(
        expected_growth=True,
        materiality="recent_restatement",
    )

    assert finding.expected_growth is True
    assert finding.expected_reason is None
    assert finding.materiality is Materiality.RECENT_RESTATEMENT
    assert assign_severity(finding) is Severity.EXPECTED


def test_growth_classes_are_warning_without_typed_or_legacy_evidence() -> None:
    finding = _finding(finding_class=FindingClass.ROW_GROWTH)

    assert Severity.EXPECTED not in DEFAULT_SEVERITIES.values()
    assert assign_severity(finding) is Severity.WARNING


def test_magnitude_temporal_and_expected_precedence() -> None:
    accepted_recent = _finding(
        materiality=Materiality.WITHIN_TOLERANCE,
        temporal_context=FindingTemporalContext.RECENT_WINDOW,
    )
    material_recent = _finding(
        materiality=Materiality.MATERIAL,
        temporal_context=FindingTemporalContext.RECENT_WINDOW,
    )
    material_historical = _finding(
        materiality=Materiality.MATERIAL,
        temporal_context=FindingTemporalContext.HISTORICAL,
    )
    expected_historical = _finding(
        materiality=Materiality.MATERIAL,
        temporal_context=FindingTemporalContext.HISTORICAL,
        expected_reason=FindingExpectedReason.PROFILE_REFRESH,
    )

    assert assign_severity(accepted_recent) is Severity.INFO
    assert assign_severity(material_recent) is Severity.WARNING
    assert assign_severity(material_historical) is Severity.CRITICAL
    assert assign_severity(expected_historical) is Severity.EXPECTED


def test_production_finding_producers_never_write_expected_growth() -> None:
    violations: list[str] = []
    for path in sorted((_ROOT / "qc_tool").rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            if not isinstance(node.func, ast.Name) or node.func.id != "Finding":
                continue
            if any(keyword.arg == "expected_growth" for keyword in node.keywords):
                violations.append(f"{path.relative_to(_ROOT)}:{node.lineno}")

    assert violations == []


def test_production_excel_never_emits_legacy_recent_materiality() -> None:
    violations: list[str] = []
    for path in sorted((_ROOT / "qc_tool" / "excel").rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Attribute):
                continue
            if not isinstance(node.value, ast.Name) or node.value.id != "Materiality":
                continue
            if node.attr == "RECENT_RESTATEMENT":
                violations.append(f"{path.relative_to(_ROOT)}:{node.lineno}")

    assert violations == []


def test_full_cycle_expected_findings_have_typed_reasons(
    qc_result: QCRunResult,
) -> None:
    expected = [
        finding
        for finding in qc_result.findings
        if finding.severity is Severity.EXPECTED
    ]

    assert expected
    assert all(finding.expected_reason is not None for finding in expected)
    assert all(finding.expected_growth for finding in expected)
    assert {finding.expected_reason for finding in expected} == {
        FindingExpectedReason.CADENCE_EXTENSION,
        FindingExpectedReason.FIGURE_REFRESH,
        FindingExpectedReason.PERIOD_PROGRESSION,
        FindingExpectedReason.PRESENTATION_REORDER,
        FindingExpectedReason.PROFILE_REFRESH,
        FindingExpectedReason.ROLLING_WINDOW,
    }
