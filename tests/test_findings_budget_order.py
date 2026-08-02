"""Step 7: findings budget applied before impact and excerpt enrichment."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from openpyxl import Workbook

from qc_tool import engine as engine_module
from qc_tool.coverage import CoverageItem, CoverageState, QCRunMode
from qc_tool.engine import _apply_findings_budget, run_qc
from qc_tool.excel import context as context_module
from qc_tool.excel import dependency as dependency_module
from qc_tool.findings import Finding, FindingClass, limit_findings


def _capped_workbook(path: Path, *, rows: int, offset: int) -> None:
    workbook = Workbook()
    sheet = workbook.active
    if sheet is None:  # pragma: no cover - openpyxl always creates one sheet
        raise RuntimeError("openpyxl did not create a default worksheet")
    sheet.title = "Data"
    sheet.append(["Key", "Amount", "Derived"])
    for index in range(1, rows + 1):
        sheet.append([f"k{index}", index + offset, f"=B{index + 1}*2"])
    workbook.save(path)


@pytest.fixture
def tiny_budget(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        engine_module,
        "limit_findings",
        lambda items: limit_findings(items, max_per_class_scope=5, max_total=10_000),
    )


def test_step7_checkpoint_reconciles_with_the_delta_ledger() -> None:
    step6 = json.loads(
        (Path(__file__).parent / "oracles" / "real_workload_step6.json").read_text(
            encoding="utf-8"
        )
    )
    step7 = json.loads(
        (Path(__file__).parent / "oracles" / "real_workload_step7.json").read_text(
            encoding="utf-8"
        )
    )
    ledger = json.loads(
        (Path(__file__).parent / "oracles" / "step_delta_ledger.json").read_text(
            encoding="utf-8"
        )
    )["steps"]["7"]

    assert step7["atomic_findings"] - step6["atomic_findings"] == (
        ledger["expected_atomic_delta"]
    )
    assert step7["budget_omitted"] == 0
    assert step7["severity"] == step6["severity"]
    assert step7["review_counts"] == step6["review_counts"]
    assert step7["coverage_states"] == step6["coverage_states"]
    assert step7["dependency_index"] == step6["dependency_index"]
    assert step7["median_peak_rss_mib"] <= step6["median_peak_rss_mib"] * 1.05
    assert step7["median_elapsed_seconds"] <= step6["median_elapsed_seconds"] * 1.10
    assert step7["source_hashes_unchanged"] is True


def test_impacts_and_excerpts_are_computed_only_for_retained_findings(
    tmp_path: Path, tiny_budget: None, monkeypatch: pytest.MonkeyPatch
) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _capped_workbook(baseline, rows=60, offset=0)
    _capped_workbook(current, rows=60, offset=1)

    closure_sources: list[tuple[str, int, int]] = []
    original_closure = dependency_module.dependent_nodes_of
    monkeypatch.setattr(
        dependency_module,
        "dependent_nodes_of",
        lambda graph, node: (closure_sources.append(node), original_closure(graph, node))[1],
    )
    excerpt_calls: list[int] = []
    original_attach = context_module.attach_excerpts
    monkeypatch.setattr(
        engine_module,
        "attach_excerpts",
        lambda findings, base, curr: (
            excerpt_calls.append(len(findings)),
            original_attach(findings, base, curr),
        )[1],
    )

    result = run_qc(baseline_excel=baseline, current_excel=current)

    value_findings = [
        finding
        for finding in result.findings
        if finding.finding_class is FindingClass.VALUE_CHANGED
    ]
    capped = [
        finding
        for finding in result.findings
        if finding.finding_class is FindingClass.FINDINGS_CAPPED
    ]

    assert len(value_findings) == 5
    assert capped, "the budget must disclose the omission"
    assert excerpt_calls == [len(result.findings)]
    assert sorted(closure_sources) == [("Data", row, 2) for row in range(2, 7)]


def test_dependency_capability_stays_checked_while_output_degrades() -> None:
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
            check_id="excel-dependencies",
            label="Formula dependency impact tracing",
            artifact="excel",
            state=CoverageState.CHECKED,
            detail="All parsed formula references were resolved",
        ),
        CoverageItem(
            check_id="excel-values",
            label="Excel values",
            artifact="excel",
            state=CoverageState.CHECKED,
        ),
    ]
    original = limit_findings

    def _tiny(items: list[Finding]):
        return original(items, max_per_class_scope=2, max_total=10)

    engine_module.limit_findings = _tiny  # type: ignore[assignment]
    try:
        _apply_findings_budget(findings, coverage)
    finally:
        engine_module.limit_findings = original  # type: ignore[assignment]

    dependencies, values = coverage
    assert dependencies.state is CoverageState.CHECKED
    assert "extraction and indexing were complete" in dependencies.detail
    assert "impacts and cell context were not computed" in dependencies.detail
    assert "Output budget" not in dependencies.detail
    assert values.state is CoverageState.DEGRADED
    assert "Output budget omitted 3 findings" in values.detail


def test_retained_atomic_contract_and_fail_on_behavior_survive_the_reordering(
    tmp_path: Path,
) -> None:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _capped_workbook(baseline, rows=8, offset=0)
    _capped_workbook(current, rows=8, offset=1)

    result = run_qc(baseline_excel=baseline, current_excel=current)

    changed = [
        finding
        for finding in result.findings
        if finding.finding_class is FindingClass.VALUE_CHANGED
    ]
    assert len(changed) == 8
    assert all(finding.current_excerpt is not None for finding in changed)
    assert all(finding.baseline_excerpt is not None for finding in changed)
    assert any(finding.impacts for finding in changed)
    assert result.counts
    assert result.mode is QCRunMode.CYCLE_COMPARISON
