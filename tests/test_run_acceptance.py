"""Step 4: run-level analyst acceptance threshold (visible-INFO, never hidden)."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from qc_tool.config.profile import NumericTolerance, default_profile
from qc_tool.coverage import QCRunMode
from qc_tool.engine import QCRunResult, run_qc
from qc_tool.findings import Finding, FindingClass, Materiality, Severity


def _write(path: Path, c3: float, d3: float) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Data"
    sheet.append(["Month", "A", "B"])
    sheet.append(["Jan-25", 50.0, 60.0])
    sheet.append(["Feb-25", c3, d3])  # old history: outside the 2-month window
    sheet.append(["Mar-25", 70.0, 80.0])
    sheet.append(["Apr-25", 71.0, 81.0])
    sheet.append(["May-25", 72.0, 82.0])
    sheet.append(["Jun-25", 73.0, 83.0])
    workbook.save(path)


@pytest.fixture()
def pair(tmp_path: Path) -> tuple[Path, Path]:
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    _write(baseline, 100.0, 1000.0)
    _write(current, 100.4, 1080.0)  # +0.4% and +8%
    return baseline, current


def _value_findings(result: QCRunResult) -> dict[str | None, Finding]:
    return {
        f.location: f
        for f in result.findings
        if f.finding_class is FindingClass.VALUE_CHANGED
    }


def test_threshold_off_keeps_material_criticals(pair: tuple[Path, Path]) -> None:
    baseline, current = pair
    result = run_qc(
        baseline_excel=baseline, current_excel=current, profile=default_profile()
    )
    values = _value_findings(result)
    assert values["B3"].materiality is Materiality.MATERIAL
    assert values["B3"].severity is Severity.CRITICAL
    assert not any("acceptance threshold" in d for d in result.disclosures)


def test_threshold_tiers_in_band_changes_to_visible_info(
    pair: tuple[Path, Path],
) -> None:
    baseline, current = pair
    result = run_qc(
        baseline_excel=baseline,
        current_excel=current,
        profile=default_profile(),
        run_acceptance=NumericTolerance(absolute=0.0, relative=0.01),
    )
    values = _value_findings(result)
    # +0.4% is inside the band; +8% is not.
    assert values["B3"].materiality is Materiality.WITHIN_TOLERANCE
    assert values["B3"].severity is Severity.INFO
    assert values["C3"].materiality is Materiality.MATERIAL
    assert values["C3"].severity is Severity.CRITICAL
    assert any(
        "acceptance threshold" in d and "1%" in d for d in result.disclosures
    )


def test_either_bound_accepts(pair: tuple[Path, Path]) -> None:
    baseline, current = pair
    result = run_qc(
        baseline_excel=baseline,
        current_excel=current,
        profile=default_profile(),
        run_acceptance=NumericTolerance(absolute=100.0, relative=0.0),
    )
    values = _value_findings(result)
    # |delta| 0.4 and 80 both within the absolute bound.
    assert values["B3"].materiality is Materiality.WITHIN_TOLERANCE
    assert values["C3"].materiality is Materiality.WITHIN_TOLERANCE


def test_zero_threshold_object_is_treated_as_off(pair: tuple[Path, Path]) -> None:
    baseline, current = pair
    result = run_qc(
        baseline_excel=baseline,
        current_excel=current,
        profile=default_profile(),
        run_acceptance=NumericTolerance(),
    )
    values = _value_findings(result)
    assert values["B3"].materiality is Materiality.MATERIAL
    assert not any("acceptance threshold" in d for d in result.disclosures)


@pytest.mark.parametrize(
    "mode",
    [QCRunMode.CURRENT_FILE_PREFLIGHT, QCRunMode.FINAL_PACKAGE],
)
def test_threshold_is_visibly_not_applied_without_a_pairwise_value_comparison(
    pair: tuple[Path, Path],
    fixture_dir: Path,
    mode: QCRunMode,
) -> None:
    _, current = pair
    kwargs: dict[str, object] = {
        "current_excel": current,
        "mode": mode,
        "run_acceptance": NumericTolerance(absolute=1.0),
    }
    if mode is QCRunMode.FINAL_PACKAGE:
        kwargs["current_ppt"] = fixture_dir / "current.pptx"

    result = run_qc(**kwargs)  # type: ignore[arg-type]

    assert any(
        "acceptance threshold was not applied" in disclosure
        for disclosure in result.disclosures
    )
    assert not any(
        "acceptance threshold active" in disclosure.casefold()
        for disclosure in result.disclosures
    )


def test_cli_rejects_negative_thresholds() -> None:
    from qc_tool.cli import _cmd_run

    with pytest.raises(ValueError, match="cannot be negative"):
        _cmd_run(["--current-excel", "x.xlsx", "--accept-percent", "-1"])


def test_service_rejects_negative_thresholds(tmp_path: Path) -> None:
    from qc_tool.run_service import perform_run

    with pytest.raises(ValueError, match="cannot be negative"):
        perform_run(
            tmp_path,
            {},
            {},
            default_profile(),
            acceptance_absolute=-0.1,
        )
