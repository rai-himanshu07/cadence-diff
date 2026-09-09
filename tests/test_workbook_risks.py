"""Typed intrinsic workbook risks shared across every QC run mode."""

from __future__ import annotations

import io
import json
import zipfile
from pathlib import Path

import pytest

from qc_tool import engine as engine_module
from qc_tool.coverage import CoverageState, QCRunMode
from qc_tool.engine import run_qc
from qc_tool.excel.workbook_risks import (
    external_link_reachability_coverage,
    workbook_risk_findings,
)
from qc_tool.findings import FindingClass, FindingProvenance, Severity
from qc_tool.io.loader import _extract_ooxml_risks, load_workbook_snapshot
from qc_tool.io.model import (
    CellRecord,
    ExternalLinkReachability,
    SheetSnapshot,
    WorkbookRisk,
    WorkbookRiskKind,
    WorkbookSnapshot,
)
from qc_tool.ppt.model import DeckSnapshot
from qc_tool.triage.rules import triage


def _snapshot(
    source_name: str,
    risks: list[WorkbookRisk] | None = None,
    *,
    external_link_reachability: ExternalLinkReachability | None = None,
) -> WorkbookSnapshot:
    return WorkbookSnapshot(
        source_name=source_name,
        file_format="xlsb",
        formulas_available=False,
        styles_available=False,
        formula_presence_available=False,
        tables_available=False,
        charts_available=False,
        interaction_rules_available=False,
        intrinsic_risks=risks or [],
        external_link_reachability=external_link_reachability,
        sheets=[
            SheetSnapshot(
                "Data",
                "visible",
                1,
                1,
                {(1, 1): CellRecord(1, 1, 1.0)},
            )
        ],
    )


def test_workbook_risk_kind_inventory_is_closed() -> None:
    assert {kind.value for kind in WorkbookRiskKind} == {
        "external_workbook_link",
        "external_relationship",
        "external_data_connection",
        "query_table",
        "vba_project",
        "excel4_macro_sheet",
        "activex_control",
        "embedded_ole",
        "control_content",
        "dialog_sheet",
        "custom_office_ui",
        "unreadable_relationship_metadata",
    }


def test_risk_classification_and_cycle_provenance_matrix() -> None:
    kinds = list(WorkbookRiskKind)
    baseline = _snapshot(
        "baseline.xlsb",
        [WorkbookRisk(kind) for kind in kinds[::2]],
    )
    current = _snapshot(
        "current.xlsb",
        [WorkbookRisk(kind, count=2) for kind in kinds],
    )

    findings = triage(workbook_risk_findings(current, baseline))

    assert len(findings) == len(kinds)
    assert {finding.severity for finding in findings} == {Severity.CRITICAL}
    assert {
        finding.finding_class
        for finding in findings
        if finding.element
        in {
            "external_workbook_link",
            "external_relationship",
            "external_data_connection",
            "query_table",
        }
    } == {FindingClass.EXTERNAL_LINK}
    assert {
        finding.finding_class
        for finding in findings
        if finding.element
        not in {
            "external_workbook_link",
            "external_relationship",
            "external_data_connection",
            "query_table",
        }
    } == {FindingClass.ACTIVE_CONTENT}
    by_kind = {finding.element: finding for finding in findings}
    for kind in kinds:
        expected = (
            FindingProvenance.INHERITED
            if kind in kinds[::2]
            else FindingProvenance.NEW
        )
        assert by_kind[kind.value].provenance is expected
        assert by_kind[kind.value].current_value == "2"


def test_ooxml_risk_scan_never_retains_external_targets() -> None:
    secret_target = "file:///sensitive/client/source.xlsx"
    relationships = f"""<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
        Target="{secret_target}" TargetMode="External"/>
    </Relationships>""".encode()
    workbook_relationships = b"""<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdExternal"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
        Target="externalLinks/externalLink1.xml"/>
    </Relationships>"""
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("xl/externalLinks/externalLink1.xml", b"<externalLink/>")
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_relationships)
        archive.writestr(
            "xl/externalLinks/_rels/externalLink1.xml.rels",
            relationships,
        )
        archive.writestr("xl/vbaProject.bin", b"synthetic")

    risks = _extract_ooxml_risks(stream.getvalue())
    snapshot = _snapshot("current.xlsx", risks)
    findings = workbook_risk_findings(snapshot)
    serialized = json.dumps(
        {
            "risks": [risk.__dict__ if hasattr(risk, "__dict__") else repr(risk) for risk in risks],
            "findings": [finding.model_dump(mode="json") for finding in findings],
        }
    )

    assert {risk.kind: risk.count for risk in risks} == {
        WorkbookRiskKind.EXTERNAL_RELATIONSHIP: 1,
        WorkbookRiskKind.EXTERNAL_WORKBOOK_LINK: 1,
        WorkbookRiskKind.VBA_PROJECT: 1,
    }
    assert secret_target not in serialized
    assert all(secret_target not in finding.message for finding in findings)


@pytest.mark.parametrize(
    "mode",
    [
        QCRunMode.CURRENT_FILE_PREFLIGHT,
        QCRunMode.CYCLE_COMPARISON,
        QCRunMode.FINAL_PACKAGE,
    ],
)
def test_intrinsic_risks_surface_in_every_run_mode(
    mode: QCRunMode,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    baseline = _snapshot("baseline.xlsb")
    current = _snapshot(
        "current.xlsb",
        [
            WorkbookRisk(WorkbookRiskKind.EXTERNAL_WORKBOOK_LINK),
            WorkbookRisk(WorkbookRiskKind.VBA_PROJECT),
        ],
    )

    def fake_excel(path: Path, **_: object) -> WorkbookSnapshot:
        return baseline if path.name.startswith("baseline") else current

    monkeypatch.setattr(engine_module, "_load_excel_file", fake_excel)
    monkeypatch.setattr(
        engine_module,
        "_load_powerpoint_file",
        lambda *_args, **_kwargs: DeckSnapshot("current.pptx"),
    )
    kwargs: dict[str, object] = {
        "current_excel": Path("current.xlsb"),
        "mode": mode,
    }
    if mode is QCRunMode.CYCLE_COMPARISON:
        kwargs["baseline_excel"] = Path("baseline.xlsb")
    elif mode is QCRunMode.FINAL_PACKAGE:
        kwargs["current_ppt"] = Path("current.pptx")

    result = run_qc(**kwargs)  # type: ignore[arg-type]
    risks = [
        finding
        for finding in result.findings
        if finding.finding_class
        in {FindingClass.EXTERNAL_LINK, FindingClass.ACTIVE_CONTENT}
    ]

    assert {finding.finding_class for finding in risks} == {
        FindingClass.EXTERNAL_LINK,
        FindingClass.ACTIVE_CONTENT,
    }
    assert {finding.severity for finding in risks} == {Severity.CRITICAL}
    assert all("source.xlsx" not in finding.message for finding in risks)


def test_loaded_snapshots_expose_only_typed_intrinsic_risks(
    fixture_dir: Path,
) -> None:
    snapshot = load_workbook_snapshot(fixture_dir / "current.xlsx")

    assert not hasattr(snapshot, "external_links")
    assert all(isinstance(risk, WorkbookRisk) for risk in snapshot.intrinsic_risks)


# --- Step 4b: passive-link suppression requires proven reachability ------


def _passive_snapshot(
    *, external_link_reachability: ExternalLinkReachability | None
) -> WorkbookSnapshot:
    return _snapshot(
        "current.xlsb",
        [
            WorkbookRisk(WorkbookRiskKind.EXTERNAL_WORKBOOK_LINK),
            WorkbookRisk(WorkbookRiskKind.EXTERNAL_RELATIONSHIP),
            WorkbookRisk(WorkbookRiskKind.VBA_PROJECT),
        ],
        external_link_reachability=external_link_reachability,
    )


def test_proven_inactive_passive_link_is_suppressed_from_findings() -> None:
    current = _passive_snapshot(
        external_link_reachability=ExternalLinkReachability(proven=True, live=False)
    )

    findings = workbook_risk_findings(current)

    assert {finding.element for finding in findings} == {"vba_project"}


def test_proven_live_passive_link_is_not_suppressed() -> None:
    current = _passive_snapshot(
        external_link_reachability=ExternalLinkReachability(
            proven=True, live=True, direct_reference_count=1
        )
    )

    findings = workbook_risk_findings(current)

    assert {finding.element for finding in findings} == {
        "external_workbook_link",
        "external_relationship",
        "vba_project",
    }


@pytest.mark.parametrize(
    "reachability",
    [None, ExternalLinkReachability(proven=False)],
)
def test_unproven_reachability_never_suppresses(
    reachability: ExternalLinkReachability | None,
) -> None:
    current = _passive_snapshot(external_link_reachability=reachability)

    findings = workbook_risk_findings(current)

    assert {finding.element for finding in findings} == {
        "external_workbook_link",
        "external_relationship",
        "vba_project",
    }


def test_reachability_coverage_checked_when_no_passive_risk() -> None:
    current = _snapshot("current.xlsb", [WorkbookRisk(WorkbookRiskKind.VBA_PROJECT)])

    item = external_link_reachability_coverage(current)

    assert item.state is CoverageState.CHECKED
    assert item.detail == ""


def test_reachability_coverage_checked_when_proven_either_way() -> None:
    inactive = _passive_snapshot(
        external_link_reachability=ExternalLinkReachability(proven=True, live=False)
    )
    live = _passive_snapshot(
        external_link_reachability=ExternalLinkReachability(proven=True, live=True)
    )

    assert external_link_reachability_coverage(inactive).state is CoverageState.CHECKED
    assert external_link_reachability_coverage(live).state is CoverageState.CHECKED
    assert "not used" in external_link_reachability_coverage(inactive).detail
    assert "in use" in external_link_reachability_coverage(live).detail


def test_reachability_coverage_degraded_when_unproven() -> None:
    current = _passive_snapshot(external_link_reachability=None)

    item = external_link_reachability_coverage(current)

    assert item.state is CoverageState.DEGRADED
    assert "cannot be proven" in item.detail


def test_reachability_coverage_degrades_across_multiple_workbooks() -> None:
    proven = _passive_snapshot(
        external_link_reachability=ExternalLinkReachability(proven=True, live=False)
    )
    unproven = _passive_snapshot(external_link_reachability=None)

    item = external_link_reachability_coverage(proven, unproven)

    assert item.state is CoverageState.DEGRADED
