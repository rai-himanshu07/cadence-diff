"""Step 4: wrapper/subtree detection and skeleton pattern grouping."""

from __future__ import annotations

import logging

import pytest

import qc_tool.excel.formulas as formulas_module
from qc_tool.excel.align import AxisAlignment, RegionAlignment, WorkbookAlignment
from qc_tool.excel.formulas import (
    detect_formula_wrapper,
    diff_workbook_formulas,
)
from qc_tool.excel.regions import TableRegion
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingEvidenceTag,
    FindingSubtype,
    Severity,
)
from qc_tool.io.model import CellRecord, SheetSnapshot, WorkbookSnapshot
from qc_tool.review import build_pattern_groups
from qc_tool.triage.rules import assign_severity

GUARD = '=IF(XLOOKUP(R1C8,C[3],C[4],FALSE),{core},NA())'


class TestDetectFormulaWrapper:
    def test_exact_wrapped(self) -> None:
        base = "=IF(R[1]C[-2],RC[3],NA())"
        curr = GUARD.format(core="IF(R[1]C[-2],RC[3],NA())")
        match = detect_formula_wrapper(base, curr)
        assert match is not None
        assert match.kind == "wrapped"
        assert match.exact is True

    def test_exact_unwrapped(self) -> None:
        base = GUARD.format(core="IF(R[1]C[-2],RC[3],NA())")
        curr = "=IF(R[1]C[-2],RC[3],NA())"
        match = detect_formula_wrapper(base, curr)
        assert match is not None
        assert match.kind == "unwrapped"
        assert match.exact is True

    def test_shifted_inner_references_match_by_shape(self) -> None:
        # The rolling-window case: the inner logic is identical in shape but
        # its relative references moved one row.
        base = "=IF(R[1]C[-2],RC[3],NA())"
        curr = GUARD.format(core="IF(R[2]C[-2],R[1]C[3],NA())")
        match = detect_formula_wrapper(base, curr)
        assert match is not None
        assert match.kind == "wrapped"
        assert match.exact is False

    def test_same_skeleton_for_shifted_and_exact_cores(self) -> None:
        base = "=IF(R[1]C[-2],RC[3],NA())"
        exact = detect_formula_wrapper(base, GUARD.format(core=base[1:]))
        shifted = detect_formula_wrapper(
            base, GUARD.format(core="IF(R[2]C[-2],R[1]C[3],NA())")
        )
        assert exact is not None and shifted is not None
        assert exact.skeleton_key == shifted.skeleton_key

    def test_operator_extension_is_not_a_wrapper(self) -> None:
        # `A1+B1` inside `A1+B1*2` is not an argument-level subtree.
        assert detect_formula_wrapper("=RC[1]+RC[2]", "=RC[1]+RC[2]*2") is None

    def test_token_prefix_is_not_a_match(self) -> None:
        # D1 vs D1000: string containment but not token containment.
        assert detect_formula_wrapper("=RC[-1]-R1C4", "=RC[-1]-R1000C4") is None

    def test_refactor_is_not_a_wrapper(self) -> None:
        assert (
            detect_formula_wrapper("=SUM(R1C1:R5C1)", "=R1C1+R2C1+R3C1+R4C1+R5C1")
            is None
        )

    def test_string_literal_cannot_fake_containment(self) -> None:
        base = "=R[1]C"
        curr = '=IF(RC[1],"R[1]C,x",R[9]C[9])'
        assert detect_formula_wrapper(base, curr) is None

    def test_bare_literal_core_is_not_meaningful(self) -> None:
        assert detect_formula_wrapper("=5", "=IF(RC[1],5,NA())") is None

    def test_malformed_formula_fails_closed(self) -> None:
        assert detect_formula_wrapper('="unterminated', "=IF(A1,1,2)") is None


def _formula_workbook(
    formulas: dict[tuple[int, int], str], *, rows: int, cols: int
) -> WorkbookSnapshot:
    cells = {
        (row, col): CellRecord(row, col, 1.0, formula=formula, is_formula=True)
        for (row, col), formula in formulas.items()
    }
    return WorkbookSnapshot(
        "synthetic.xlsx",
        "xlsx",
        True,
        True,
        formula_presence_available=True,
        formula_source="openpyxl",
        sheets=[SheetSnapshot("Synthetic", "visible", rows, cols, cells)],
    )


def _block_alignment(rows: int, cols: int) -> WorkbookAlignment:
    region = TableRegion("Synthetic", 1, 1, rows, cols, "block", None, 1, "none")
    return WorkbookAlignment(
        common_sheets=["Synthetic"],
        regions={
            "Synthetic": [
                RegionAlignment(
                    region,
                    region,
                    AxisAlignment(pairs=[(row, row) for row in range(1, rows + 1)]),
                    AxisAlignment(pairs=[(col, col) for col in range(1, cols + 1)]),
                )
            ]
        },
    )


class TestDiffIntegration:
    def test_reference_tag_parse_failure_keeps_finding_and_logs_fixed_code(
        self,
        monkeypatch: pytest.MonkeyPatch,
        caplog: pytest.LogCaptureFixture,
    ) -> None:
        baseline = _formula_workbook({(1, 1): "=B1+C1"}, rows=1, cols=3)
        current = _formula_workbook({(1, 1): "=B1+D1"}, rows=1, cols=4)

        def fail_reference_parse(_formula: str) -> list[object]:
            raise ValueError("synthetic malformed formula")

        monkeypatch.setattr(
            formulas_module,
            "formula_reference_operands",
            fail_reference_parse,
        )
        with caplog.at_level(logging.WARNING, logger=formulas_module.__name__):
            findings = diff_workbook_formulas(
                baseline,
                current,
                _block_alignment(1, 4),
                _use_native_delta=False,
            )

        logic = [
            finding
            for finding in findings
            if finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
        ]
        assert len(logic) == 1
        assert FindingEvidenceTag.ADDED_REFERENCE not in logic[0].evidence_tags
        assert "formula-reference-tag-unavailable" in caplog.text

    def test_rollout_collapses_to_one_skeleton_group(self) -> None:
        wrap = "=IF(XLOOKUP($H$1,$F:$F,$G:$G,FALSE),{core},NA())"
        baseline_formulas = {
            (1, 1): "=IF(B1,C1,NA())",
            (2, 1): "=IF(B2,C2,NA())",
            (3, 1): "=IF(B3,C3,NA())",
            (4, 1): "=D4*2",
        }
        current_formulas = {
            (1, 1): wrap.format(core="IF(B1,C1,NA())"),
            (2, 1): wrap.format(core="IF(B2,C2,NA())"),
            (3, 1): wrap.format(core="IF(B3,C3,NA())"),
            (4, 1): "=D4*3",  # unrelated logic change
        }
        baseline = _formula_workbook(baseline_formulas, rows=4, cols=8)
        current = _formula_workbook(current_formulas, rows=4, cols=8)
        findings = [
            f
            for f in diff_workbook_formulas(
                baseline, current, _block_alignment(4, 8)
            )
            if f.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
        ]
        assert len(findings) == 4

        wrapped = [f for f in findings if f.subtype is FindingSubtype.FORMULA_WRAPPED]
        plain = [f for f in findings if f.subtype is None]
        assert len(wrapped) == 3
        assert len(plain) == 1
        assert len({f.event_key for f in wrapped}) == 1
        assert wrapped[0].event_key.startswith("formula-wrapper:wrapped:")
        assert "preserved inside a new wrapper" in wrapped[0].message
        assert all(FindingEvidenceTag.EXACT_WRAPPER in f.evidence_tags for f in wrapped)
        assert all(FindingEvidenceTag.ADDED_REFERENCE in f.evidence_tags for f in wrapped)
        assert all(assign_severity(f) is Severity.WARNING for f in findings)

        for finding in findings:
            finding.severity = assign_severity(finding)
        groups = [
            g
            for g in build_pattern_groups(findings)
            if g.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
        ]
        assert len(groups) == 2  # one skeleton group + the unrelated change
        by_size = sorted(groups, key=lambda g: -g.member_count)
        assert by_size[0].member_count == 3
        assert by_size[1].member_count == 1

    def test_mixed_wrappers_split_by_skeleton(self) -> None:
        baseline_formulas = {
            (1, 1): "=IF(B1,C1,NA())",
            (2, 1): "=IF(B2,C2,NA())",
        }
        current_formulas = {
            (1, 1): "=IF(XLOOKUP($H$1,$F:$F,$G:$G,FALSE),IF(B1,C1,NA()),NA())",
            (2, 1): "=IFERROR(IF(B2,C2,NA()),0)",  # different guard family
        }
        baseline = _formula_workbook(baseline_formulas, rows=2, cols=8)
        current = _formula_workbook(current_formulas, rows=2, cols=8)
        findings = [
            f
            for f in diff_workbook_formulas(
                baseline, current, _block_alignment(2, 8)
            )
            if f.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
        ]
        assert len(findings) == 2
        assert all(f.subtype is FindingSubtype.FORMULA_WRAPPED for f in findings)
        assert len({f.event_key for f in findings}) == 2

    def test_shifted_wrapper_carries_shape_evidence(self) -> None:
        baseline = _formula_workbook(
            {(1, 1): "=IF(B1,C1,NA())"}, rows=1, cols=8
        )
        current = _formula_workbook(
            {
                (1, 1): (
                    "=IF(XLOOKUP($H$1,$F:$F,$G:$G,FALSE),"
                    "IF(B2,C2,NA()),NA())"
                )
            },
            rows=1,
            cols=8,
        )

        finding = next(
            f
            for f in diff_workbook_formulas(
                baseline, current, _block_alignment(1, 8)
            )
            if f.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
        )

        assert FindingEvidenceTag.SHAPE_WRAPPER in finding.evidence_tags
        assert FindingEvidenceTag.EXACT_WRAPPER not in finding.evidence_tags
        assert FindingEvidenceTag.ADDED_REFERENCE in finding.evidence_tags

    def test_expected_extension_skips_wrapper_detection(self) -> None:
        baseline_formulas = {(1, 1): "=SUM(B1:B5)"}
        current_formulas = {(1, 1): "=SUM(B1:B6)"}
        baseline = _formula_workbook(baseline_formulas, rows=1, cols=3)
        current = _formula_workbook(current_formulas, rows=1, cols=3)
        findings = [
            f
            for f in diff_workbook_formulas(
                baseline, current, _block_alignment(1, 3)
            )
            if f.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
        ]
        assert len(findings) == 1
        assert findings[0].expected_growth
        assert findings[0].subtype is None
        assert findings[0].event_key == ""


def test_wrapper_subtype_survives_serialization() -> None:
    finding = Finding(
        artifact="excel",
        finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
        subtype=FindingSubtype.FORMULA_WRAPPED,
        event_key="formula-wrapper:wrapped:abc123def456",
        message="x",
    )
    dumped = finding.model_dump(mode="json")
    restored = Finding.model_validate(dumped)
    assert restored.subtype is FindingSubtype.FORMULA_WRAPPED
    assert restored.event_key == finding.event_key
