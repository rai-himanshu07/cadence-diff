"""Contracts for the block-compressed findings store (plan step 2)."""

from __future__ import annotations

import json
import tracemalloc
import zlib
from pathlib import Path

import pytest

from qc_tool.findings import Finding, FindingClass, Severity
from qc_tool.findings_store import (
    BlockFile,
    FindingSequence,
    FindingsStoreError,
    SpillWriter,
    encode_block,
    merge_spill,
    write_finding_blocks,
)


def _finding(index: int) -> Finding:
    return Finding(
        finding_id=f"F{index + 1:04d}",
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        severity=Severity.CRITICAL,
        sheet="Data",
        location=f"B{index + 2}",
        baseline_value=str(index),
        current_value=str(index + 1),
        message=f"Data!B{index + 2}: value changed",
    )


def test_sequence_matches_in_memory_list_exactly(tmp_path: Path) -> None:
    findings = [_finding(index) for index in range(2_517)]
    container = write_finding_blocks(
        tmp_path / "findings.qcfb",
        (finding.model_dump(mode="json") for finding in findings),
        block_rows=500,
    )
    sequence = FindingSequence(container, cache_blocks=2)

    assert len(sequence) == len(findings)
    assert list(sequence) == findings
    assert sequence[0] == findings[0]
    assert sequence[1_234] == findings[1_234]
    assert sequence[-1] == findings[-1]
    assert sequence[500:503] == findings[500:503]
    with pytest.raises(IndexError):
        sequence[len(findings)]


def test_corrupted_block_fails_closed(tmp_path: Path) -> None:
    path = tmp_path / "findings.qcfb"
    container = write_finding_blocks(
        path,
        (_finding(index).model_dump(mode="json") for index in range(10)),
        block_rows=5,
    )
    info = container.block_infos()[0]
    raw = bytearray(path.read_bytes())
    raw[info.offset + 2] ^= 0xFF
    path.write_bytes(bytes(raw))

    sequence = FindingSequence(BlockFile.open(path))
    with pytest.raises(FindingsStoreError):
        sequence[0]


def test_invalid_payload_fails_closed(tmp_path: Path) -> None:
    path = tmp_path / "findings.qcfb"
    with BlockFile(path) as container:
        container.append_rows([{"finding_class": "not_a_real_class"}])
    sequence = FindingSequence(BlockFile.open(path))
    with pytest.raises(FindingsStoreError):
        sequence[0]


def test_footer_corruption_is_a_typed_error(tmp_path: Path) -> None:
    path = tmp_path / "findings.qcfb"
    write_finding_blocks(
        path, (_finding(index).model_dump(mode="json") for index in range(3))
    )
    raw = bytearray(path.read_bytes())
    raw[-9] ^= 0xFF  # inside the JSON footer
    path.write_bytes(bytes(raw))
    with pytest.raises(FindingsStoreError):
        BlockFile.open(path)


def test_spill_merge_yields_global_sort_order(tmp_path: Path) -> None:
    rows = [((index * 7919) % 100_003, f"payload-{index}") for index in range(20_000)]
    with SpillWriter(tmp_path / "spill.qcfb", block_rows=750) as spill:
        for key, payload in rows:
            spill.append([key], payload)
    assert spill.total_rows == len(rows)

    merged = list(merge_spill(spill.path))
    expected = [payload for _, payload in sorted(rows, key=lambda row: row[0])]
    assert merged == expected


def test_spill_merge_reduces_fan_in_with_scratch_files(tmp_path: Path) -> None:
    rows = [((index * 31) % 977, index) for index in range(6_000)]
    with SpillWriter(tmp_path / "spill.qcfb", block_rows=100) as spill:
        for key, payload in rows:
            spill.append([key, payload], payload)

    merged = list(merge_spill(spill.path, fan_in=4))
    assert merged == [payload for _, payload in sorted(rows, key=lambda r: (r[0], r[1]))]
    assert not list(tmp_path.glob("*.merge*"))


def test_spill_merge_converges_past_fan_in_times_block_findings(
    tmp_path: Path,
) -> None:
    """Runs, not blocks, must drive the merge recursion.

    With more rows than ``fan_in * BLOCK_FINDINGS`` a repacked scratch file
    holds as many blocks as its input, so block counting loops forever; run
    counting divides by ``fan_in`` every pass and terminates.
    """
    rows = [((index * 37) % 1_009, index) for index in range(12_000)]
    with SpillWriter(tmp_path / "spill.qcfb", block_rows=100) as spill:
        for key, payload in rows:
            spill.append([key, payload], payload)

    merged = list(merge_spill(spill.path, fan_in=2))
    assert merged == [payload for _, payload in sorted(rows, key=lambda r: (r[0], r[1]))]
    assert not list(tmp_path.glob("*.merge*"))


def test_iteration_memory_stays_flat_for_large_stores(tmp_path: Path) -> None:
    payload = _finding(0).model_dump(mode="json")
    count = 500_000

    def payloads():
        for index in range(count):
            payload["location"] = f"B{index}"
            yield payload

    container = write_finding_blocks(tmp_path / "big.qcfb", payloads())

    sequence = FindingSequence(container)
    tracemalloc.start()
    seen = 0
    for _ in sequence.iter_payloads():
        seen += 1
    _, peak = tracemalloc.get_traced_memory()
    tracemalloc.stop()

    assert seen == count
    assert peak < 128 * 1024 * 1024


def test_compression_ratio_is_material_on_repetitive_findings(tmp_path: Path) -> None:
    payloads = [_finding(index).model_dump(mode="json") for index in range(5_000)]
    raw = sum(len(json.dumps(p, separators=(",", ":"))) for p in payloads)
    blob = encode_block(payloads)
    assert raw / len(blob) > 8
    assert zlib.decompress(blob)  # round-trips


def test_trusted_constructor_equals_validation_on_every_field(
    qc_result,
) -> None:
    """Slice A safety: trusted decode must equal validated decode exactly."""
    import json

    from qc_tool.findings import (
        Finding,
        FindingClass,
        FindingEvidenceTag,
        FindingExpectedReason,
        FindingProvenance,
        FindingSubtype,
        FindingTemporalContext,
        GridExcerpt,
        LogicalFindingAddress,
        Materiality,
        NumericCounterfactualBasis,
        SeriesAnchorV2,
        Severity,
    )
    from qc_tool.findings_store import finding_payload

    # every real finding in the fixture corpus round-trips identically
    for finding in qc_result.findings:
        payload = json.loads(json.dumps(finding_payload(finding)))
        trusted = Finding.from_trusted_payload(payload)
        validated = Finding.model_validate(payload)
        assert trusted == validated
        assert trusted.model_dump(mode="json") == validated.model_dump(mode="json")

    # a synthetic finding populating EVERY typed field guards against drift
    full = Finding(
        finding_id="F0001",
        artifact="excel",
        artifact_member="member-b",
        finding_class=FindingClass.VALUE_CHANGED,
        severity=Severity.CRITICAL,
        expected_reason=FindingExpectedReason.PERIOD_PROGRESSION,
        provenance=FindingProvenance.NEW,
        subtype=FindingSubtype.VALUE_REPLACEMENT,
        materiality=Materiality.MATERIAL,
        temporal_context=FindingTemporalContext.HISTORICAL,
        evidence_tags={FindingEvidenceTag.DISPLAY_EQUIVALENT},
        event_key="event",
        sheet="Data",
        location="B5",
        baseline_location="B4",
        element="Total",
        baseline_value="1",
        current_value="2",
        message="value changed",
        impacts=["Summary!C1"],
        baseline_excerpt=GridExcerpt(
            cols=["A", "B"], rows=[1, 2], cells=[["a", "b"], ["c", "d"]],
            hit_row=1, hit_col=1,
        ),
        current_excerpt=GridExcerpt(
            cols=["A"], rows=[1], cells=[["x"]], hit_row=0, hit_col=0,
        ),
        analyst_comment="note",
        severity_overridden=True,
        root_cause_key="root",
        waiver_reason="waived",
        waiver_expires="2027-01-01",
        focus_shape_id=3,
        baseline_focus_shape_id=2,
        counterfactual_basis=NumericCounterfactualBasis(
            baseline=1.0, current=2.0, number_format="0.0",
            sheet="Data", location="B5",
        ),
        series_anchor=SeriesAnchorV2(
            sheet="Data", current_region_id="r1", period_axis="rows",
            series_index=2, period_index=3, segment="new_period",
        ),
        logical_address=LogicalFindingAddress(
            member_id="primary", sheet_id="data", region_id="r1", column_id="id",
            row_key_digest="a" * 64,
        ),
    )
    payload = json.loads(json.dumps(finding_payload(full)))
    trusted = Finding.from_trusted_payload(payload)
    validated = Finding.model_validate(payload)
    assert trusted == validated
    assert trusted.model_dump(mode="json") == validated.model_dump(mode="json")
    assert trusted.counterfactual_basis == validated.counterfactual_basis
    assert trusted.series_anchor == validated.series_anchor
    assert trusted.logical_address == validated.logical_address
    assert trusted.expected_growth is True  # derived flag must still derive

    # drift guard: a new model field forces this test to know about it
    known = {
        "finding_id", "artifact", "artifact_member", "finding_class",
        "severity", "expected_growth", "expected_reason", "provenance",
        "subtype", "materiality", "temporal_context", "evidence_tags",
        "event_key", "sheet", "location", "baseline_location", "element",
        "slide", "slide_index", "baseline_slide_index", "focus_shape_id",
        "baseline_focus_shape_id", "baseline_value", "current_value",
        "message", "impacts", "baseline_excerpt", "current_excerpt",
        "analyst_comment", "severity_overridden", "root_cause_key",
        "waiver_reason", "waiver_expires", "counterfactual_basis",
        "series_anchor", "population", "logical_address",
    }
    assert set(Finding.model_fields) == known


def test_iter_trusted_equals_validating_iteration(tmp_path) -> None:
    """The trusted sequence pass yields the same findings as __iter__."""
    from qc_tool.findings import Finding, FindingClass, Severity
    from qc_tool.findings_store import (
        FindingSequence,
        finding_payload,
        write_finding_blocks,
    )

    findings = [
        Finding(
            finding_id=f"F{index:04d}",
            artifact="excel",
            finding_class=FindingClass.VALUE_CHANGED,
            severity=Severity.WARNING,
            sheet="Data",
            location=f"B{index}",
            message=f"changed {index}",
        )
        for index in range(1, 12)
    ]
    path = tmp_path / "trusted.qcfb"
    container = write_finding_blocks(
        path, (finding_payload(f) for f in findings)
    )
    sequence = FindingSequence(container)
    assert list(sequence.iter_trusted()) == list(iter(sequence))
