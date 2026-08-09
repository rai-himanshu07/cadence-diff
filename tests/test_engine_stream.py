"""Step 4 contracts: the engine's spill-backed cycle pipeline.

The full suite plus the compatibility harness prove byte-equality with the
old in-memory triage; these tests pin the streaming-specific behaviors that
equality alone would not catch if they silently degraded.
"""

from __future__ import annotations

import gc
from pathlib import Path

from qc_tool.engine import QCRunResult, run_qc
from qc_tool.findings import Severity
from qc_tool.findings_store import BlockFile, FindingSequence


def test_cycle_result_is_a_lazy_sequence_with_ordered_ids(
    qc_result: QCRunResult,
) -> None:
    assert isinstance(qc_result.findings, FindingSequence)
    ids = [finding.finding_id for finding in qc_result.findings]
    assert ids == [f"F{index:04d}" for index in range(1, len(ids) + 1)]


def test_severity_counts_cache_matches_recomputation(
    qc_result: QCRunResult,
) -> None:
    assert qc_result.severity_counts is not None
    recomputed = dict.fromkeys(Severity, 0)
    for finding in qc_result.findings:
        assert finding.severity is not None
        recomputed[finding.severity] += 1
    assert qc_result.counts == recomputed


def test_private_sidecar_fields_survive_the_spill(qc_result: QCRunResult) -> None:
    """Excluded-from-serialization fields must ride through the round-trip.

    History stores counterfactual bases and series anchors from the in-memory
    result; losing them in the spill would silently empty both sidecars.
    """
    bases = [
        finding
        for finding in qc_result.findings
        if finding.counterfactual_basis is not None
    ]
    anchors = [
        finding
        for finding in qc_result.findings
        if finding.series_anchor is not None
    ]
    assert bases, "fixture pair must produce numeric counterfactual bases"
    assert anchors, "fixture pair must produce series anchors"
    for finding in bases:
        basis = finding.counterfactual_basis
        assert basis is not None
        assert basis.sheet == finding.sheet
    for finding in anchors:
        anchor = finding.series_anchor
        assert anchor is not None
        assert anchor.sheet == finding.sheet


def test_spill_directory_is_removed_when_the_result_is_collected(
    fixture_dir: Path,
) -> None:
    result = run_qc(
        baseline_excel=fixture_dir / "baseline.xlsx",
        current_excel=fixture_dir / "current.xlsx",
    )
    findings = result.findings
    assert isinstance(findings, FindingSequence)
    source = findings._source
    assert isinstance(source, BlockFile)
    store_dir = source.path.parent
    assert store_dir.exists()
    assert not (store_dir / "spill.qcfb").exists()  # merged spill is deleted
    del result
    del findings
    gc.collect()
    assert not store_dir.exists()
