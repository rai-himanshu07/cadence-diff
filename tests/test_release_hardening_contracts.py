"""Frozen failing contracts for the confirmed release-hardening blockers
(plan-20260909-release-hardening-and-ranked-table-review.md, Step 1).

Each test below reproduces one "Confirmed Blocker"/"Evidence" item from the
plan's own text against the CURRENT code, using only safe synthetic
fixtures and monkeypatched adapters -- never a real workbook or private
data. Every test is expected to FAIL until the step named in its docstring
lands; do not "fix" a test here without also shipping the corresponding
production change and updating the docstring/step reference.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import hmac
import json
import zipfile
from pathlib import Path

import pytest

import qc_tool.io.native_formula as native_formula_module
from qc_tool.attestation import create_attestation, verify_attestation
from qc_tool.config.profile import PopulationPolicy, default_profile
from qc_tool.engine import QCRunResult, compare_findings
from qc_tool.excel.population import CandidateSpill, finalize_populations
from qc_tool.findings import (
    Finding,
    FindingClass,
    FindingEvidenceTag,
    MembershipCodec,
    PopulationEvidence,
)
from qc_tool.io.formula_cache import (
    FormulaCacheKey,
    FormulaExtractionCache,
    coordinate_digest,
    package_digest,
)
from qc_tool.io.formula_enrichment import FormulaEnrichmentError, FormulaExtraction
from qc_tool.io.native_formula import (
    FormulaSheetSurface,
    WorkbookFormulaSurface,
    extract_formulas_with_native_kernel,
)
from qc_tool.io.xlsb_formula import XlsbFormulaScan
from qc_tool.scope import ComparisonScope
from tests.conftest import fixture_profile

_TODAY = dt.date(2026, 9, 9)


def _scan() -> XlsbFormulaScan:
    return XlsbFormulaScan(formula_cells={"Data": frozenset({(1, 1), (2, 2)})})


# --- Criterion 1: native A1 render failures must never become "=" ----------


def test_extract_formulas_with_native_kernel_must_not_turn_a_failed_a1_render_into_a_bare_equals(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Step 2. A Rust-side A1 render failure currently reaches Python as an
    empty string (`render(...).unwrap_or_default()` in
    `native/xlsbkernel/src/surface.rs`), and `extract_formulas_with_native_
    kernel` blindly prefixes every cell with '=' -- producing a literal '='
    formula that is accepted as real, complete coverage. It must instead be
    absent, exactly like an unresolved shared-formula follower already is
    in the same Rust function.
    """
    surface = WorkbookFormulaSurface(
        sheets=(
            FormulaSheetSurface(
                name="Data",
                cell_rows=(0, 1),
                cell_cols=(0, 1),
                cell_a1=("A2+1", ""),  # second cell failed to render
                cell_definition_id=(0, 1),
                definition_r1c1=("RC[1]+1", ""),
            ),
        ),
        defined_names=(),
    )
    monkeypatch.setattr(native_formula_module, "_xlsbkernel", object())
    monkeypatch.setattr(native_formula_module, "formula_surface_report", lambda data: surface)
    monkeypatch.setattr(
        native_formula_module.importlib.metadata, "version", lambda name: "9.9.9"
    )

    extraction = extract_formulas_with_native_kernel(b"unused", _scan())

    assert (2, 2) not in extraction.formulas.get("Data", {}), (
        "a failed A1 render must be absent, never a literal '=' formula"
    )


# --- Criterion 2: the Python native boundary must be total/degradable ------


def test_extract_formulas_with_native_kernel_degrades_on_an_out_of_range_definition_id(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Step 2. `_validated_r1c1_cells` indexes
    `sheet.definition_r1c1[definition_id]` with a bare Python list index and
    raises `IndexError` for a malformed/mismatched native return -- it must
    instead surface as `FormulaEnrichmentError` so the loader's existing
    `except FormulaEnrichmentError` degrade-gracefully clause catches it
    exactly like every other native-boundary failure.
    """
    surface = WorkbookFormulaSurface(
        sheets=(
            FormulaSheetSurface(
                name="Data",
                cell_rows=(0,),
                cell_cols=(0,),
                cell_a1=("A2",),
                cell_definition_id=(7,),  # out of range: definition_r1c1 has 1 entry
                definition_r1c1=("R[1]C",),
            ),
        ),
        defined_names=(),
    )
    monkeypatch.setattr(native_formula_module, "_xlsbkernel", object())
    monkeypatch.setattr(native_formula_module, "formula_surface_report", lambda data: surface)
    monkeypatch.setattr(
        native_formula_module.importlib.metadata, "version", lambda name: "9.9.9"
    )

    with pytest.raises(FormulaEnrichmentError):
        extract_formulas_with_native_kernel(b"unused", _scan())


def test_extract_formulas_with_native_kernel_degrades_on_mismatched_surface_vector_lengths(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Step 2. A native surface whose per-cell vectors disagree in length
    (a malformed or truncated PyO3 return) currently propagates a bare
    `ValueError` from `zip(..., strict=True)` instead of the shared
    `FormulaEnrichmentError` contract every other native-boundary failure
    uses.
    """
    surface = WorkbookFormulaSurface(
        sheets=(
            FormulaSheetSurface(
                name="Data",
                cell_rows=(0, 1),
                cell_cols=(0,),  # one element short
                cell_a1=("A2", "A3"),
                cell_definition_id=(0, 0),
                definition_r1c1=("R[1]C",),
            ),
        ),
        defined_names=(),
    )
    monkeypatch.setattr(native_formula_module, "_xlsbkernel", object())
    monkeypatch.setattr(native_formula_module, "formula_surface_report", lambda data: surface)
    monkeypatch.setattr(
        native_formula_module.importlib.metadata, "version", lambda name: "9.9.9"
    )

    with pytest.raises(FormulaEnrichmentError):
        extract_formulas_with_native_kernel(b"unused", _scan())


# --- Criterion 3: formula-cache v2 must preserve formulas_r1c1 --------------


def test_formula_cache_round_trip_preserves_canonical_r1c1_evidence(tmp_path: Path) -> None:
    """Step 3. Cache schema v1 never encodes `formulas_r1c1`
    (`_entry_blocks`/`_read_entry` in `qc_tool/io/formula_cache.py`), so a
    warm cache hit silently drops it even when the original extraction had
    it -- bypassing B5's validated-R1C1 fast path on every warm run.
    """
    extraction = FormulaExtraction(
        formulas={"Data": {(1, 1): "=A2+1", (2, 2): "=SUM(A1:A5)"}},
        engine="native-biff12:1.2.3",
        detail="native kernel",
        formulas_r1c1={"Data": {(1, 1): "=R[1]C", (2, 2): "=SUM(R[-1]C:R[-5]C)"}},
    )
    scan = _scan()
    key = FormulaCacheKey(
        package_sha256=package_digest(b"native-bytes"),
        coordinate_digest=coordinate_digest(scan),
        coordinate_count=scan.formula_count,
        adapter_family="native",
        adapter_fingerprint="native:xlsbkernel:1.2.3",
    )
    cache = FormulaExtractionCache(tmp_path / "formula-cache")

    cache.store(key, extraction)
    restored = cache.lookup(key)

    assert restored is not None
    assert restored.formulas_r1c1 == extraction.formulas_r1c1


# --- Criterion 10: attestation v4 must compose prior contracts -------------


def _population_finding(
    *, location: str = "B2:B16", shape_after_digest: str = "shape-a"
) -> Finding:
    first, last = location.split(":")
    membership = MembershipCodec(
        current_rectangles=(location,),
        baseline_mode="shift",
        shift=(-1, 0),
        member_count=15,
    )
    return Finding(
        artifact="excel",
        finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
        sheet="Data",
        location=location,
        element="population",
        message="15 cells summarised as one population",
        population=PopulationEvidence(
            member_count=15,
            membership=membership,
            first=first,
            last=last,
            shape_before_digest="a" * 64,
            shape_after_digest=shape_after_digest,
        ),
    )


def test_v4_verifier_accepts_a_malformed_population_manifest_entry_when_correctly_signed(
    fixture_dir: Path, tmp_path: Path
) -> None:
    """Step 7. `verify_attestation` only checks
    `isinstance(manifest["population_manifest"], list)`; it never validates
    each entry against `PopulationManifestEntry`'s own schema, so a
    structurally broken entry with a valid HMAC signature over the broken
    content verifies successfully.
    """
    result = QCRunResult(profile_name="pop", findings=[_population_finding()])
    key = b"z" * 32
    bundle = create_attestation(
        tmp_path / "pop.qca",
        result=result,
        profile=fixture_profile(),
        input_files={"current_excel": fixture_dir / "current.xlsx"},
        report_paths={},
        key=key,
    )

    with zipfile.ZipFile(bundle) as archive:
        manifest = json.loads(archive.read("manifest.json"))
        members = {
            name: archive.read(name) for name in archive.namelist() if name != "manifest.json"
        }

    assert manifest["schema_version"] == 4
    entry = manifest["population_manifest"][0]
    del entry["identity_key_digest"]
    entry["member_count"] = -1

    unsigned = {k: v for k, v in manifest.items() if k != "signature"}
    payload = json.dumps(unsigned, sort_keys=True, separators=(",", ":")).encode("utf-8")
    manifest["signature"]["value"] = hmac.new(key, payload, hashlib.sha256).hexdigest()

    tampered = tmp_path / "tampered.qca"
    with zipfile.ZipFile(tampered, "w") as archive:
        archive.writestr("manifest.json", json.dumps(manifest))
        for name, data in members.items():
            archive.writestr(name, data)

    verification = verify_attestation(tampered, key=key)

    assert not verification.valid, (
        "a malformed but correctly-signed v4 population-manifest entry must "
        "fail verification"
    )


# --- Criterion 6: population evidence tags/event identity ------------------


def _homogeneity_candidate(
    *,
    location: str,
    baseline_location: str,
    event_key: str = "",
    evidence_tags: set[FindingEvidenceTag] | None = None,
) -> Finding:
    return Finding.model_construct(
        artifact="excel",
        finding_class=FindingClass.FORMULA_LOGIC_CHANGED,
        sheet="Sheet1",
        location=location,
        baseline_location=baseline_location,
        baseline_value="=A1",
        current_value="=A2",
        event_key=event_key,
        evidence_tags=evidence_tags or set(),
        message=f"Sheet1!{location}: formula logic changed",
    )


def test_population_finding_preserves_a_homogeneous_event_key_and_evidence_tags() -> None:
    """Step 5. `_build_population` (`qc_tool/excel/population.py`) never
    copies `event_key`/`evidence_tags` from the representative candidate --
    a population always reports empty tags and an empty event key even
    when every member shares identical producer-authored evidence.
    """
    spill = CandidateSpill(default_profile(), _TODAY)
    tags = {FindingEvidenceTag.EXACT_WRAPPER}
    for row in range(2, 17):  # 15 members, all sharing one event key/tag set
        candidate = _homogeneity_candidate(
            location=f"B{row}",
            baseline_location=f"B{row - 1}",
            event_key="wrapper-rollout-1",
            evidence_tags=tags,
        )
        spill.add(candidate, shape_before="digest-before", shape_after="digest-after")

    policy = PopulationPolicy(enabled=True, threshold=10)
    outcome = finalize_populations(spill, policy, ComparisonScope())

    assert len(outcome.population_findings) == 1
    population = outcome.population_findings[0]
    assert population.event_key == "wrapper-rollout-1"
    assert population.evidence_tags == tags


def test_heterogeneous_evidence_tags_are_not_silently_merged_into_one_population() -> None:
    """Step 5. `population_key` deliberately excludes `evidence_tags` from
    the grouping key, but nothing then checks tag homogeneity before
    building one population `Finding` -- two subsets of members with
    different producer evidence silently collapse into one group with no
    tags at all.
    """
    spill = CandidateSpill(default_profile(), _TODAY)
    for row in range(2, 8):  # 6 members tagged A
        candidate = _homogeneity_candidate(
            location=f"C{row}",
            baseline_location=f"C{row - 1}",
            evidence_tags={FindingEvidenceTag.DISPLAY_EQUIVALENT},
        )
        spill.add(candidate, shape_before="digest-before", shape_after="digest-after")
    for row in range(8, 14):  # 6 members tagged B, otherwise identical
        candidate = _homogeneity_candidate(
            location=f"C{row}",
            baseline_location=f"C{row - 1}",
            evidence_tags={FindingEvidenceTag.ULP_SCALE},
        )
        spill.add(candidate, shape_before="digest-before", shape_after="digest-after")

    policy = PopulationPolicy(enabled=True, threshold=10)
    outcome = finalize_populations(spill, policy, ComparisonScope())

    assert outcome.population_findings == [], (
        "members with different producer evidence tags share one "
        "population_key and must not silently merge into one population"
    )
    assert len(outcome.replay_findings) == 12


# --- Criterion 8: Re-QC delta must use multiset semantics -------------------


def test_compare_findings_uses_multiset_semantics_for_duplicate_population_identities() -> None:
    """Step 5. `compare_findings` (`qc_tool/engine.py`) builds plain Python
    `set`s of `requeue_identity_key(finding)`. Two genuinely different
    populations that share one `population_identity_digest` (identical
    artifact/finding_class/sheet/shape digests/member, differing only in
    geometry -- which the digest deliberately excludes) collapse into one
    set entry, so losing one of the two is invisible to the delta.
    """
    previous = [
        _population_finding(location="B2:B16", shape_after_digest="shape-a"),
        _population_finding(location="D2:D16", shape_after_digest="shape-a"),
    ]
    current = [previous[0]]  # only the B-column population still exists

    delta = compare_findings(previous, current)

    assert delta.resolved == 1, (
        "one of two same-identity populations disappeared and must count as "
        "resolved, not be hidden by set-based delta accounting"
    )
    assert delta.persisting == 1
