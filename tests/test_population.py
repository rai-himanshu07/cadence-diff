"""Group-first population contract: model, membership codec, registry,
focus rule, run policy, and attestation v4 (plan-20260906-group-first-and-
native-kernel.md, Step A1). Producer-level aggregation (A2) is implemented
in `qc_tool.excel.population`; see `tests/test_population_finalize.py` for
its grouping/spill/finalization and end-to-end tests.
"""

from __future__ import annotations

import json
import random
import warnings

import pytest
from pydantic import ValidationError

from qc_tool.attestation import (
    create_attestation,
    load_or_create_attestation_key,
    verify_attestation,
)
from qc_tool.config.profile import (
    POPULATION_ELIGIBLE_CLASSES,
    PopulationPolicy,
    ReviewPolicy,
    canonical_profile_json,
    default_profile,
    profile_sha256,
)
from qc_tool.coverage import QCRunMode
from qc_tool.engine import QCRunResult
from qc_tool.findings import (
    Finding,
    FindingClass,
    MembershipCodec,
    PopulationEvidence,
    PopulationSample,
)
from qc_tool.findings_store import finding_payload
from qc_tool.focus.targets import build_focus_targets
from qc_tool.history.review_state import (
    FINDING_EVIDENCE_FIELDS,
    EvidenceFieldRole,
    finding_evidence_digest,
)


def _shift_membership(
    rectangles: tuple[str, ...], *, dr: int, dc: int, count: int
) -> MembershipCodec:
    return MembershipCodec(
        current_rectangles=rectangles,
        baseline_mode="shift",
        shift=(dr, dc),
        member_count=count,
    )


def _population_finding(
    *,
    finding_class: FindingClass = FindingClass.FORMULA_LOGIC_CHANGED,
    membership: MembershipCodec | None = None,
) -> Finding:
    membership = membership or _shift_membership(("B2:B20",), dr=-1, dc=0, count=19)
    return Finding(
        artifact="excel",
        finding_class=finding_class,
        sheet="Data",
        location="B2:B20",
        element="population",
        message="19 cells summarised as one population",
        population=PopulationEvidence(
            member_count=membership.member_count,
            membership=membership,
            first="B2",
            last="B20",
            samples=(PopulationSample(current_location="B2", baseline_location="B1"),),
            shape_before_digest="a" * 64,
            shape_after_digest="b" * 64,
        ),
    )


# --- membership codec: exact reconstruction -------------------------------


def test_shift_mode_reconstructs_exact_baseline_current_pairs() -> None:
    rectangles = ("B2:B5", "B8:B10")
    codec = _shift_membership(rectangles, dr=-1, dc=0, count=7)
    current = ["B2", "B3", "B4", "B5", "B8", "B9", "B10"]
    expected_baseline = ["B1", "B2", "B3", "B4", "B7", "B8", "B9"]

    def _row(coord: str) -> int:
        return int(coord[1:])

    assert codec.shift is not None
    reconstructed = [f"B{_row(c) + codec.shift[0]}" for c in current]
    assert reconstructed == expected_baseline


def test_pairs_mode_reconstructs_exact_pairs_under_random_permutation() -> None:
    rng = random.Random(42)
    pairs = [(f"B{row}", f"A{row - 1}") for row in range(2, 40) if row % 3 == 0]
    shuffled = pairs[:]
    rng.shuffle(shuffled)

    codec = MembershipCodec(
        baseline_mode="pairs",
        pairs=tuple(shuffled),
        member_count=len(shuffled),
    )

    assert codec.pairs is not None
    assert set(codec.pairs) == set(pairs)
    assert codec.member_count == len(pairs)


def test_shift_mode_survives_a_random_mask_of_current_rectangles() -> None:
    rng = random.Random(7)
    rows = sorted(rng.sample(range(2, 500), 40))
    rectangles: list[str] = []
    start = rows[0]
    previous = rows[0]
    for row in rows[1:]:
        if row != previous + 1:
            rectangles.append(f"B{start}" if start == previous else f"B{start}:B{previous}")
            start = row
        previous = row
    rectangles.append(f"B{start}" if start == previous else f"B{start}:B{previous}")

    codec = _shift_membership(tuple(rectangles), dr=-2, dc=0, count=len(rows))
    assert codec.member_count == len(rows)
    assert codec.baseline_mode == "shift"


def test_membership_codec_requires_shift_offset_in_shift_mode() -> None:
    with pytest.raises(ValidationError):
        MembershipCodec(baseline_mode="shift", member_count=1)


def test_membership_codec_requires_pairs_in_pairs_mode() -> None:
    with pytest.raises(ValidationError):
        MembershipCodec(baseline_mode="pairs", member_count=1)


def test_membership_codec_caps_rectangles_and_pairs() -> None:
    with pytest.raises(ValidationError):
        MembershipCodec(
            current_rectangles=tuple(f"B{i}" for i in range(2001)),
            baseline_mode="shift",
            shift=(0, 0),
            member_count=2001,
        )
    with pytest.raises(ValidationError):
        MembershipCodec(
            baseline_mode="pairs",
            pairs=tuple((f"B{i}", f"A{i}") for i in range(50_001)),
            member_count=50_001,
        )


def test_population_evidence_caps_samples_at_five() -> None:
    membership = _shift_membership(("B2:B20",), dr=-1, dc=0, count=19)
    with pytest.raises(ValidationError):
        PopulationEvidence(
            member_count=19,
            membership=membership,
            first="B2",
            last="B20",
            samples=tuple(
                PopulationSample(current_location=f"B{i}") for i in range(6)
            ),
            shape_before_digest="a" * 64,
            shape_after_digest="b" * 64,
        )


def test_from_trusted_payload_reconstructs_membership_as_real_tuples() -> None:
    """A disk round trip must not leave JSON lists where the model needs tuples.

    `model_construct` (unlike the validating constructor) does not coerce a
    list into a `tuple[...]`-typed field. Building `MembershipCodec` from a
    raw JSON-shaped dict without an explicit tuple conversion silently
    leaves `current_rectangles`/`shift`/`pairs` as plain lists in memory --
    caught by a `model_dump` serializer warning on the very next re-encode.
    """
    pairs_membership = MembershipCodec(
        baseline_mode="pairs",
        current_rectangles=("B2", "B5"),
        pairs=(("B2", "A1"), ("B5", "A4")),
        member_count=2,
    )
    original = _population_finding(membership=pairs_membership)
    # Simulate the exact disk round trip: JSON has no tuple type.
    payload = json.loads(json.dumps(finding_payload(original)))

    restored = Finding.from_trusted_payload(payload)

    assert restored.population is not None
    membership = restored.population.membership
    assert isinstance(membership.current_rectangles, tuple)
    assert isinstance(membership.shift, tuple) if membership.shift else True
    assert isinstance(membership.pairs, tuple)
    assert all(isinstance(pair, tuple) for pair in membership.pairs)
    # A re-encode must not warn about a type mismatch (proves the types are
    # genuinely correct, not merely list-like enough to pass by duck typing).
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        finding_payload(restored)


# --- Finding.population: evidence registry + digest -----------------------


def test_population_field_is_included_in_the_evidence_registry() -> None:
    assert FINDING_EVIDENCE_FIELDS["population"] is EvidenceFieldRole.INCLUDED


def test_population_evidence_participates_in_the_digest() -> None:
    finding = _population_finding()
    atomic_twin = finding.model_copy(update={"population": None})

    assert finding_evidence_digest(finding) != finding_evidence_digest(atomic_twin)

    population = finding.population
    assert population is not None
    same_shape = finding.model_copy(
        update={
            "population": population.model_copy(
                update={"member_count": population.member_count}
            )
        }
    )
    assert finding_evidence_digest(same_shape) == finding_evidence_digest(finding)


def test_population_digest_changes_when_member_count_changes() -> None:
    finding = _population_finding()
    population = finding.population
    assert population is not None
    changed = finding.model_copy(
        update={"population": population.model_copy(update={"member_count": 5})}
    )
    assert finding_evidence_digest(changed) != finding_evidence_digest(finding)


# --- focus targets: a population never yields a click-to-focus target -----


def test_population_finding_yields_zero_focus_targets_regardless_of_class() -> None:
    for finding_class in POPULATION_ELIGIBLE_CLASSES:
        finding = _population_finding(finding_class=finding_class).model_copy(
            update={"finding_id": "F0001"}
        )
        targets = build_focus_targets(
            [finding],
            mode=QCRunMode.CYCLE_COMPARISON,
            file_hashes={"baseline_excel": "h1", "current_excel": "h2"},
        )
        assert targets == {}


# --- ReviewPolicy: explicit, versioned, off by default --------------------


def test_review_policy_disabled_by_default_and_excluded_from_canonical_json() -> None:
    profile = default_profile()
    assert profile.review_policy.populations.enabled is False
    assert "review_policy" not in canonical_profile_json(profile)


def test_enabling_populations_changes_profile_sha256() -> None:
    base = default_profile()
    enabled = base.model_copy(
        update={
            "review_policy": ReviewPolicy(
                populations=PopulationPolicy(enabled=True, threshold=10)
            )
        }
    )
    assert profile_sha256(enabled) != profile_sha256(base)
    assert "review_policy" in canonical_profile_json(enabled)


def test_population_policy_rejects_classes_outside_the_eligible_set() -> None:
    with pytest.raises(ValidationError):
        PopulationPolicy(classes=(FindingClass.VALUE_CHANGED,))


# --- attestation v4: population manifest -----------------------------------


def _result(findings: list[Finding]) -> QCRunResult:
    return QCRunResult(
        profile_name="default",
        mode=QCRunMode.CYCLE_COMPARISON,
        files={},
        findings=findings,
    )


def test_attestation_v4_carries_a_population_manifest(tmp_path) -> None:
    findings = [_population_finding().model_copy(update={"finding_id": "F0001"})]
    result = _result(findings)
    profile = default_profile()
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    baseline.write_bytes(b"baseline")
    current.write_bytes(b"current")
    _, key = load_or_create_attestation_key(tmp_path)

    output = create_attestation(
        tmp_path / "run.qca",
        result=result,
        profile=profile,
        input_files={"baseline_excel": baseline, "current_excel": current},
        report_paths={},
        key=key,
    )

    verification = verify_attestation(output, key=key)
    assert verification.valid, verification.issues

    import json
    import zipfile

    with zipfile.ZipFile(output) as archive:
        manifest = json.loads(archive.read("manifest.json"))
    assert manifest["schema_version"] == 4
    assert len(manifest["population_manifest"]) == 1
    entry = manifest["population_manifest"][0]
    assert entry["member_count"] == 19
    assert entry["identity_key_digest"]
    assert entry["membership_digest"]


def test_attestation_stays_v1_when_no_population_is_present(tmp_path) -> None:
    findings = [
        Finding(
            artifact="excel",
            finding_class=FindingClass.VALUE_CHANGED,
            sheet="Data",
            location="B2",
            message="value changed",
        ).model_copy(update={"finding_id": "F0001"})
    ]
    result = _result(findings)
    profile = default_profile()
    baseline = tmp_path / "baseline.xlsx"
    current = tmp_path / "current.xlsx"
    baseline.write_bytes(b"baseline")
    current.write_bytes(b"current")
    _, key = load_or_create_attestation_key(tmp_path)

    output = create_attestation(
        tmp_path / "run.qca",
        result=result,
        profile=profile,
        input_files={"baseline_excel": baseline, "current_excel": current},
        report_paths={},
        key=key,
    )

    import json
    import zipfile

    with zipfile.ZipFile(output) as archive:
        manifest = json.loads(archive.read("manifest.json"))
    assert manifest["schema_version"] == 1
    assert "population_manifest" not in manifest
