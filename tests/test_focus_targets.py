"""Step 2 contracts: the focus target matrix and the private sidecar."""

import json
import sqlite3
import zipfile
from pathlib import Path

import pytest

from qc_tool.coverage import QCRunMode
from qc_tool.engine import QCRunResult
from qc_tool.findings import Finding, FindingClass, Severity
from qc_tool.focus.locator import (
    MAX_FOCUS_CELLS,
    normalize_address,
    parse_axis_span,
    parse_excel_location,
    parse_qualified_location,
)
from qc_tool.focus.model import (
    EMPTY_SIDECAR_JSON,
    FOCUS_SIDECAR_VERSION,
    FocusArtifact,
    FocusRole,
    FocusTargetSeed,
    FocusTargetSidecar,
    decode_focus_targets,
    encode_focus_targets,
)
from qc_tool.focus.targets import (
    FINDING_CLASS_RULES,
    FocusTargetContractError,
    build_focus_targets,
)
from qc_tool.history.store import RunHistory, export_runs_archive

ALL_HASHES = {
    "baseline_excel": "a" * 64,
    "current_excel": "b" * 64,
    "baseline_ppt": "c" * 64,
    "current_ppt": "d" * 64,
}


def _targets(
    findings: list[Finding],
    *,
    mode: QCRunMode = QCRunMode.CYCLE_COMPARISON,
    file_hashes: dict[str, str] | None = None,
) -> dict[str, tuple[FocusTargetSeed, ...]]:
    for index, finding in enumerate(findings, start=1):
        finding.finding_id = f"F{index:04d}"
        finding.severity = Severity.WARNING
    return build_focus_targets(
        findings, mode=mode, file_hashes=file_hashes or ALL_HASHES
    )


# --------------------------------------------------------------------------
# locator grammar
# --------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("B5", "B5"),
        ("$B$5", "B5"),
        ("a1", "A1"),
        ("A1:D6", "A1:D6"),
        ("D6:A1", "A1:D6"),
        ("XFD1048576", "XFD1048576"),
        ("5:5", "5:5"),
        ("C:C", "C:C"),
    ],
)
def test_normalize_address_accepts_bounded_references(raw: str, expected: str) -> None:
    assert normalize_address(raw) == expected


@pytest.mark.parametrize(
    "raw",
    [
        None,
        "",
        "sheet!A1",
        "A1:B2:C3",
        "XFE1",
        "A1048577",
        "A0",
        "row 5",
        "total revenue",
        "1:XFD",
        "A:5",
    ],
)
def test_normalize_address_refuses_non_addresses(raw: str | None) -> None:
    assert normalize_address(raw) is None


def test_normalize_address_bounds_the_selected_area() -> None:
    assert normalize_address("A:A") == "A:A"
    # Two whole columns exceed the one-column focus budget.
    assert normalize_address("A:B") is None
    assert MAX_FOCUS_CELLS == 1_048_576


def test_axis_span_wording_becomes_an_a1_span() -> None:
    assert parse_axis_span("row 5") == "5:5"
    assert parse_axis_span("column C") == "C:C"
    assert parse_axis_span("row zero") is None
    assert parse_axis_span("B5") is None
    assert parse_excel_location("row 12") == "12:12"
    assert parse_excel_location("B5") == "B5"


def test_qualified_location_splits_on_the_final_separator() -> None:
    assert parse_qualified_location("Summary!B2") == ("Summary", "B2")
    assert parse_qualified_location("KPI!Q3!B2") == ("KPI!Q3", "B2")
    assert parse_qualified_location("'Summary'!B2") == ("Summary", "B2")
    assert parse_qualified_location("Summary") is None
    assert parse_qualified_location("Summary!") is None
    assert parse_qualified_location("!B2") is None


# --------------------------------------------------------------------------
# seed model
# --------------------------------------------------------------------------


def test_excel_seed_requires_a_sheet_and_refuses_slide_fields() -> None:
    with pytest.raises(ValueError):
        FocusTargetSeed(
            artifact=FocusArtifact.EXCEL, role=FocusRole.CURRENT_EXCEL, address="A1"
        )
    with pytest.raises(ValueError):
        FocusTargetSeed(
            artifact=FocusArtifact.EXCEL,
            role=FocusRole.CURRENT_EXCEL,
            sheet="Summary",
            slide_index=2,
        )


def test_ppt_seed_requires_a_slide_index_and_refuses_sheet_fields() -> None:
    with pytest.raises(ValueError):
        FocusTargetSeed(artifact=FocusArtifact.PPT, role=FocusRole.CURRENT_PPT)
    with pytest.raises(ValueError):
        FocusTargetSeed(
            artifact=FocusArtifact.PPT,
            role=FocusRole.CURRENT_PPT,
            slide_index=1,
            sheet="Summary",
        )


def test_seed_role_and_artifact_must_agree() -> None:
    with pytest.raises(ValueError):
        FocusTargetSeed(
            artifact=FocusArtifact.EXCEL, role=FocusRole.CURRENT_PPT, slide_index=1
        )


# --------------------------------------------------------------------------
# target matrix
# --------------------------------------------------------------------------


def test_every_finding_class_has_an_explicit_rule() -> None:
    assert set(FINDING_CLASS_RULES) == set(FindingClass)


def test_changed_excel_cell_seeds_both_sides() -> None:
    finding = Finding(
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        sheet="Summary",
        location="B5",
        baseline_location="B4",
        message="value changed",
    )
    seeds = _targets([finding])["F0001"]
    assert {(seed.role, seed.sheet, seed.address) for seed in seeds} == {
        (FocusRole.CURRENT_EXCEL, "Summary", "B5"),
        (FocusRole.BASELINE_EXCEL, "Summary", "B4"),
    }


def test_deleted_row_seeds_only_the_baseline_side() -> None:
    finding = Finding(
        artifact="excel",
        finding_class=FindingClass.ROW_DELETED,
        sheet="Detail",
        baseline_location="row 8",
        message="row deleted",
    )
    seeds = _targets([finding])["F0001"]
    assert [(seed.role, seed.address) for seed in seeds] == [
        (FocusRole.BASELINE_EXCEL, "8:8")
    ]


def test_inserted_column_seeds_only_the_current_side() -> None:
    finding = Finding(
        artifact="excel",
        finding_class=FindingClass.COLUMN_INSERTED,
        sheet="Detail",
        location="column D",
        message="column inserted",
    )
    seeds = _targets([finding])["F0001"]
    assert [(seed.role, seed.address) for seed in seeds] == [
        (FocusRole.CURRENT_EXCEL, "D:D")
    ]


def test_sheet_level_evidence_produces_a_sheet_only_seed() -> None:
    added = Finding(
        artifact="excel",
        finding_class=FindingClass.SHEET_ADDED,
        sheet="New",
        message="sheet added",
    )
    removed = Finding(
        artifact="excel",
        finding_class=FindingClass.SHEET_REMOVED,
        sheet="Old",
        message="sheet removed",
    )
    generated = _targets([added, removed])
    assert [(seed.role, seed.sheet, seed.address) for seed in generated["F0001"]] == [
        (FocusRole.CURRENT_EXCEL, "New", None)
    ]
    assert [(seed.role, seed.sheet, seed.address) for seed in generated["F0002"]] == [
        (FocusRole.BASELINE_EXCEL, "Old", None)
    ]


def test_duplicate_key_marker_never_becomes_a_baseline_role() -> None:
    finding = Finding(
        artifact="excel",
        finding_class=FindingClass.DUPLICATE_KEY,
        sheet="Detail",
        location="A9:C9",
        baseline_location="row 4",
        message="duplicate key",
    )
    seeds = _targets([finding])["F0001"]
    assert [(seed.role, seed.address) for seed in seeds] == [
        (FocusRole.CURRENT_EXCEL, "A9:C9")
    ]


@pytest.mark.parametrize(
    "finding_class",
    [
        FindingClass.NAMED_RANGE_CHANGED,
        FindingClass.PIVOT_SOURCE_CHANGED,
        FindingClass.FINDINGS_CAPPED,
        FindingClass.CALCULATION_MODE,
        FindingClass.HIDDEN_CHANGED,
        FindingClass.PACKAGE_PERIOD_MISMATCH,
        FindingClass.PPT_REQUIRED_SLIDE_MISSING,
    ],
)
def test_untargetable_classes_produce_nothing(finding_class: FindingClass) -> None:
    finding = Finding(
        artifact="excel",
        finding_class=finding_class,
        sheet="Summary",
        location="B5",
        element="something",
        message="no structured target",
    )
    assert _targets([finding]) == {}


def test_matched_slide_findings_seed_both_decks() -> None:
    finding = Finding(
        artifact="ppt",
        finding_class=FindingClass.SLIDE_TEXT_CHANGED,
        slide="3 Revenue",
        slide_index=3,
        baseline_slide_index=2,
        message="text changed",
    )
    seeds = _targets([finding])["F0001"]
    assert {(seed.role, seed.slide_index) for seed in seeds} == {
        (FocusRole.CURRENT_PPT, 3),
        (FocusRole.BASELINE_PPT, 2),
    }


def test_added_and_removed_slides_seed_only_the_existing_side() -> None:
    added = Finding(
        artifact="ppt",
        finding_class=FindingClass.SLIDE_ADDED,
        slide="9 New",
        slide_index=9,
        message="slide added",
    )
    removed = Finding(
        artifact="ppt",
        finding_class=FindingClass.SLIDE_REMOVED,
        slide="4 Gone",
        baseline_slide_index=4,
        message="slide removed",
    )
    generated = _targets([added, removed])
    assert [(seed.role, seed.slide_index) for seed in generated["F0001"]] == [
        (FocusRole.CURRENT_PPT, 9)
    ]
    assert [(seed.role, seed.slide_index) for seed in generated["F0002"]] == [
        (FocusRole.BASELINE_PPT, 4)
    ]


def test_added_slide_with_both_indices_is_a_contract_violation() -> None:
    finding = Finding(
        artifact="ppt",
        finding_class=FindingClass.SLIDE_ADDED,
        slide="9 New",
        slide_index=9,
        baseline_slide_index=4,
        message="slide added",
    )
    with pytest.raises(FocusTargetContractError):
        _targets([finding])


def test_removed_slide_with_a_current_index_is_a_contract_violation() -> None:
    finding = Finding(
        artifact="ppt",
        finding_class=FindingClass.SLIDE_REMOVED,
        slide="4 Gone",
        slide_index=4,
        baseline_slide_index=4,
        message="slide removed",
    )
    with pytest.raises(FocusTargetContractError):
        _targets([finding])


def test_crosscheck_mismatch_seeds_current_excel_and_current_ppt() -> None:
    finding = Finding(
        artifact="crosscheck",
        finding_class=FindingClass.CROSSCHECK_MISMATCH,
        slide="2 Highlights",
        slide_index=2,
        element="Total revenue",
        location="Summary!B2",
        message="deck does not match Excel",
    )
    seeds = _targets([finding], mode=QCRunMode.FINAL_PACKAGE)["F0001"]
    assert {
        (seed.role, seed.sheet, seed.address, seed.slide_index) for seed in seeds
    } == {
        (FocusRole.CURRENT_EXCEL, "Summary", "B2", None),
        (FocusRole.CURRENT_PPT, None, None, 2),
    }


def test_unlocated_crosscheck_finding_has_no_target() -> None:
    finding = Finding(
        artifact="crosscheck",
        finding_class=FindingClass.CROSSCHECK_UNRESOLVED,
        slide="2 Highlights",
        element="Total revenue",
        message="mapped figure not found",
    )
    assert _targets([finding], mode=QCRunMode.FINAL_PACKAGE) == {}


def test_preflight_mode_never_creates_baseline_roles() -> None:
    finding = Finding(
        artifact="excel",
        finding_class=FindingClass.FORMULA_HARDCODED,
        sheet="Summary",
        location="B5",
        baseline_location="B5",
        message="hardcoded",
    )
    seeds = _targets(
        [finding],
        mode=QCRunMode.CURRENT_FILE_PREFLIGHT,
        file_hashes={"current_excel": "b" * 64},
    )["F0001"]
    assert [seed.role for seed in seeds] == [FocusRole.CURRENT_EXCEL]


def test_a_missing_role_hash_suppresses_that_seed() -> None:
    finding = Finding(
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        sheet="Summary",
        location="B5",
        baseline_location="B5",
        message="value changed",
    )
    seeds = _targets([finding], file_hashes={"current_excel": "b" * 64})["F0001"]
    assert [seed.role for seed in seeds] == [FocusRole.CURRENT_EXCEL]


def test_untriaged_findings_are_a_contract_violation() -> None:
    finding = Finding(
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        sheet="Summary",
        location="B5",
        message="value changed",
    )
    with pytest.raises(FocusTargetContractError):
        build_focus_targets(
            [finding], mode=QCRunMode.CYCLE_COMPARISON, file_hashes=ALL_HASHES
        )


def test_seeds_never_carry_paths_hashes_values_or_prose() -> None:
    finding = Finding(
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        sheet="Summary",
        location="B5",
        baseline_location="B5",
        baseline_value="1000",
        current_value="2000",
        element="Total revenue",
        message="value changed from 1000 to 2000",
    )
    payload = encode_focus_targets(_targets([finding]))
    for secret in ("1000", "2000", "Total revenue", "value changed", "b" * 64):
        assert secret not in payload


# --------------------------------------------------------------------------
# sidecar encoding
# --------------------------------------------------------------------------


def test_sidecar_round_trip_is_versioned() -> None:
    seed = FocusTargetSeed(
        artifact=FocusArtifact.EXCEL,
        role=FocusRole.CURRENT_EXCEL,
        sheet="Summary",
        address="B5",
    )
    encoded = encode_focus_targets({"F0001": (seed,)})
    assert json.loads(encoded)["version"] == FOCUS_SIDECAR_VERSION
    sidecar = decode_focus_targets(encoded)
    assert sidecar.usable
    assert sidecar.seed("F0001", FocusRole.CURRENT_EXCEL) == seed
    assert sidecar.seed("F0001", FocusRole.BASELINE_EXCEL) is None


def test_empty_generation_stores_the_legacy_default() -> None:
    assert encode_focus_targets({}) == EMPTY_SIDECAR_JSON
    assert decode_focus_targets(EMPTY_SIDECAR_JSON) == FocusTargetSidecar()


@pytest.mark.parametrize(
    "raw",
    [
        None,
        "",
        "not json",
        "[]",
        '{"version": 99, "targets": {}}',
        '{"version": 1, "targets": {"F0001": [{"artifact": "excel"}]}}',
        '{"version": 1, "unexpected": true}',
    ],
)
def test_unreadable_sidecars_yield_no_targets(raw: str | None) -> None:
    sidecar = decode_focus_targets(raw)
    assert not sidecar.usable
    assert sidecar.seeds("F0001") == ()


def test_ambiguous_role_seeds_are_not_offered() -> None:
    sidecar = FocusTargetSidecar(
        version=FOCUS_SIDECAR_VERSION,
        targets={
            "F0001": (
                FocusTargetSeed(
                    artifact=FocusArtifact.EXCEL,
                    role=FocusRole.CURRENT_EXCEL,
                    sheet="A",
                    address="B5",
                ),
                FocusTargetSeed(
                    artifact=FocusArtifact.EXCEL,
                    role=FocusRole.CURRENT_EXCEL,
                    sheet="B",
                    address="B6",
                ),
            )
        },
    )
    assert sidecar.seed("F0001", FocusRole.CURRENT_EXCEL) is None


# --------------------------------------------------------------------------
# history integration
# --------------------------------------------------------------------------


def _result() -> QCRunResult:
    finding = Finding(
        finding_id="F0001",
        artifact="excel",
        finding_class=FindingClass.VALUE_CHANGED,
        severity=Severity.CRITICAL,
        sheet="Summary",
        location="B5",
        message="value changed",
    )
    return QCRunResult(
        profile_name="default",
        mode=QCRunMode.CYCLE_COMPARISON,
        files={"current_excel": "current.xlsx"},
        findings=[finding],
    )


def test_history_round_trips_the_private_sidecar(tmp_path: Path) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    seed = FocusTargetSeed(
        artifact=FocusArtifact.EXCEL,
        role=FocusRole.CURRENT_EXCEL,
        sheet="Summary",
        address="B5",
    )
    run_id = history.record_run(
        _result(),
        file_hashes={"current_excel": "b" * 64},
        report_paths={},
        focus_targets={"F0001": (seed,)},
    )
    record = history.get_run(run_id)
    assert record.focus_targets.seed("F0001", FocusRole.CURRENT_EXCEL) == seed


def test_runs_recorded_without_targets_have_no_focus(tmp_path: Path) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    run_id = history.record_run(
        _result(), file_hashes={"current_excel": "b" * 64}, report_paths={}
    )
    record = history.get_run(run_id)
    assert not record.focus_targets.usable
    assert record.focus_targets.seeds("F0001") == ()


def test_legacy_database_gains_the_column(tmp_path: Path) -> None:
    db_path = tmp_path / "history.sqlite3"
    with sqlite3.connect(db_path) as conn:
        conn.executescript(
            """
            CREATE TABLE runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                started_at TEXT NOT NULL,
                profile TEXT NOT NULL,
                files TEXT NOT NULL,
                file_hashes TEXT NOT NULL,
                counts TEXT NOT NULL,
                disclosures TEXT NOT NULL,
                verified_crosschecks INTEGER NOT NULL,
                findings TEXT NOT NULL,
                report_paths TEXT NOT NULL
            );
            """
        )
    history = RunHistory(db_path)
    with sqlite3.connect(db_path) as conn:
        columns = {row[1] for row in conn.execute("PRAGMA table_info(runs)")}
    assert "focus_targets" in columns
    run_id = history.record_run(
        _result(), file_hashes={"current_excel": "b" * 64}, report_paths={}
    )
    assert not history.get_run(run_id).focus_targets.usable


def test_exported_archive_contains_no_focus_locators(tmp_path: Path) -> None:
    history = RunHistory(tmp_path / "history.sqlite3")
    managed = tmp_path / "runs"
    managed.mkdir()
    report = managed / "qc_report.html"
    report.write_text("<html></html>", encoding="utf-8")
    seed = FocusTargetSeed(
        artifact=FocusArtifact.EXCEL,
        role=FocusRole.CURRENT_EXCEL,
        sheet="ZzSecretSheet",
        address="B5",
    )
    run_id = history.record_run(
        _result(),
        file_hashes={"current_excel": "b" * 64},
        report_paths={"html": str(report)},
        focus_targets={"F0001": (seed,)},
    )
    destination = tmp_path / "export.zip"
    export_runs_archive(
        [history.get_run(run_id)], destination, managed_root=managed
    )
    payload = destination.read_bytes()
    assert b"ZzSecretSheet" not in payload
    assert b"focus_targets" not in payload
    with zipfile.ZipFile(destination) as bundle:
        manifest = json.loads(bundle.read("manifest.json"))
    assert "focus_targets" not in json.dumps(manifest)
