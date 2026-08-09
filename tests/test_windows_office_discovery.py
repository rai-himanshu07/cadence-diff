"""Contracts for the aggregate-only Windows Office discovery spike."""

import argparse
import json
import os
from pathlib import Path
from typing import cast

import pytest

import scripts.windows_office_discovery as discovery


def _reasons(record: dict[str, object]) -> list[str]:
    return cast(list[str], record["incomplete_reasons"])


def _window(**overrides: object) -> discovery.WindowFact:
    fields: dict[str, object] = {
        "application": discovery._EXCEL,
        "process_id": 100,
        "class_name": discovery._EXCEL_FRAME_CLASS,
        "visible": True,
        "minimized": False,
        "same_user": True,
        "same_session": True,
        "identity_known": True,
        "protected_view": False,
        "document_key": "key-a",
        "mechanism": discovery._MECHANISM_PYTHONCOM,
    }
    fields.update(overrides)
    return discovery.WindowFact(**fields)  # pyright: ignore[reportArgumentType]


def _document(**overrides: object) -> discovery.DocumentFact:
    fields: dict[str, object] = {
        "application": discovery._EXCEL,
        "document_key": "key-a",
        "process_id": 100,
        "window_count": 1,
        "visible_window_count": 1,
        "saved": True,
        "autosave": False,
        "path_kind": "local",
        "extension_kind": "xlsx",
        "hidden_instance": False,
        "is_addin": False,
        "visibility_unproved": False,
        "source": "window",
    }
    fields.update(overrides)
    return discovery.DocumentFact(**fields)  # pyright: ignore[reportArgumentType]


def _summarize(
    windows: list[discovery.WindowFact],
    documents: list[discovery.DocumentFact],
    *,
    expectation: discovery.Expectation,
    rot_only_documents: int = 0,
    native_om_available: bool = True,
    object_model_attempted: bool = True,
) -> dict[str, object]:
    return discovery.summarize_application(
        discovery._EXCEL,
        windows,
        documents,
        expectation=expectation,
        rot_only_documents=rot_only_documents,
        native_om_available=native_om_available,
        object_model_attempted=object_model_attempted,
    )


@pytest.mark.skipif(os.name == "nt", reason="asserts the non-Windows refusal path")
def test_discovery_refuses_non_windows_without_touching_com(tmp_path: Path) -> None:
    args = argparse.Namespace(
        scenario="baseline",
        output=tmp_path / "result.json",
        hidden_worker_file=None,
        spawn_hidden_worker=False,
    )

    exit_code, payload = discovery._run_discovery(args)

    assert exit_code == 2
    assert payload == {
        "schema_version": 1,
        "overall": "unsupported_platform",
        "passed": False,
        "scenario": "baseline",
    }


@pytest.mark.parametrize(
    ("full_name", "expected"),
    [
        (r"C:\Users\a\Book.xlsx", "local"),
        (r"\\server\share\Book.xlsx", "unc"),
        ("https://contoso.sharepoint.com/Book.xlsx", "url"),
        ("Book1", "unrecognized"),
        ("   ", "none"),
    ],
)
def test_path_kind_classifies_every_supported_location(
    full_name: str, expected: str
) -> None:
    assert discovery.classify_path_kind(full_name) == expected


@pytest.mark.parametrize(
    ("full_name", "expected"),
    [
        (r"C:\a\Book.XLSB", "xlsb"),
        (r"C:\a\Macro.xlsm", "xlsm"),
        (r"C:\a\Deck.pptx", "pptx"),
        (r"C:\a\Deck.pptm", "pptm"),
        (r"C:\a\notes.txt", "other"),
        ("Presentation1", "none"),
    ],
)
def test_extension_kind_classifies_office_formats(
    full_name: str, expected: str
) -> None:
    assert discovery.classify_extension(full_name) == expected


def test_exact_discovery_matching_declared_scenario_is_complete() -> None:
    record = _summarize(
        [_window(), _window(process_id=200, document_key="key-b")],
        [_document(), _document(document_key="key-b", process_id=200)],
        expectation=discovery.Expectation(processes=2, documents=2, windows=2),
    )

    assert record["coverage"] == "complete"
    assert record["verdict"] == discovery._VERDICT_COMPLETE
    assert record["fatal"] is False
    assert record["passed"] is True


def test_missing_document_under_complete_coverage_is_fatal() -> None:
    record = _summarize(
        [_window()],
        [_document()],
        expectation=discovery.Expectation(processes=1, documents=2, windows=1),
    )

    assert record["coverage"] == "complete"
    assert record["verdict"] == discovery._VERDICT_UNDERCOUNT
    assert record["fatal"] is True
    assert record["passed"] is False


def test_duplicate_alias_inflating_documents_is_fatal() -> None:
    record = _summarize(
        [_window(), _window(document_key="key-b")],
        [_document(), _document(document_key="key-b")],
        expectation=discovery.Expectation(documents=1),
    )

    assert record["verdict"] == discovery._VERDICT_OVERCOUNT
    assert record["fatal"] is True


def test_protected_view_window_refuses_instead_of_claiming_coverage() -> None:
    record = _summarize(
        [_window(), _window(protected_view=True, document_key=None)],
        [_document()],
        expectation=discovery.Expectation(documents=2),
    )

    assert record["coverage"] == "enumeration_incomplete"
    assert "protected_view_window" in _reasons(record)
    assert record["verdict"] == discovery._VERDICT_REFUSED
    assert record["fatal"] is False
    assert record["passed"] is True


def test_unreachable_visible_window_refuses_instead_of_undercounting() -> None:
    record = _summarize(
        [_window(), _window(document_key=None, mechanism=discovery._MECHANISM_NONE)],
        [_document()],
        expectation=discovery.Expectation(documents=2),
    )

    assert "document_window_object_model_unreachable" in _reasons(record)
    assert record["verdict"] == discovery._VERDICT_REFUSED


def test_running_object_table_document_missed_by_windows_refuses() -> None:
    record = _summarize(
        [_window()],
        [_document()],
        expectation=discovery.Expectation(documents=1),
        rot_only_documents=1,
    )

    assert "running_object_table_document_not_enumerated" in _reasons(record)
    assert record["coverage"] == "enumeration_incomplete"


def test_cross_session_or_cross_user_office_window_refuses() -> None:
    record = _summarize(
        [_window(), _window(same_user=False, same_session=False)],
        [_document()],
        expectation=discovery.Expectation(documents=1),
    )

    reasons = _reasons(record)
    assert "cross_user_office_window" in reasons
    assert "cross_session_office_window" in reasons


def test_uninspectable_process_refuses_before_other_checks() -> None:
    record = _summarize(
        [_window(identity_known=False, same_user=False, same_session=False)],
        [_document()],
        expectation=discovery.Expectation(documents=1),
    )

    reasons = _reasons(record)
    assert reasons == ["process_identity_unavailable"]


def test_hidden_qc_worker_is_counted_but_never_an_analyst_target() -> None:
    record = _summarize(
        [_window(), _window(visible=False, document_key="key-hidden")],
        [
            _document(),
            _document(
                document_key="key-hidden",
                hidden_instance=True,
                visible_window_count=0,
                autosave=None,
            ),
        ],
        expectation=discovery.Expectation(processes=1, documents=1, windows=1),
    )

    assert record["documents_discovered"] == 2
    assert record["hidden_instance_documents"] == 1
    assert record["analyst_documents"] == 1
    assert record["verdict"] == discovery._VERDICT_COMPLETE


def test_loaded_addins_are_discovered_but_never_analyst_documents() -> None:
    record = _summarize(
        [_window(), _window(visible=False, document_key="key-addin")],
        [
            _document(),
            _document(
                document_key="key-addin",
                is_addin=True,
                visible_window_count=0,
                window_count=0,
                saved=None,
                autosave=None,
            ),
        ],
        expectation=discovery.Expectation(processes=1, documents=1, windows=1),
    )

    assert record["documents_discovered"] == 2
    assert record["addin_documents"] == 1
    assert record["windowless_documents"] == 1
    assert record["analyst_documents"] == 1
    assert record["incomplete_reasons"] == []
    assert record["verdict"] == discovery._VERDICT_COMPLETE
    assert record["fatal"] is False


def test_workbook_open_without_a_visible_window_is_not_an_analyst_target() -> None:
    record = _summarize(
        [_window()],
        [_document(visible_window_count=0)],
        expectation=discovery.Expectation(documents=0),
    )

    assert record["documents_discovered"] == 1
    assert record["analyst_documents"] == 0
    assert record["verdict"] == discovery._VERDICT_COMPLETE


def test_unproved_window_visibility_refuses_instead_of_dropping_a_document() -> None:
    record = _summarize(
        [_window()],
        [_document(visible_window_count=0, visibility_unproved=True)],
        expectation=discovery.Expectation(processes=1, documents=1, windows=1),
    )

    assert "window_visibility_unproved" in _reasons(record)
    assert record["verdict"] == discovery._VERDICT_REFUSED
    assert record["fatal"] is False


def test_visible_frame_resolves_a_document_com_cannot_report() -> None:
    key: tuple[int, str] = (100, "key-a")
    documents: dict[tuple[int, str], discovery.DocumentFact] = {
        key: _document(visible_window_count=0, visibility_unproved=True)
    }

    resolved, from_frame, unresolved = discovery.resolve_frame_visibility(
        documents, {key: 2}
    )

    assert from_frame == 1
    assert unresolved == 0
    assert resolved[key].visible_window_count == 2
    assert resolved[key].visibility_unproved is False
    assert resolved[key].analyst_candidate is True


def test_document_with_no_visible_frame_stays_unproved() -> None:
    key: tuple[int, str] = (100, "key-a")
    documents: dict[tuple[int, str], discovery.DocumentFact] = {
        key: _document(visible_window_count=0, visibility_unproved=True)
    }

    resolved, from_frame, unresolved = discovery.resolve_frame_visibility(
        documents, {}
    )

    assert from_frame == 0
    assert unresolved == 1
    assert resolved[key].visibility_unproved is True
    assert resolved[key].analyst_candidate is False


def test_frame_fallback_never_overrides_a_reported_visibility() -> None:
    key: tuple[int, str] = (100, "key-a")
    documents: dict[tuple[int, str], discovery.DocumentFact] = {
        key: _document(visible_window_count=1)
    }

    resolved, from_frame, _unresolved = discovery.resolve_frame_visibility(
        documents, {key: 9}
    )

    assert from_frame == 0
    assert resolved[key].visible_window_count == 1


def test_undeclared_expectations_never_pass() -> None:
    record = _summarize(
        [_window()], [_document()], expectation=discovery.Expectation()
    )

    assert record["verdict"] == discovery._VERDICT_UNDECLARED
    assert record["passed"] is False
    assert record["fatal"] is False


def test_untested_application_with_nothing_open_is_not_applicable() -> None:
    record = _summarize([], [], expectation=discovery.Expectation())

    assert record["verdict"] == discovery._VERDICT_NOT_APPLICABLE
    assert record["passed"] is True
    assert record["fatal"] is False


def test_identical_file_open_in_two_processes_stays_ambiguous() -> None:
    record = _summarize(
        [_window(), _window(process_id=200)],
        [_document(), _document(process_id=200)],
        expectation=discovery.Expectation(processes=2, documents=2, windows=2),
    )

    assert record["analyst_documents"] == 2
    assert record["duplicate_path_analyst_documents"] == 2
    assert record["verdict"] == discovery._VERDICT_COMPLETE


def test_surplus_windows_are_reported_as_an_overcount() -> None:
    record = _summarize(
        [_window()],
        [_document(visible_window_count=2)],
        expectation=discovery.Expectation(processes=1, documents=1, windows=1),
    )

    assert record["verdict"] == discovery._VERDICT_OVERCOUNT
    assert record["fatal"] is True


def test_missing_evidence_outranks_a_surplus() -> None:
    record = _summarize(
        [_window()],
        [_document(visible_window_count=5)],
        expectation=discovery.Expectation(processes=1, documents=2, windows=1),
    )

    assert record["verdict"] == discovery._VERDICT_UNDERCOUNT


def test_unproved_saved_or_autosave_state_refuses() -> None:
    record = _summarize(
        [_window()],
        [_document(saved=None, autosave=None)],
        expectation=discovery.Expectation(documents=1),
    )

    reasons = _reasons(record)
    assert "saved_state_unproved" in reasons
    assert "autosave_state_unproved" in reasons


def test_absent_native_object_model_refuses() -> None:
    record = _summarize(
        [_window()],
        [_document()],
        expectation=discovery.Expectation(documents=1),
        native_om_available=False,
    )

    assert "native_object_model_unavailable" in _reasons(record)


def test_object_model_is_not_blamed_when_it_was_never_attempted() -> None:
    record = _summarize(
        [],
        [],
        expectation=discovery.Expectation(processes=0, documents=0, windows=0),
        native_om_available=False,
        object_model_attempted=False,
    )

    assert record["incomplete_reasons"] == []
    assert record["coverage"] == "complete"
    assert record["verdict"] == discovery._VERDICT_COMPLETE


def test_unstaged_application_refuses_instead_of_reporting_a_fatal_miss() -> None:
    record = _summarize(
        [],
        [],
        expectation=discovery.Expectation(processes=1, documents=1, windows=1),
        native_om_available=False,
        object_model_attempted=False,
    )

    assert "office_windows_not_found" in _reasons(record)
    assert record["verdict"] == discovery._VERDICT_REFUSED
    assert record["fatal"] is False


def test_diagnostics_expose_why_enumeration_returned_nothing() -> None:
    diagnostics = discovery.Diagnostics(
        top_level_windows_total=214,
        office_frame_windows=0,
        frame_class_counts={},
        running_object_table_entries=37,
        running_object_table_error="TypeError",
    )

    payload = diagnostics.as_payload()

    assert payload["top_level_windows_total"] == 214
    assert payload["office_frame_windows"] == 0
    assert payload["running_object_table_error"] == "TypeError"
    assert payload["object_model_attempts"] == 0
    assert "top_level_class_counts" not in payload


def test_window_class_dump_is_opt_in() -> None:
    diagnostics = discovery.Diagnostics(
        top_level_class_counts={"XLMAIN": 2, "Shell_TrayWnd": 1}
    )

    payload = diagnostics.as_payload()

    assert payload["top_level_class_counts"] == {"Shell_TrayWnd": 1, "XLMAIN": 2}


@pytest.mark.skipif(os.name == "nt", reason="asserts the non-Windows refusal path")
def test_running_object_table_reports_its_failure_code() -> None:
    documents, entries, error = discovery._running_object_table_documents()

    assert documents == {}
    assert entries == 0
    assert error == "unsupported_platform"


def test_overall_outcome_folds_every_application() -> None:
    complete = {"passed": True, "fatal": False, "coverage": "complete"}
    incomplete = {
        "passed": True,
        "fatal": False,
        "coverage": "enumeration_incomplete",
    }
    fatal = {"passed": False, "fatal": True, "coverage": "complete"}

    assert discovery.evaluate({"excel": complete, "powerpoint": complete}) == (
        "complete",
        True,
    )
    assert discovery.evaluate({"excel": complete, "powerpoint": incomplete}) == (
        "enumeration_incomplete",
        True,
    )
    assert discovery.evaluate({"excel": fatal, "powerpoint": complete}) == (
        "fatal_silent_miss",
        False,
    )
    assert discovery.evaluate({}) == ("failed", False)


def test_excel_document_windows_prefer_the_workbook_class(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        discovery,
        "_enumerate_child_windows",
        lambda _hwnd: [(2, "XLDESK"), (3, discovery._EXCEL_DOCUMENT_CLASS)],
    )

    candidates = discovery._document_windows(1, discovery._EXCEL)

    assert candidates == [(3, discovery._EXCEL_DOCUMENT_CLASS)]


def test_powerpoint_document_windows_try_frame_then_known_panes(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        discovery,
        "_enumerate_child_windows",
        lambda _hwnd: [(2, "SysHeader32"), (3, "mdiClass")],
    )

    candidates = discovery._document_windows(1, discovery._POWERPOINT)

    assert candidates[:2] == [
        (1, discovery._POWERPOINT_FRAME_CLASS),
        (3, "mdiClass"),
    ]
    assert (2, "SysHeader32") in candidates


def test_object_model_attempts_stay_bounded(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        discovery,
        "_enumerate_child_windows",
        lambda _hwnd: [(index, "SysHeader32") for index in range(500)],
    )

    candidates = discovery._document_windows(1, discovery._POWERPOINT)

    assert len(candidates) == discovery._MAX_OBJECT_MODEL_ATTEMPTS


@pytest.mark.skipif(os.name == "nt", reason="asserts the non-Windows refusal path")
def test_com_apartment_is_not_claimed_off_windows() -> None:
    assert discovery._com_initialize() is False


def test_document_keys_do_not_reveal_the_location() -> None:
    location = r"C:\Users\analyst\Deliverables\Q3 Revenue.xlsx"

    key = discovery._document_key(discovery._EXCEL, location)

    assert location not in key
    assert "revenue" not in key.casefold()
    assert key == discovery._document_key(discovery._EXCEL, location.upper())
    assert key != discovery._document_key(discovery._POWERPOINT, location)


def test_result_payload_carries_only_aggregate_values(tmp_path: Path) -> None:
    record = _summarize(
        [_window()],
        [_document()],
        expectation=discovery.Expectation(processes=1, documents=1, windows=1),
    )
    payload: dict[str, object] = {
        "schema_version": 1,
        "applications": {"excel": record},
        "overall": "complete",
        "passed": True,
        "scenario": "baseline",
    }
    output = tmp_path / "discovery.json"

    discovery._write_json(output, payload)
    text = output.read_text(encoding="utf-8")

    assert json.loads(text)["applications"]["excel"]["verdict"] == (
        discovery._VERDICT_COMPLETE
    )
    assert "key-a" not in text
    assert ":\\" not in text
    assert "\\\\" not in text
