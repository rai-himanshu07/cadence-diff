"""Step 3 contracts: discovery, package risk, exact byte matching, and binding."""

import datetime as dt
import hashlib
import zipfile
from pathlib import Path

import pytest

from qc_tool.focus.binding import (
    BINDING_TTL,
    BindingRegistry,
    BindingRequest,
    BindOutcome,
    FileChangedError,
    HashTimeoutError,
    hash_open_file,
    is_managed_copy,
    resolve_binding,
    revalidate_binding,
)
from qc_tool.focus.discovery import (
    DiscoveryReason,
    DiscoveryResult,
    FocusApplication,
    FrameObservation,
    OpenDocument,
    PathKind,
    assess_discovery,
    classify_path_kind,
    read_excel_document,
    read_powerpoint_document,
    read_window_collection,
    resolve_frame_visibility,
)
from qc_tool.focus.model import FocusRole
from qc_tool.focus.package_risk import (
    PackageRiskKind,
    PackageScanError,
    scan_focus_package,
)

# --------------------------------------------------------------------------
# mocked COM doubles
# --------------------------------------------------------------------------


class FakeWindow:
    def __init__(self, *, visible: bool | None = None, hwnd: int = 0) -> None:
        if visible is not None:
            self.Visible = visible
        if hwnd:
            self.HWND = hwnd


class FakeCollection:
    def __init__(self, items: list[object]) -> None:
        self._items = items
        self.Count = len(items)

    def Item(self, index: int) -> object:
        return self._items[index - 1]


class FakeApplication:
    def __init__(self, *, visible: bool = True, protected_views: int = 0) -> None:
        self.Visible = visible
        self.ProtectedViewWindows = FakeCollection([object()] * protected_views)


class FakeWorkbook:
    def __init__(
        self,
        full_name: str,
        *,
        saved: bool = True,
        autosave: bool = False,
        is_addin: bool = False,
        windows: list[object] | None = None,
    ) -> None:
        self.FullName = full_name
        self.Saved = saved
        self.AutoSaveOn = autosave
        self.IsAddin = is_addin
        self.Windows = FakeCollection(
            windows if windows is not None else [FakeWindow(visible=True, hwnd=101)]
        )


class FakeExcelWindowObject:
    def __init__(self, workbook: FakeWorkbook, application: FakeApplication) -> None:
        self.Parent = workbook
        self.Application = application


class FakePresentation:
    def __init__(
        self,
        full_name: str,
        *,
        saved: bool = True,
        autosave: bool = False,
        windows: list[object] | None = None,
    ) -> None:
        self.FullName = full_name
        self.Saved = saved
        self.AutoSaveOn = autosave
        self.Windows = FakeCollection(
            windows if windows is not None else [FakeWindow(hwnd=201)]
        )


class FakePptWindowObject:
    def __init__(
        self, presentation: FakePresentation, application: FakeApplication
    ) -> None:
        self.Presentation = presentation
        self.Application = application


def _always_visible(_handle: int) -> bool | None:
    return True


def _never_provable(_handle: int) -> bool | None:
    return None


def _excel(**kwargs: object) -> OpenDocument:
    defaults: dict[str, object] = {
        "application": FocusApplication.EXCEL,
        "process_id": 4242,
        "process_created": 1.0,
        "windows_session_id": 1,
        "full_name": r"C:\work\current.xlsx",
        "window_count": 1,
        "visible_window_count": 1,
        "visible_window_handles": (101,),
        "saved": True,
        "autosave": False,
    }
    defaults.update(kwargs)
    return OpenDocument(**defaults)  # type: ignore[arg-type]


# --------------------------------------------------------------------------
# document readers
# --------------------------------------------------------------------------


def test_excel_reader_reports_identity_and_protected_view() -> None:
    application = FakeApplication(protected_views=3)
    window = FakeExcelWindowObject(FakeWorkbook(r"C:\work\a.xlsx"), application)
    observed = read_excel_document(
        window,
        process_id=7,
        process_created=2.5,
        windows_session_id=1,
        window_visible=_always_visible,
    )
    assert observed is not None
    document, protected = observed
    assert protected is True
    assert document.process_id == 7
    assert document.process_created == 2.5
    assert document.visible_window_handles == (101,)
    assert document.analyst_candidate


def test_addin_and_windowless_workbooks_are_never_candidates() -> None:
    application = FakeApplication()
    addin = read_excel_document(
        FakeExcelWindowObject(FakeWorkbook(r"C:\x\a.xlam", is_addin=True), application),
        process_id=7,
        process_created=1.0,
        windows_session_id=1,
        window_visible=_always_visible,
    )
    windowless = read_excel_document(
        FakeExcelWindowObject(FakeWorkbook(r"C:\x\b.xlsx", windows=[]), application),
        process_id=7,
        process_created=1.0,
        windows_session_id=1,
        window_visible=_always_visible,
    )
    assert addin is not None and not addin[0].analyst_candidate
    assert windowless is not None and not windowless[0].analyst_candidate


def test_hidden_instance_documents_are_never_candidates() -> None:
    observed = read_excel_document(
        FakeExcelWindowObject(
            FakeWorkbook(r"C:\x\worker.xlsb"), FakeApplication(visible=False)
        ),
        process_id=9,
        process_created=1.0,
        windows_session_id=1,
        window_visible=_always_visible,
    )
    assert observed is not None
    assert observed[0].hidden_instance
    assert not observed[0].analyst_candidate


def test_powerpoint_window_visibility_falls_back_to_the_handle() -> None:
    observed = read_powerpoint_document(
        FakePptWindowObject(FakePresentation(r"C:\x\deck.pptx"), FakeApplication()),
        process_id=11,
        process_created=1.0,
        windows_session_id=1,
        window_visible=_always_visible,
    )
    assert observed is not None
    assert observed[0].visible_window_handles == (201,)
    assert not observed[0].visibility_unproved


def test_unprovable_window_state_is_unproved_not_dropped() -> None:
    observed = read_powerpoint_document(
        FakePptWindowObject(FakePresentation(r"C:\x\deck.pptx"), FakeApplication()),
        process_id=11,
        process_created=1.0,
        windows_session_id=1,
        window_visible=_never_provable,
    )
    assert observed is not None
    assert observed[0].visibility_unproved
    assert observed[0].visible_window_count == 0


def test_window_collection_counts_only_provable_visible_windows() -> None:
    total, visible, handles, unproved = read_window_collection(
        FakeCollection(
            [
                FakeWindow(visible=True, hwnd=1),
                FakeWindow(visible=False, hwnd=2),
                FakeWindow(hwnd=3),
            ]
        ),
        window_visible=_never_provable,
    )
    assert (total, visible, handles, unproved) == (3, 1, (1,), True)


def test_frame_visibility_repairs_an_unprovable_document() -> None:
    document = _excel(
        visible_window_count=0, visible_window_handles=(), visibility_unproved=True
    )
    repaired, unresolved = resolve_frame_visibility(
        [document], {document.instance_key: (555,)}
    )
    assert unresolved == 0
    assert repaired[0].visible_window_handles == (555,)
    assert repaired[0].visible_window_count == 1
    assert not repaired[0].visibility_unproved


def test_frame_visibility_leaves_a_truly_unprovable_document_unresolved() -> None:
    document = _excel(
        visible_window_count=0, visible_window_handles=(), visibility_unproved=True
    )
    repaired, unresolved = resolve_frame_visibility([document], {})
    assert unresolved == 1
    assert repaired[0].visibility_unproved


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        (r"C:\work\a.xlsx", PathKind.LOCAL),
        (r"\\server\share\a.xlsx", PathKind.UNC),
        ("https://contoso.sharepoint.com/a.xlsx", PathKind.URL),
        ("Book1", PathKind.UNRECOGNIZED),
        ("", PathKind.NONE),
    ],
)
def test_path_kinds(raw: str, expected: PathKind) -> None:
    assert classify_path_kind(raw) is expected


# --------------------------------------------------------------------------
# discovery completeness
# --------------------------------------------------------------------------


def _frame(**kwargs: object) -> FrameObservation:
    defaults: dict[str, object] = {
        "application": FocusApplication.EXCEL,
        "process_id": 4242,
        "visible": True,
        "identity_known": True,
        "same_user": True,
        "same_session": True,
        "protected_view": False,
        "reached_document": True,
    }
    defaults.update(kwargs)
    return FrameObservation(**defaults)  # type: ignore[arg-type]


def test_clean_discovery_is_complete() -> None:
    result = assess_discovery(
        FocusApplication.EXCEL,
        [_frame()],
        [_excel()],
        native_object_model_available=True,
    )
    assert result.complete
    assert len(result.candidates) == 1


@pytest.mark.parametrize(
    ("frame_kwargs", "reason"),
    [
        ({"identity_known": False}, DiscoveryReason.PROCESS_IDENTITY_UNAVAILABLE),
        ({"same_user": False}, DiscoveryReason.CROSS_USER_OFFICE_WINDOW),
        ({"same_session": False}, DiscoveryReason.CROSS_SESSION_OFFICE_WINDOW),
        ({"protected_view": True}, DiscoveryReason.PROTECTED_VIEW_WINDOW),
        ({"reached_document": False}, DiscoveryReason.DOCUMENT_WINDOW_UNREACHABLE),
    ],
)
def test_any_uninspectable_window_makes_discovery_incomplete(
    frame_kwargs: dict[str, object], reason: DiscoveryReason
) -> None:
    result = assess_discovery(
        FocusApplication.EXCEL,
        [_frame(**frame_kwargs)],
        [_excel()],
        native_object_model_available=True,
    )
    assert reason in result.reasons
    assert not result.complete


@pytest.mark.parametrize(
    ("document_kwargs", "reason"),
    [
        ({"saved": None}, DiscoveryReason.SAVED_STATE_UNPROVED),
        ({"autosave": None}, DiscoveryReason.AUTOSAVE_STATE_UNPROVED),
        (
            {"visibility_unproved": True},
            DiscoveryReason.WINDOW_VISIBILITY_UNPROVED,
        ),
    ],
)
def test_any_unreadable_attribute_refuses_instead_of_excluding(
    document_kwargs: dict[str, object], reason: DiscoveryReason
) -> None:
    result = assess_discovery(
        FocusApplication.EXCEL,
        [_frame()],
        [_excel(**document_kwargs)],
        native_object_model_available=True,
    )
    assert reason in result.reasons
    # The document is still discovered, never silently dropped.
    assert len(result.documents) == 1


def test_running_object_table_only_documents_refuse() -> None:
    result = assess_discovery(
        FocusApplication.EXCEL,
        [_frame()],
        [_excel()],
        native_object_model_available=True,
        rot_only_documents=1,
    )
    assert DiscoveryReason.ROT_DOCUMENT_NOT_ENUMERATED in result.reasons


def test_one_path_open_in_two_processes_stays_two_documents() -> None:
    first = _excel(process_id=1)
    second = _excel(process_id=2)
    assert first.instance_key != second.instance_key


# --------------------------------------------------------------------------
# package risk
# --------------------------------------------------------------------------


def _package(path: Path, entries: dict[str, bytes]) -> Path:
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        for name, payload in entries.items():
            archive.writestr(name, payload)
    return path


def test_clean_xlsx_is_safe_to_focus(tmp_path: Path) -> None:
    path = _package(tmp_path / "clean.xlsx", {"xl/workbook.xml": b"<workbook/>"})
    assert scan_focus_package(path).safe_to_focus


def test_clean_pptx_is_safe_to_focus(tmp_path: Path) -> None:
    path = _package(tmp_path / "clean.pptx", {"ppt/presentation.xml": b"<p/>"})
    assert scan_focus_package(path).safe_to_focus


@pytest.mark.parametrize(
    ("entries", "risk"),
    [
        ({"xl/vbaProject.bin": b"\x00"}, PackageRiskKind.VBA_PROJECT),
        ({"ppt/vbaProject.bin": b"\x00"}, PackageRiskKind.VBA_PROJECT),
        ({"xl/activeX/activeX1.xml": b"<a/>"}, PackageRiskKind.ACTIVEX_CONTROL),
        ({"ppt/embeddings/oleObject1.bin": b"\x00"}, PackageRiskKind.EMBEDDED_OLE),
        ({"customUI/customUI.xml": b"<c/>"}, PackageRiskKind.CUSTOM_OFFICE_UI),
        ({"xl/externalLinks/externalLink1.xml": b"<e/>"}, PackageRiskKind.EXTERNAL_DATA),
        ({"xl/macrosheets/sheet1.xml": b"<m/>"}, PackageRiskKind.MACRO_SHEET),
    ],
)
def test_risky_parts_refuse(
    tmp_path: Path, entries: dict[str, bytes], risk: PackageRiskKind
) -> None:
    suffix = ".pptx" if any(name.startswith("ppt/") for name in entries) else ".xlsx"
    path = _package(tmp_path / f"risky{suffix}", entries)
    scan = scan_focus_package(path)
    assert not scan.safe_to_focus
    assert risk in scan.risks


def test_external_relationship_refuses(tmp_path: Path) -> None:
    rels = (
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        b'relationships"><Relationship Id="rId1" Type="http://schemas.openxml'
        b'formats.org/officeDocument/2006/relationships/oleObject" '
        b'Target="../evil.bin" TargetMode="External"/></Relationships>'
    )
    path = _package(tmp_path / "ext.xlsx", {"xl/_rels/workbook.xml.rels": rels})
    scan = scan_focus_package(path)
    assert PackageRiskKind.EXTERNAL_RELATIONSHIP in scan.risks
    assert PackageRiskKind.EMBEDDED_OLE in scan.risks


def test_executable_relationship_target_refuses(tmp_path: Path) -> None:
    rels = (
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        b'relationships"><Relationship Id="rId1" Type="http://schemas.openxml'
        b'formats.org/officeDocument/2006/relationships/hyperlink" '
        b'Target="payload.exe" TargetMode="External"/></Relationships>'
    )
    path = _package(tmp_path / "exe.pptx", {"ppt/_rels/presentation.xml.rels": rels})
    scan = scan_focus_package(path)
    assert PackageRiskKind.UNKNOWN_EXECUTABLE_RELATIONSHIP in scan.risks


def test_traversal_relationship_target_refuses(tmp_path: Path) -> None:
    rels = (
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        b'relationships"><Relationship Id="rId1" Type="http://schemas.openxml'
        b'formats.org/officeDocument/2006/relationships/slide" '
        b'Target="../../../../etc/passwd"/></Relationships>'
    )
    path = _package(tmp_path / "trav.pptx", {"ppt/_rels/presentation.xml.rels": rels})
    assert PackageRiskKind.PATH_TRAVERSAL in scan_focus_package(path).risks


def test_malformed_relationship_metadata_refuses(tmp_path: Path) -> None:
    path = _package(tmp_path / "bad.xlsx", {"xl/_rels/workbook.xml.rels": b"<not-xml"})
    scan = scan_focus_package(path)
    assert PackageRiskKind.UNREADABLE_RELATIONSHIP_METADATA in scan.risks


def test_powerpoint_program_action_refuses(tmp_path: Path) -> None:
    slide = (
        b'<p:sld xmlns:p="p" xmlns:a="a"><a:hlinkClick action="ppaction://program"'
        b"/></p:sld>"
    )
    path = _package(tmp_path / "action.pptx", {"ppt/slides/slide1.xml": slide})
    assert PackageRiskKind.SLIDE_ACTION in scan_focus_package(path).risks


def test_powerpoint_slide_jump_action_is_allowed(tmp_path: Path) -> None:
    slide = (
        b'<p:sld xmlns:p="p" xmlns:a="a">'
        b'<a:hlinkClick action="ppaction://hlinkshowjump?jump=nextslide"/></p:sld>'
    )
    path = _package(tmp_path / "jump.pptx", {"ppt/slides/slide1.xml": slide})
    assert scan_focus_package(path).safe_to_focus


@pytest.mark.parametrize("suffix", [".xlsm", ".pptm", ".xls", ".ppt", ".xlam"])
def test_macro_enabled_formats_refuse(tmp_path: Path, suffix: str) -> None:
    path = tmp_path / f"book{suffix}"
    path.write_bytes(b"PK\x03\x04")
    assert scan_focus_package(path).error is PackageScanError.ACTIVE_CONTENT_FORMAT


def test_unknown_format_refuses(tmp_path: Path) -> None:
    path = tmp_path / "notes.txt"
    path.write_bytes(b"hello")
    assert scan_focus_package(path).error is PackageScanError.UNSUPPORTED_FORMAT


def test_encrypted_package_refuses(tmp_path: Path) -> None:
    path = tmp_path / "secret.xlsx"
    path.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 32)
    assert scan_focus_package(path).error is PackageScanError.ENCRYPTED_PACKAGE


def test_malformed_package_refuses(tmp_path: Path) -> None:
    path = tmp_path / "torn.pptx"
    path.write_bytes(b"PK\x03\x04not-a-zip")
    assert scan_focus_package(path).error is PackageScanError.MALFORMED_PACKAGE


def test_focus_risk_vocabulary_covers_the_xlsb_scanner() -> None:
    from qc_tool.focus import package_risk
    from qc_tool.io import xlsb_formula

    assert set(xlsb_formula._RISKY_PARTS) <= set(package_risk._RISKY_PARTS)
    assert set(xlsb_formula._RISKY_PATH_SEGMENTS) <= set(
        package_risk._RISKY_PATH_SEGMENTS
    )
    assert set(xlsb_formula._RISKY_RELATIONSHIP_KINDS) <= set(
        package_risk._RISKY_RELATIONSHIP_KINDS
    )


# --------------------------------------------------------------------------
# hashing and file identity
# --------------------------------------------------------------------------


def test_hash_open_file_reports_bytes_and_identity(tmp_path: Path) -> None:
    path = tmp_path / "a.bin"
    path.write_bytes(b"payload")
    digest, identity = hash_open_file(path)
    assert digest == hashlib.sha256(b"payload").hexdigest()
    assert identity.size == 7


def test_hash_open_file_detects_a_concurrent_change(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import os as os_module

    class FakeStat:
        def __init__(self, mtime_ns: int) -> None:
            self.st_dev = 1
            self.st_ino = 2
            self.st_size = 7
            self.st_mtime_ns = mtime_ns

    path = tmp_path / "a.bin"
    path.write_bytes(b"payload")
    calls = {"n": 0}

    def flaky(_fileno: int) -> object:
        calls["n"] += 1
        return FakeStat(1_000 + calls["n"])

    monkeypatch.setattr(os_module, "fstat", flaky)
    with pytest.raises(FileChangedError):
        hash_open_file(path)


def test_hash_open_file_respects_its_budget(tmp_path: Path) -> None:
    path = tmp_path / "a.bin"
    path.write_bytes(b"payload" * 1000)
    with pytest.raises(HashTimeoutError):
        hash_open_file(path, deadline_seconds=-1.0)


def test_managed_copies_are_detected_through_a_symlink(tmp_path: Path) -> None:
    managed_root = tmp_path / "data"
    (managed_root / "uploads").mkdir(parents=True)
    managed = managed_root / "uploads" / "current.xlsx"
    managed.write_bytes(b"payload")
    alias = tmp_path / "alias.xlsx"
    alias.symlink_to(managed)
    assert is_managed_copy(managed, managed_root)
    assert is_managed_copy(alias, managed_root, managed_path=managed)
    outside = tmp_path / "outside.xlsx"
    outside.write_bytes(b"payload")
    assert not is_managed_copy(outside, managed_root, managed_path=managed)


# --------------------------------------------------------------------------
# resolution precedence
# --------------------------------------------------------------------------


@pytest.fixture
def workspace(tmp_path: Path) -> tuple[Path, Path, str]:
    managed_root = tmp_path / "data"
    managed_root.mkdir()
    source = _package(tmp_path / "current.xlsx", {"xl/workbook.xml": b"<workbook/>"})
    digest = hashlib.sha256(source.read_bytes()).hexdigest()
    return managed_root, source, digest


def _request(managed_root: Path, digest: str, managed_path: Path | None = None):
    return BindingRequest(
        run_id=7,
        role=FocusRole.CURRENT_EXCEL,
        expected_sha256=digest,
        managed_root=managed_root,
        managed_path=managed_path,
    )


def _discovery(*documents: OpenDocument, reasons=()) -> DiscoveryResult:
    return DiscoveryResult(
        application=FocusApplication.EXCEL, documents=documents, reasons=reasons
    )


def test_sole_eligible_exact_match_binds(workspace) -> None:
    managed_root, source, digest = workspace
    outcome, document, identity = resolve_binding(
        _request(managed_root, digest), _discovery(_excel(full_name=str(source)))
    )
    assert outcome is BindOutcome.MATCHED
    assert document is not None and identity is not None


def test_incomplete_discovery_refuses_before_anything_else(workspace) -> None:
    managed_root, source, digest = workspace
    outcome, _document, _identity = resolve_binding(
        _request(managed_root, digest),
        _discovery(
            _excel(full_name=str(source)),
            reasons=(DiscoveryReason.CROSS_SESSION_OFFICE_WINDOW,),
        ),
    )
    assert outcome is BindOutcome.ENUMERATION_INCOMPLETE


def test_protected_view_is_reported_as_itself(workspace) -> None:
    managed_root, source, digest = workspace
    outcome, _document, _identity = resolve_binding(
        _request(managed_root, digest),
        _discovery(
            _excel(full_name=str(source)),
            reasons=(DiscoveryReason.PROTECTED_VIEW_WINDOW,),
        ),
    )
    assert outcome is BindOutcome.PROTECTED_VIEW


def test_same_bytes_open_twice_is_ambiguous_before_windows_are_inspected(
    workspace,
) -> None:
    managed_root, source, digest = workspace
    outcome, _document, _identity = resolve_binding(
        _request(managed_root, digest),
        _discovery(
            _excel(full_name=str(source), process_id=1),
            _excel(full_name=str(source), process_id=2, visible_window_count=4),
        ),
    )
    assert outcome is BindOutcome.AMBIGUOUS_EXACT_MATCHES


def test_same_name_different_bytes_is_no_match(workspace, tmp_path: Path) -> None:
    managed_root, _source, digest = workspace
    other = _package(tmp_path / "other.xlsx", {"xl/workbook.xml": b"<different/>"})
    outcome, _document, _identity = resolve_binding(
        _request(managed_root, digest), _discovery(_excel(full_name=str(other)))
    )
    assert outcome is BindOutcome.NO_EXACT_MATCH


def test_managed_upload_copy_is_refused(workspace, tmp_path: Path) -> None:
    managed_root, source, digest = workspace
    managed = managed_root / "current.xlsx"
    managed.write_bytes(source.read_bytes())
    outcome, _document, _identity = resolve_binding(
        _request(managed_root, digest, managed),
        _discovery(_excel(full_name=str(managed))),
    )
    assert outcome is BindOutcome.MANAGED_COPY_REFUSED


def test_autosave_enabled_and_unproved_both_refuse(workspace) -> None:
    managed_root, source, digest = workspace
    enabled, _d, _i = resolve_binding(
        _request(managed_root, digest),
        _discovery(_excel(full_name=str(source), autosave=True)),
    )
    unproved, _d2, _i2 = resolve_binding(
        _request(managed_root, digest),
        _discovery(_excel(full_name=str(source), autosave=None)),
    )
    assert enabled is BindOutcome.AUTOSAVE_ENABLED
    assert unproved is BindOutcome.AUTOSAVE_STATE_UNPROVED


def test_dirty_document_with_matching_bytes_still_binds(workspace) -> None:
    managed_root, source, digest = workspace
    outcome, document, _identity = resolve_binding(
        _request(managed_root, digest),
        _discovery(_excel(full_name=str(source), saved=False)),
    )
    assert outcome is BindOutcome.MATCHED
    assert document is not None and document.saved is False


def test_url_only_document_is_unsupported(workspace) -> None:
    managed_root, _source, digest = workspace
    outcome, _document, _identity = resolve_binding(
        _request(managed_root, digest),
        _discovery(_excel(full_name="https://contoso.sharepoint.com/a.xlsx")),
    )
    assert outcome is BindOutcome.UNSUPPORTED_LOCATION


def test_never_saved_document_is_unsupported(workspace) -> None:
    managed_root, _source, digest = workspace
    outcome, _document, _identity = resolve_binding(
        _request(managed_root, digest), _discovery(_excel(full_name="Book1"))
    )
    assert outcome is BindOutcome.UNSUPPORTED_LOCATION


def test_macro_enabled_document_is_refused(workspace, tmp_path: Path) -> None:
    managed_root, source, _digest = workspace
    macro = tmp_path / "current.xlsm"
    macro.write_bytes(source.read_bytes())
    digest = hashlib.sha256(macro.read_bytes()).hexdigest()
    outcome, _document, _identity = resolve_binding(
        _request(managed_root, digest), _discovery(_excel(full_name=str(macro)))
    )
    assert outcome is BindOutcome.UNSUPPORTED_ACTIVE_CONTENT


def test_risky_package_is_refused(workspace, tmp_path: Path) -> None:
    managed_root, _source, _digest = workspace
    risky = _package(tmp_path / "risky.xlsx", {"xl/vbaProject.bin": b"\x00"})
    digest = hashlib.sha256(risky.read_bytes()).hexdigest()
    outcome, _document, _identity = resolve_binding(
        _request(managed_root, digest), _discovery(_excel(full_name=str(risky)))
    )
    assert outcome is BindOutcome.UNSUPPORTED_ACTIVE_CONTENT


def test_multiple_windows_on_the_sole_match_are_ambiguous(workspace) -> None:
    managed_root, source, digest = workspace
    outcome, _document, _identity = resolve_binding(
        _request(managed_root, digest),
        _discovery(
            _excel(
                full_name=str(source),
                visible_window_count=2,
                visible_window_handles=(101, 102),
            )
        ),
    )
    assert outcome is BindOutcome.WINDOW_AMBIGUOUS


# --------------------------------------------------------------------------
# binding registry
# --------------------------------------------------------------------------


def test_binding_is_scoped_to_one_client_run_and_role(workspace) -> None:
    managed_root, source, digest = workspace
    registry = BindingRegistry()
    request = _request(managed_root, digest)
    _outcome, document, identity = resolve_binding(
        request, _discovery(_excel(full_name=str(source)))
    )
    assert document is not None and identity is not None
    registry.bind("client-a", request, document, identity)
    assert registry.get("client-a", 7, FocusRole.CURRENT_EXCEL) is not None
    assert registry.get("client-b", 7, FocusRole.CURRENT_EXCEL) is None
    assert registry.get("client-a", 8, FocusRole.CURRENT_EXCEL) is None
    assert registry.get("client-a", 7, FocusRole.BASELINE_EXCEL) is None


def test_binding_expires(workspace) -> None:
    managed_root, source, digest = workspace
    registry = BindingRegistry()
    request = _request(managed_root, digest)
    _outcome, document, identity = resolve_binding(
        request, _discovery(_excel(full_name=str(source)))
    )
    assert document is not None and identity is not None
    bound_at = dt.datetime(2026, 8, 5, 12, 0, tzinfo=dt.UTC)
    registry.bind("c", request, document, identity, now=bound_at)
    assert registry.get("c", 7, FocusRole.CURRENT_EXCEL, now=bound_at) is not None
    expired = bound_at + BINDING_TTL
    assert registry.get("c", 7, FocusRole.CURRENT_EXCEL, now=expired) is None


def test_disconnect_and_run_deletion_invalidate_bindings(workspace) -> None:
    managed_root, source, digest = workspace
    registry = BindingRegistry()
    request = _request(managed_root, digest)
    _outcome, document, identity = resolve_binding(
        request, _discovery(_excel(full_name=str(source)))
    )
    assert document is not None and identity is not None
    registry.bind("c1", request, document, identity)
    registry.bind("c2", request, document, identity)
    registry.invalidate_client("c1")
    assert registry.get("c1", 7, FocusRole.CURRENT_EXCEL) is None
    assert registry.get("c2", 7, FocusRole.CURRENT_EXCEL) is not None
    registry.invalidate_run(7)
    assert len(registry) == 0


def test_path_hmac_is_keyed_per_registry(workspace) -> None:
    _managed_root, source, _digest = workspace
    first = BindingRegistry()
    second = BindingRegistry()
    assert first.path_hmac(str(source)) != second.path_hmac(str(source))
    assert first.path_hmac(str(source)) == first.path_hmac(str(source))
    assert str(source) not in first.path_hmac(str(source))


def _bound(workspace, **document_kwargs):
    managed_root, source, digest = workspace
    registry = BindingRegistry()
    request = _request(managed_root, digest)
    document = _excel(full_name=str(source), **document_kwargs)
    _outcome, matched, identity = resolve_binding(request, _discovery(document))
    assert matched is not None and identity is not None
    binding = registry.bind("c", request, matched, identity)
    return registry, binding, document


def test_revalidation_accepts_an_unchanged_document(workspace) -> None:
    registry, binding, document = _bound(workspace)
    result = revalidate_binding(binding, _discovery(document), registry)
    assert result.bound


def test_revalidation_rejects_a_moved_process_or_window(workspace) -> None:
    registry, binding, document = _bound(workspace)
    from dataclasses import replace

    moved = replace(document, process_created=document.process_created + 1)
    assert (
        revalidate_binding(binding, _discovery(moved), registry).outcome
        is BindOutcome.BINDING_IDENTITY_CHANGED
    )
    rewindowed = replace(document, visible_window_handles=(999,))
    assert (
        revalidate_binding(binding, _discovery(rewindowed), registry).outcome
        is BindOutcome.BINDING_IDENTITY_CHANGED
    )


def test_revalidation_rejects_a_closed_document(workspace) -> None:
    registry, binding, _document = _bound(workspace)
    assert (
        revalidate_binding(binding, _discovery(), registry).outcome
        is BindOutcome.BINDING_IDENTITY_CHANGED
    )


def test_revalidation_reports_a_dirty_document_whose_bytes_changed(
    workspace,
) -> None:
    registry, binding, document = _bound(workspace, saved=False)
    Path(document.full_name).write_bytes(b"PK\x03\x04 different bytes")
    result = revalidate_binding(binding, _discovery(document), registry)
    assert result.outcome is BindOutcome.MATCHING_DOCUMENT_DIRTY_AND_CHANGED


def test_revalidation_reports_a_saved_document_whose_bytes_changed(
    workspace,
) -> None:
    registry, binding, document = _bound(workspace)
    Path(document.full_name).write_bytes(b"PK\x03\x04 different bytes")
    result = revalidate_binding(binding, _discovery(document), registry)
    assert result.outcome is BindOutcome.DOCUMENT_CHANGING


def test_revalidation_refuses_when_autosave_turned_on(workspace) -> None:
    registry, binding, document = _bound(workspace)
    from dataclasses import replace

    result = revalidate_binding(
        binding, _discovery(replace(document, autosave=True)), registry
    )
    assert result.outcome is BindOutcome.BINDING_IDENTITY_CHANGED


def test_bindings_are_never_persisted() -> None:
    import inspect

    from qc_tool.history import store

    source = inspect.getsource(store)
    for symbol in ("DocumentBinding", "BindingRegistry", "canonical_path_hmac"):
        assert symbol not in source
