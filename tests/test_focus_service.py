"""Step 7 contracts: startup guards, action tokens, and role actions."""

import datetime as dt
from pathlib import Path

import pytest

from qc_tool.coverage import QCRunMode
from qc_tool.findings import Finding, FindingClass, Severity
from qc_tool.focus.binding import BindOutcome
from qc_tool.focus.discovery import DiscoveryResult, FocusApplication, OpenDocument
from qc_tool.focus.model import (
    FOCUS_SIDECAR_VERSION,
    FocusArtifact,
    FocusRole,
    FocusTargetSeed,
    FocusTargetSidecar,
)
from qc_tool.focus.navigator import FocusNavigator, HelperRun
from qc_tool.focus.protocol import SCHEMA_VERSION, FocusAction, FocusOutcome
from qc_tool.focus.service import (
    TOKEN_TTL,
    ActionClaim,
    FocusService,
    FocusUnavailable,
    TokenRejection,
)
from qc_tool.history.store import RunRecord
from qc_tool.server_config import NetworkMode

CLIENT = "client-1"


def _record(
    *,
    hashes: dict[str, str] | None = None,
    seeds: dict[str, tuple[FocusTargetSeed, ...]] | None = None,
    paths: dict[str, str] | None = None,
) -> RunRecord:
    return RunRecord(
        run_id=7,
        started_at=dt.datetime(2026, 8, 5, tzinfo=dt.UTC),
        profile="default",
        mode=QCRunMode.CYCLE_COMPARISON,
        files={"current_excel": "current.xlsx", "baseline_excel": "baseline.xlsx"},
        file_hashes=hashes if hashes is not None else {
            "current_excel": "a" * 64,
            "baseline_excel": "b" * 64,
        },
        counts={},
        review_counts={},
        disclosures=[],
        verified_crosschecks=0,
        report_paths={},
        file_paths=paths or {},
        findings=[
            Finding(
                finding_id="F0001",
                artifact="excel",
                finding_class=FindingClass.VALUE_CHANGED,
                severity=Severity.CRITICAL,
                sheet="Summary",
                location="B5",
                message="value changed",
            )
        ],
        focus_targets=FocusTargetSidecar(
            version=FOCUS_SIDECAR_VERSION,
            targets=seeds
            if seeds is not None
            else {
                "F0001": (
                    FocusTargetSeed(
                        artifact=FocusArtifact.EXCEL,
                        role=FocusRole.CURRENT_EXCEL,
                        sheet="Summary",
                        address="B5",
                    ),
                    FocusTargetSeed(
                        artifact=FocusArtifact.EXCEL,
                        role=FocusRole.BASELINE_EXCEL,
                        sheet="Summary",
                        address="B5",
                    ),
                )
            },
        ),
    )


def _service(tmp_path: Path, **kwargs) -> FocusService:
    options: dict[str, object] = {
        "enabled": True,
        "platform": "win32",
        "network_mode": NetworkMode.LOCAL,
    }
    options.update(kwargs)
    return FocusService(tmp_path, **options)  # type: ignore[arg-type]


# --------------------------------------------------------------------------
# availability
# --------------------------------------------------------------------------


def test_focus_is_off_by_default(tmp_path: Path) -> None:
    service = FocusService(tmp_path, enabled=False, platform="win32")
    assert service.availability() is FocusUnavailable.FEATURE_OFF
    assert not service.available


def test_focus_refuses_off_windows(tmp_path: Path) -> None:
    service = _service(tmp_path, platform="linux")
    assert service.availability() is FocusUnavailable.UNSUPPORTED_PLATFORM


def test_focus_refuses_in_lan_mode(tmp_path: Path) -> None:
    service = _service(tmp_path, network_mode=NetworkMode.LAN)
    assert service.availability() is FocusUnavailable.NETWORK_NOT_LOOPBACK


def test_focus_is_available_only_when_all_three_hold(tmp_path: Path) -> None:
    assert _service(tmp_path).available


async def test_disabling_focus_revokes_all_client_authority(
    tmp_path: Path,
    open_workbook: tuple[Path, str],
) -> None:
    source, digest = open_workbook
    navigator, _seen = _navigator(
        [
            HelperRun(
                payload={
                    "outcome": FocusOutcome.DISCOVERED.value,
                    "discovery": _payload(_document(source)),
                }
            )
        ]
    )
    service = _service(tmp_path / "data", navigator=navigator)
    record = _record(hashes={"current_excel": digest})
    assert (await service.bind(CLIENT, record, FocusRole.CURRENT_EXCEL)).offered
    assert service.confirm(
        CLIENT,
        record,
        FocusRole.CURRENT_EXCEL,
    ) is BindOutcome.MATCHED
    service.acknowledge(CLIENT)
    token = service.issue_token(CLIENT, 7, "F0001", FocusRole.CURRENT_EXCEL)

    service.set_enabled(False)

    assert service.availability() is FocusUnavailable.FEATURE_OFF
    assert service.binding(CLIENT, 7, FocusRole.CURRENT_EXCEL) is None
    assert not service.acknowledged(CLIENT)
    assert service.consume_token(CLIENT, token) is TokenRejection.UNKNOWN_TOKEN
    assert service.confirm(
        CLIENT,
        record,
        FocusRole.CURRENT_EXCEL,
    ) is BindOutcome.NO_EXACT_MATCH

    service.set_enabled(True)
    assert service.available
    assert service.binding(CLIENT, 7, FocusRole.CURRENT_EXCEL) is None


# --------------------------------------------------------------------------
# targets and role suppression
# --------------------------------------------------------------------------


def test_both_roles_are_offered_when_each_has_a_hash(tmp_path: Path) -> None:
    service = _service(tmp_path)
    roles = {seed.role for seed in service.seeds(_record(), "F0001")}
    assert roles == {FocusRole.CURRENT_EXCEL, FocusRole.BASELINE_EXCEL}


def test_a_missing_role_hash_removes_that_action(tmp_path: Path) -> None:
    service = _service(tmp_path)
    record = _record(hashes={"current_excel": "a" * 64})
    roles = {seed.role for seed in service.seeds(record, "F0001")}
    assert roles == {FocusRole.CURRENT_EXCEL}


def test_identical_baseline_and_current_hashes_suppress_both_roles(
    tmp_path: Path,
) -> None:
    service = _service(tmp_path)
    same = {"current_excel": "a" * 64, "baseline_excel": "a" * 64}
    assert service.seeds(_record(hashes=same), "F0001") == ()


def test_a_legacy_run_has_no_actions(tmp_path: Path) -> None:
    service = _service(tmp_path)
    record = _record()
    record.focus_targets = FocusTargetSidecar()
    assert service.seeds(record, "F0001") == ()


def test_an_ambiguous_role_seed_is_not_actionable(tmp_path: Path) -> None:
    service = _service(tmp_path)
    duplicated = (
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
    record = _record(seeds={"F0001": duplicated})
    assert service.seed(record, "F0001", FocusRole.CURRENT_EXCEL) is None


def test_identical_hash_suppression_is_scoped_to_one_member(tmp_path: Path) -> None:
    service = _service(tmp_path)
    seeds = tuple(
        FocusTargetSeed(
            artifact=FocusArtifact.EXCEL,
            role=role,
            member_id=member_id,
            sheet="Data",
            address="A1",
        )
        for member_id in ("core", "ops")
        for role in (FocusRole.BASELINE_EXCEL, FocusRole.CURRENT_EXCEL)
    )
    record = _record(
        hashes={
            "baseline_excel:core": "a" * 64,
            "current_excel:core": "a" * 64,
            "baseline_excel:ops": "b" * 64,
            "current_excel:ops": "c" * 64,
        },
        seeds={"F0001": seeds},
    )

    actionable = service.seeds(record, "F0001")

    assert {(seed.role, seed.member_id) for seed in actionable} == {
        (FocusRole.BASELINE_EXCEL, "ops"),
        (FocusRole.CURRENT_EXCEL, "ops"),
    }


# --------------------------------------------------------------------------
# action tokens
# --------------------------------------------------------------------------


def test_a_token_is_single_use(tmp_path: Path) -> None:
    service = _service(tmp_path)
    token = service.issue_token(CLIENT, 7, "F0001", FocusRole.CURRENT_EXCEL)
    assert isinstance(service.consume_token(CLIENT, token), ActionClaim)
    assert service.consume_token(CLIENT, token) is TokenRejection.UNKNOWN_TOKEN


def test_action_token_authorizes_one_exact_member(tmp_path: Path) -> None:
    service = _service(tmp_path)
    token = service.issue_token(
        CLIENT,
        7,
        "F0001",
        FocusRole.CURRENT_EXCEL,
        member_id="ops",
    )

    claim = service.consume_token(CLIENT, token)

    assert isinstance(claim, ActionClaim)
    assert claim.role is FocusRole.CURRENT_EXCEL
    assert claim.member_id == "ops"


def test_a_token_expires(tmp_path: Path) -> None:
    service = _service(tmp_path)
    issued = dt.datetime(2026, 8, 5, 12, 0, tzinfo=dt.UTC)
    token = service.issue_token(
        CLIENT, 7, "F0001", FocusRole.CURRENT_EXCEL, now=issued
    )
    assert (
        service.consume_token(CLIENT, token, now=issued + TOKEN_TTL)
        is TokenRejection.EXPIRED_TOKEN
    )


def test_a_token_from_another_client_is_rejected(tmp_path: Path) -> None:
    service = _service(tmp_path)
    token = service.issue_token(CLIENT, 7, "F0001", FocusRole.CURRENT_EXCEL)
    assert service.consume_token("other", token) is TokenRejection.UNKNOWN_TOKEN


def test_a_rerender_retires_every_earlier_token(tmp_path: Path) -> None:
    service = _service(tmp_path)
    token = service.issue_token(CLIENT, 7, "F0001", FocusRole.CURRENT_EXCEL)
    service.new_revision(CLIENT)
    assert service.consume_token(CLIENT, token) is TokenRejection.UNKNOWN_TOKEN


def test_disconnect_and_deletion_clear_tokens(tmp_path: Path) -> None:
    service = _service(tmp_path)
    token = service.issue_token(CLIENT, 7, "F0001", FocusRole.CURRENT_EXCEL)
    service.forget_run(7)
    assert service.consume_token(CLIENT, token) is TokenRejection.UNKNOWN_TOKEN
    other = service.issue_token(CLIENT, 9, "F0001", FocusRole.CURRENT_EXCEL)
    service.forget_client(CLIENT)
    assert service.consume_token(CLIENT, other) is TokenRejection.UNKNOWN_TOKEN


def test_side_effect_acknowledgement_is_per_client_and_sticky(tmp_path: Path) -> None:
    service = _service(tmp_path)
    assert not service.acknowledged(CLIENT)
    service.acknowledge(CLIENT)
    assert service.acknowledged(CLIENT)
    assert not service.acknowledged("other")


# --------------------------------------------------------------------------
# bind and focus dispatch
# --------------------------------------------------------------------------


def _document(
    source: Path,
    *,
    application: FocusApplication = FocusApplication.EXCEL,
) -> OpenDocument:
    return OpenDocument(
        application=application,
        process_id=42,
        process_created=1.0,
        windows_session_id=1,
        full_name=str(source),
        window_count=1,
        visible_window_count=1,
        visible_window_handles=(101,),
        saved=True,
        autosave=False,
    )


def _payload(document: OpenDocument) -> dict[str, object]:
    from dataclasses import asdict

    item = asdict(document)
    item["application"] = document.application.value
    return {
        "application": document.application.value,
        "documents": [item],
        "reasons": [],
    }


@pytest.fixture
def open_workbook(tmp_path: Path) -> tuple[Path, str]:
    import hashlib
    import zipfile

    source = tmp_path / "current.xlsx"
    with zipfile.ZipFile(source, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("xl/workbook.xml", "<workbook/>")
    return source, hashlib.sha256(source.read_bytes()).hexdigest()


def _navigator(replies: list[HelperRun]) -> tuple[FocusNavigator, list[dict]]:
    seen: list[dict] = []

    def runner(request: dict[str, object], _timeout: float) -> HelperRun:
        seen.append(request)
        return replies[min(len(seen) - 1, len(replies) - 1)]

    return FocusNavigator(runner=runner), seen


async def test_bind_offers_one_document_and_confirmation_creates_the_binding(
    tmp_path: Path, open_workbook: tuple[Path, str]
) -> None:
    source, digest = open_workbook
    navigator, _seen = _navigator(
        [
            HelperRun(
                payload={
                    "outcome": FocusOutcome.DISCOVERED.value,
                    "discovery": _payload(_document(source)),
                }
            )
        ]
    )
    service = _service(tmp_path / "data", navigator=navigator)
    record = _record(hashes={"current_excel": digest})
    report = await service.bind(CLIENT, record, FocusRole.CURRENT_EXCEL)
    assert report.offered
    assert report.folder_label == source.parent.name
    assert service.binding(CLIENT, 7, FocusRole.CURRENT_EXCEL) is None
    assert (
        service.confirm(CLIENT, record, FocusRole.CURRENT_EXCEL)
        is BindOutcome.MATCHED
    )
    assert service.binding(CLIENT, 7, FocusRole.CURRENT_EXCEL) is not None


async def test_non_primary_binding_never_occupies_primary_role_slot(
    tmp_path: Path,
    open_workbook: tuple[Path, str],
) -> None:
    source, digest = open_workbook
    navigator, _seen = _navigator(
        [
            HelperRun(
                payload={
                    "outcome": FocusOutcome.DISCOVERED.value,
                    "discovery": _payload(_document(source)),
                }
            )
        ]
    )
    service = _service(tmp_path / "data", navigator=navigator)
    seed = FocusTargetSeed(
        artifact=FocusArtifact.EXCEL,
        role=FocusRole.CURRENT_EXCEL,
        member_id="ops",
        sheet="Summary",
        address="B5",
    )
    record = _record(
        hashes={"current_excel:ops": digest},
        seeds={"F0001": (seed,)},
    )

    report = await service.bind(
        CLIENT,
        record,
        FocusRole.CURRENT_EXCEL,
        "ops",
    )

    assert report.offered
    assert report.role_key == "current_excel:ops"
    assert service.confirm(
        CLIENT,
        record,
        FocusRole.CURRENT_EXCEL,
        "ops",
    ) is BindOutcome.MATCHED
    assert service.binding(CLIENT, 7, FocusRole.CURRENT_EXCEL) is None
    assert service.binding(CLIENT, 7, FocusRole.CURRENT_EXCEL, "ops") is not None


async def test_focus_requires_a_confirmed_binding(
    tmp_path: Path, open_workbook: tuple[Path, str]
) -> None:
    source, digest = open_workbook
    navigator, _seen = _navigator(
        [
            HelperRun(
                payload={
                    "outcome": FocusOutcome.DISCOVERED.value,
                    "discovery": _payload(_document(source)),
                }
            )
        ]
    )
    service = _service(tmp_path / "data", navigator=navigator)
    service.acknowledge(CLIENT)
    record = _record(hashes={"current_excel": digest})
    claim = ActionClaim(
        run_id=7,
        finding_id="F0001",
        role=FocusRole.CURRENT_EXCEL,
        revision=0,
        issued_at=dt.datetime.now(dt.UTC),
    )
    reply = await service.focus(CLIENT, record, claim)
    assert reply.code == BindOutcome.BINDING_EXPIRED.value


async def test_focus_requires_the_side_effect_acknowledgement(
    tmp_path: Path, open_workbook: tuple[Path, str]
) -> None:
    _source, digest = open_workbook
    navigator, _seen = _navigator([HelperRun(payload={"outcome": "focused"})])
    service = _service(tmp_path / "data", navigator=navigator)
    record = _record(hashes={"current_excel": digest})
    claim = ActionClaim(
        run_id=7,
        finding_id="F0001",
        role=FocusRole.CURRENT_EXCEL,
        revision=0,
        issued_at=dt.datetime.now(dt.UTC),
    )
    reply = await service.focus(CLIENT, record, claim)
    assert reply.outcome is FocusOutcome.ACTION_UNAVAILABLE


async def test_focus_sends_only_fixed_identity_and_never_a_path(
    tmp_path: Path, open_workbook: tuple[Path, str]
) -> None:
    source, digest = open_workbook
    discovery = HelperRun(
        payload={
            "outcome": FocusOutcome.DISCOVERED.value,
            "discovery": _payload(_document(source)),
        }
    )
    navigator, seen = _navigator([discovery, discovery, discovery])
    service = _service(tmp_path / "data", navigator=navigator)
    service.acknowledge(CLIENT)
    record = _record(hashes={"current_excel": digest})
    await service.bind(CLIENT, record, FocusRole.CURRENT_EXCEL)
    service.confirm(CLIENT, record, FocusRole.CURRENT_EXCEL)
    claim = ActionClaim(
        run_id=7,
        finding_id="F0001",
        role=FocusRole.CURRENT_EXCEL,
        revision=0,
        issued_at=dt.datetime.now(dt.UTC),
    )
    await service.focus(CLIENT, record, claim)
    focus_request = [
        item for item in seen if item.get("action") == FocusAction.FOCUS.value
    ]
    assert len(focus_request) == 1
    request = focus_request[0]
    assert request["schema_version"] == SCHEMA_VERSION
    assert request["sheet"] == "Summary"
    assert request["address"] == "B5"
    assert request["shape_id"] is None
    assert request["process_id"] == 42
    assert request["window_handle"] == 101
    assert str(source) not in repr(request)
    assert source.name not in repr(request)


async def test_powerpoint_shape_id_reaches_only_the_private_helper_request(
    tmp_path: Path,
) -> None:
    import hashlib
    import zipfile

    source = tmp_path / "current.pptx"
    with zipfile.ZipFile(source, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("ppt/presentation.xml", "<presentation/>")
    digest = hashlib.sha256(source.read_bytes()).hexdigest()
    document = _document(source, application=FocusApplication.POWERPOINT)
    discovery = HelperRun(
        payload={
            "outcome": FocusOutcome.DISCOVERED.value,
            "discovery": _payload(document),
        }
    )
    navigator, seen = _navigator([discovery, discovery, discovery])
    service = _service(tmp_path / "data", navigator=navigator)
    service.acknowledge(CLIENT)
    seed = FocusTargetSeed(
        artifact=FocusArtifact.PPT,
        role=FocusRole.CURRENT_PPT,
        slide_index=2,
        shape_id=77,
    )
    record = _record(
        hashes={"current_ppt": digest},
        seeds={"F0001": (seed,)},
    )
    await service.bind(CLIENT, record, FocusRole.CURRENT_PPT)
    assert service.confirm(CLIENT, record, FocusRole.CURRENT_PPT) is BindOutcome.MATCHED
    claim = ActionClaim(
        run_id=7,
        finding_id="F0001",
        role=FocusRole.CURRENT_PPT,
        revision=0,
        issued_at=dt.datetime.now(dt.UTC),
    )
    await service.focus(CLIENT, record, claim)
    request = next(
        item for item in seen if item.get("action") == FocusAction.FOCUS.value
    )
    assert request["slide_index"] == 2
    assert request["shape_id"] == 77
    assert str(source) not in repr(request)


async def test_focus_refuses_when_the_feature_is_off(tmp_path: Path) -> None:
    service = FocusService(tmp_path, enabled=False, platform="win32")
    claim = ActionClaim(
        run_id=7,
        finding_id="F0001",
        role=FocusRole.CURRENT_EXCEL,
        revision=0,
        issued_at=dt.datetime.now(dt.UTC),
    )
    reply = await service.focus(CLIENT, _record(), claim)
    assert reply.outcome is FocusOutcome.UNSUPPORTED_PLATFORM


async def test_bind_refuses_when_discovery_is_incomplete(
    tmp_path: Path, open_workbook: tuple[Path, str]
) -> None:
    _source, digest = open_workbook
    navigator, _seen = _navigator(
        [
            HelperRun(
                payload={
                    "outcome": FocusOutcome.DISCOVERED.value,
                    "discovery": {
                        "application": "excel",
                        "documents": [],
                        "reasons": ["cross_session_office_window"],
                    },
                }
            )
        ]
    )
    service = _service(tmp_path / "data", navigator=navigator)
    report = await service.bind(
        CLIENT, _record(hashes={"current_excel": digest}), FocusRole.CURRENT_EXCEL
    )
    assert report.outcome is BindOutcome.ENUMERATION_INCOMPLETE


def test_discovery_payload_round_trips(tmp_path: Path) -> None:
    document = _document(tmp_path / "a.xlsx")
    rebuilt = DiscoveryResult.from_payload(_payload(document))
    assert rebuilt is not None
    assert rebuilt.documents[0] == document
    assert DiscoveryResult.from_payload({"application": "word"}) is None
    assert DiscoveryResult.from_payload("nonsense") is None
