"""NiceGUI local web application for the QC tool.

The UI stays thin: uploads, profile selection, passwords, one button.
All comparison logic lives in `qc_tool.engine`; `qc_tool.run_service.perform_run`
wraps a run with report writing and history recording, and heavy runs execute
in one owned worker process managed by `qc_tool.runqueue`. Everything runs on
localhost; sources are read-only. The visual language lives in
`qc_tool.ui.theme`.
"""

import asyncio
import dataclasses
import datetime as dt
import html
import json
import logging
import math
import re
import secrets
import sqlite3
import tempfile
from collections.abc import Callable, Iterable, Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, cast, overload

from fastapi import HTTPException
from nicegui import app, events, ui
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple

from qc_tool.config.lint import lint_profile
from qc_tool.config.profile import (
    CrosscheckMapping,
    DeliverableProfile,
    NumericTolerance,
    list_profiles,
    load_profile,
    load_profile_by_name,
    new_profile,
    profile_path,
    profile_sha256,
    save_profile,
)
from qc_tool.config.promotion import (
    PromotionKind,
    PromotionRequest,
    available_promotions,
    promote_finding,
)
from qc_tool.coverage import (
    CoverageItem,
    CoverageState,
    MappingCoverage,
    QCRunMode,
    capability_limited,
)
from qc_tool.crosscheck.trace import MappingSuggestion, SuggestedSource
from qc_tool.engine import FindingsDelta, QCRunResult, compare_findings
from qc_tool.excel.formulas import formula_token_diff
from qc_tool.findings import (
    Finding,
    FindingClass,
    GridExcerpt,
    SeriesAnchor,
    Severity,
)
from qc_tool.findings_store import finding_by_id
from qc_tool.focus.binding import BindOutcome
from qc_tool.focus.model import FocusTargetSeed
from qc_tool.focus.protocol import FocusOutcome
from qc_tool.focus.service import ROLE_LABELS as FOCUS_ROLE_LABELS
from qc_tool.focus.service import BindReport, FocusService, TokenRejection
from qc_tool.history.carry_forward import apply_carry_forward, preview_carry_forward
from qc_tool.history.run_state import RunStateRecord, RunStatus
from qc_tool.history.store import RunHistory, RunRecord, export_runs_archive, sha256_file
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.io.peek import peek_sheet_names, peek_slide_titles
from qc_tool.launcher import (
    HEALTH_PATH,
    active_instance_health,
    claim_local_instance,
    launcher_log_path,
    release_local_instance,
)
from qc_tool.package import (
    MAX_WORKBOOKS_PER_SIDE,
    MEMBER_ID_PATTERN,
    PackageArtifact,
    PackageManifest,
    PackageSide,
)
from qc_tool.ppt.extract import load_deck_snapshot
from qc_tool.progress import RunPhase
from qc_tool.projection import VolumeProjection, project_cycle_volume
from qc_tool.report.excel_report import write_excel_report
from qc_tool.report.html_report import write_html_report
from qc_tool.review import (
    ReviewAnnotation,
    ReviewGroup,
    apply_group_review,
    build_pattern_groups,
    count_pattern_groups,
    format_group_ranges,
    format_ranges,
    prioritize_review,
)
from qc_tool.review_series import (
    ReviewSlice,
    SeriesCluster,
    SeriesLensError,
    SeriesReviewLens,
    anchor_matches_finding,
    anchor_segment,
    build_series_review_lens,
    cluster_confirmation_updates,
    series_position,
)
from qc_tool.review_stream import (
    GroupPriorityAggregate,
    GroupSummary,
    counts_from_summaries,
    prioritize_review_summaries,
)
from qc_tool.run_preflight import duplicate_byte_conflicts
from qc_tool.run_service import (
    REPORT_DEFER_FINDINGS,
    RunArtifacts,
    perform_run,
)
from qc_tool.runqueue import (
    QueueBusyError,
    RunQueueManager,
    RunRequest,
    get_manager,
    new_request_id,
)
from qc_tool.security import private_directory, private_file, secure_managed_tree
from qc_tool.server_config import (
    NetworkMode,
    lan_config_matches,
    load_server_config,
    local_config,
    save_server_config,
)
from qc_tool.shortcut import (
    ShortcutState,
    install_shortcut,
    remove_shortcut,
    shortcut_status,
)
from qc_tool.signoff import assess_signoff, finalize_run
from qc_tool.story import ChangeStory, build_stories
from qc_tool.triage.preview import (
    POLICY_PREVIEW_MAX_FINDINGS,
    CounterfactualPolicy,
    PolicyPreviewUnavailable,
    PreviewReviewFloor,
    preview_policy,
)
from qc_tool.ui.guide import render_guide
from qc_tool.ui.profile_editor import ProfileEditorController, open_profile_editor
from qc_tool.ui.theme import (
    COL_RESIZE_JS,
    FINDINGS_BODY_SLOT,
    HISTORY_BODY_SLOT,
    REVIEW_GROUPS_BODY_SLOT,
    REVIEW_MEMBER_ROWS_SLOT,
    page_frame,
    section,
    status_chip,
)

logger = logging.getLogger(__name__)


@app.get(HEALTH_PATH, include_in_schema=False)
async def _instance_health(challenge: str) -> dict[str, object]:
    payload = active_instance_health(challenge)
    if payload is None:
        raise HTTPException(status_code=404)
    return payload

__all__ = [
    "RunArtifacts",
    "create_pages",
    "list_profiles",
    "load_profile_by_name",
    "perform_run",
    "persist_confirmed_mapping",
    "run_app",
]

MAX_UPLOAD_BYTES = 500 * 1024 * 1024  # accept large workbooks locally

ROLES = ("baseline_excel", "current_excel", "baseline_ppt", "current_ppt")
ROLE_LABELS = {
    "baseline_excel": "Baseline — Excel workbook",
    "current_excel": "Current — Excel workbook",
    "baseline_ppt": "Baseline — PowerPoint deck",
    "current_ppt": "Current — PowerPoint deck",
}
ROLE_ACCEPT = {
    "baseline_excel": ".xlsx,.xlsm,.xlsb",
    "current_excel": ".xlsx,.xlsm,.xlsb",
    "baseline_ppt": ".pptx",
    "current_ppt": ".pptx",
}
ROLE_SHORT = {
    "baseline_excel": "baseline Excel",
    "current_excel": "current Excel",
    "baseline_ppt": "baseline PPT",
    "current_ppt": "current PPT",
}
#: Inside a baseline/current panel the cycle is already stated by the panel.
ROLE_ARTIFACT = {
    "baseline_excel": "Excel workbook",
    "current_excel": "Excel workbook",
    "baseline_ppt": "PowerPoint deck",
    "current_ppt": "PowerPoint deck",
}
#: Uploads are grouped by cycle so a current file cannot be dropped into a
#: baseline slot just by scanning left to right.
ROLE_GROUPS = (
    (
        "baseline",
        "Baseline",
        "the previous cycle you compare against",
        ("baseline_excel", "baseline_ppt"),
    ),
    (
        "current",
        "Current",
        "the cycle you are signing off",
        ("current_excel", "current_ppt"),
    ),
)
MODE_LABELS = {
    QCRunMode.CURRENT_FILE_PREFLIGHT: "Current-file preflight",
    QCRunMode.CYCLE_COMPARISON: "Cycle comparison",
    QCRunMode.FINAL_PACKAGE: "Final-package QC",
}
PHASE_LABELS = {
    RunPhase.PREPARING: "Preparing run",
    RunPhase.LOADING_BASELINE_EXCEL: "Loading baseline Excel",
    RunPhase.LOADING_CURRENT_EXCEL: "Loading current Excel",
    RunPhase.LOADING_BASELINE_POWERPOINT: "Loading baseline PowerPoint",
    RunPhase.LOADING_CURRENT_POWERPOINT: "Loading current PowerPoint",
    RunPhase.ANALYZING_EXCEL: "Analyzing Excel",
    RunPhase.DIFFING_EXCEL: "Comparing values and structure",
    RunPhase.COMPARING_FORMULAS: "Comparing formulas",
    RunPhase.INDEXING_DEPENDENCIES: "Indexing dependencies",
    RunPhase.QUERYING_IMPACTS: "Querying impacts",
    RunPhase.BUILDING_REVIEW: "Building review groups",
    RunPhase.ANALYZING_POWERPOINT: "Analyzing PowerPoint",
    RunPhase.CROSSCHECKING: "Cross-checking package",
    RunPhase.WRITING_REPORTS: "Writing reports",
    RunPhase.RECORDING_HISTORY: "Recording history",
    RunPhase.COMPLETE: "Complete",
}
QUEUE_STATUS_LABELS = {
    RunStatus.QUEUED: "Queued",
    RunStatus.STARTING: "Starting",
    RunStatus.RUNNING: "Running",
    RunStatus.CANCELLING: "Cancelling",
}


def _download_handler(path: str | Path):
    """Zero-arg click handler (NiceGUI passes event args to 1-arg lambdas)."""

    def handler() -> None:
        ui.download(str(path))

    return handler


#: Rendered client-side just before the server stops, so the dropped websocket
#: reads as an intentional shutdown rather than a lost connection. It overlays
#: the app instead of replacing it, because NiceGUI's own disconnect handler
#: still expects its DOM to exist.
_STOPPED_PAGE_JS = """
document.title = 'QC Tool stopped';
if (!document.querySelector('.stopped-page')) {
  const shell = document.createElement('div');
  shell.className = 'stopped-page';
  shell.innerHTML =
    '<div class="t">QC Tool has stopped</div>' +
    '<div class="d">The local server is no longer running. You can close this tab.</div>' +
    '<div class="d">Start it again from your terminal when you need it.</div>';
  document.body.appendChild(shell);
}
"""


@dataclass(slots=True)
class SessionState:
    files: dict[str, Path] = field(default_factory=dict)
    file_sizes: dict[str, int] = field(default_factory=dict)
    file_hashes: dict[str, str] = field(default_factory=dict)
    upload_generations: dict[str, int] = field(default_factory=dict)
    passwords: dict[str, str] = field(default_factory=dict)  # role -> password
    profile_name: str = "default"
    mode: QCRunMode = QCRunMode.CURRENT_FILE_PREFLIGHT
    allow_large_workbooks: bool = False
    acceptance_absolute: float = 0.0
    acceptance_percent: float = 0.0  # analyst-facing percent; engine gets a fraction
    selected_sheets: set[str] = field(default_factory=set)
    available_sheets: list[str] = field(default_factory=list)
    selected_slides: set[int] = field(default_factory=set)
    available_slides: list[tuple[int, str]] = field(default_factory=list)
    rerun_of: int | None = None
    rerun_required: frozenset[str] = frozenset()  # roles the previous run used
    # Package/manifest aware session state
    package_manifest: PackageManifest | None = None
    selected_member_sheets: dict[str, set[str]] = field(default_factory=dict)
    available_member_sheets: dict[str, list[str]] = field(default_factory=dict)
    member_order: dict[str, list[str]] = field(
        default_factory=lambda: {"baseline": [], "current": []}
    )


def _initial_mode(
    stored: object,
    *,
    rerun_mode: QCRunMode | None = None,
) -> QCRunMode:
    if rerun_mode is not None:
        return rerun_mode
    try:
        return QCRunMode(str(stored))
    except ValueError:
        return QCRunMode.CURRENT_FILE_PREFLIGHT


def _set_desktop_focus_preference(
    work_dir: Path,
    service: FocusService,
    enabled: bool,
) -> None:
    config = load_server_config(work_dir)
    updated = config.model_copy(update={"desktop_focus": enabled})
    if not enabled:
        service.set_enabled(False)
        save_server_config(work_dir, updated)
        return
    save_server_config(work_dir, updated)
    service.set_enabled(True)


def _persist_expired_lan_config(work_dir: Path) -> None:
    current = load_server_config(work_dir)
    save_server_config(
        work_dir,
        local_config(desktop_focus=current.desktop_focus),
    )


# --- pure helpers (unit-testable without a browser) -------------------------


def _safe_upload_name(raw_name: str) -> str:
    """Return a browser-supplied basename that cannot escape managed storage."""
    name = Path(raw_name.replace("\\", "/")).name.strip()
    if not name or name in {".", ".."}:
        raise ValueError("uploaded file has no usable filename")
    return name


def _profile_path(profiles_dir: Path, name: str) -> Path:
    """Resolve a validated profile name inside the managed profile directory."""
    return profile_path(profiles_dir, name)


def _storage_secret(work_dir: Path) -> str:
    """Load or create the private signing secret used by NiceGUI storage."""
    secret_path = work_dir / ".nicegui-storage-secret"
    if secret_path.exists():
        secret = secret_path.read_text(encoding="utf-8").strip()
        if secret:
            return secret
    private_directory(work_dir)
    secret = secrets.token_urlsafe(32)
    secret_path.write_text(secret, encoding="utf-8")
    secret_path.chmod(0o600)
    return secret


def _files_for_mode(
    mode: QCRunMode, files: dict[str, Path]
) -> dict[str, Path]:
    """Validate and select only file roles used by one QC mode."""
    # Helper to match roles by prefix (supports dynamic member-qualified roles)
    def keys_with_prefix(prefix: str) -> list[str]:
        return sorted(
            key for key in files if key == prefix or key.startswith(prefix + ":")
        )

    if mode is QCRunMode.CYCLE_COMPARISON:
        excel_baseline = keys_with_prefix("baseline_excel")
        excel_current = keys_with_prefix("current_excel")
        ppt_baseline = [k for k in files if k == "baseline_ppt"]
        ppt_current = [k for k in files if k == "current_ppt"]
        has_excel = bool(excel_baseline and excel_current)
        has_ppt = bool(ppt_baseline and ppt_current)
        if not has_excel and not has_ppt:
            raise ValueError(
                "Upload a baseline and current Excel pair and/or PowerPoint pair"
            )
        selected: dict[str, Path] = {}
        if has_excel:
            for k in excel_baseline + excel_current:
                selected[k] = files[k]
        if has_ppt:
            for k in ppt_baseline + ppt_current:
                selected[k] = files[k]
        return selected
    if mode is QCRunMode.CURRENT_FILE_PREFLIGHT:
        selected = {
            role: path
            for role, path in files.items()
            if role == "current_excel"
            or role.startswith("current_excel:")
            or role == "current_ppt"
        }
        if not selected:
            raise ValueError("Upload a current Excel workbook and/or PowerPoint deck")
        return selected
    # final package: require at least one current Excel member and exactly one current PPT
    current_exc = sorted(
        key
        for key in files
        if key == "current_excel" or key.startswith("current_excel:")
    )
    current_ppts = [key for key in files if key == "current_ppt"]
    if not current_exc or len(current_ppts) != 1:
        raise ValueError("Upload both current Excel and current PowerPoint files")
    selected = {k: files[k] for k in current_exc}
    # include the single primary current_ppt
    if "current_ppt" in files:
        selected["current_ppt"] = files["current_ppt"]
    return selected


def _role_requirement(mode: QCRunMode, role: str) -> str:
    """Badge text describing how one file role participates in a QC mode."""
    if mode is QCRunMode.FINAL_PACKAGE:
        return "required"
    if mode is QCRunMode.CURRENT_FILE_PREFLIGHT:
        return "Excel and/or PPT"
    return "Excel pair" if "excel" in role else "PPT pair"


def _run_blockers(
    mode: QCRunMode,
    files: dict[str, Path],
    *,
    file_hashes: Mapping[str, str] | None = None,
    rerun_of: int | None = None,
    rerun_required: frozenset[str] = frozenset(),
) -> list[str]:
    """Everything preventing a run, listed before the analyst presses Run QC."""
    blockers: list[str] = []
    missing = rerun_required - files.keys()
    if rerun_of is not None and missing:
        names = ", ".join(role_label(role) for role in sorted(missing))
        blockers.append(
            f"Re-QC of run #{rerun_of} is blocked until these files are "
            f"selected again: {names}"
        )
    try:
        selected = _files_for_mode(mode, files)
    except ValueError as exc:
        blockers.append(str(exc))
        selected = {}
    if file_hashes is not None:
        verifying = sorted(set(selected) - file_hashes.keys())
        if verifying:
            names = ", ".join(role_label(role) for role in verifying)
            blockers.append(f"Verifying selected file bytes: {names}")
        blockers.extend(
            conflict.message(role_label)
            for conflict in duplicate_byte_conflicts(mode, selected, file_hashes)
        )
    return blockers


def role_label(role: str) -> str:
    """Human label for a role key, adding member id when present."""
    prefix, _, member = role.partition(":")
    base = ROLE_LABELS.get(prefix, prefix.replace("_", " "))
    if member and member != "primary":
        return f"{base} · {member}"
    return base


def _scope_summary(state: SessionState) -> str:
    """The comparison scope actually sent to the engine, in one phrase."""
    parts: list[str] = []
    if state.selected_sheets and set(state.available_sheets) != state.selected_sheets:
        parts.append(
            f"{len(state.selected_sheets)}/{len(state.available_sheets)} sheets"
        )
    scoped_members = [
        member
        for member, selected in state.selected_member_sheets.items()
        if selected
        and set(state.available_member_sheets.get(member, ())) != selected
    ]
    if scoped_members:
        parts.append(f"{len(scoped_members)} member-specific sheet scopes")
    slide_indexes = {index for index, _ in state.available_slides}
    if state.selected_slides and slide_indexes != state.selected_slides:
        parts.append(f"{len(state.selected_slides)}/{len(slide_indexes)} slides")
    return "scope: " + (", ".join(parts) if parts else "everything")


def _acceptance_summary(state: SessionState) -> str:
    """The run-level acceptance threshold in one phrase."""
    bounds: list[str] = []
    if state.acceptance_absolute > 0:
        bounds.append(f"±{state.acceptance_absolute:g}")
    if state.acceptance_percent > 0:
        bounds.append(f"±{state.acceptance_percent:g}%")
    return "acceptance: " + (" or ".join(bounds) if bounds else "strict")


def _input_cautions(state: SessionState) -> list[str]:
    """Likely mix-ups that would otherwise produce a meaningless clean result."""
    cautions: list[str] = []
    for baseline, current, artifact in (
        ("baseline_excel", "current_excel", "Excel"),
        ("baseline_ppt", "current_ppt", "PowerPoint"),
    ):
        if baseline not in state.files or current not in state.files:
            continue
        if (
            state.files[baseline].name == state.files[current].name
            and state.file_sizes.get(baseline) == state.file_sizes.get(current)
            and state.file_hashes.get(baseline)
            and state.file_hashes.get(current)
            and state.file_hashes.get(baseline) != state.file_hashes.get(current)
        ):
            cautions.append(
                f"baseline and current {artifact} look like the same file "
                f"({state.files[baseline].name}) — a clean result would prove nothing"
            )
    try:
        manifest = PackageManifest.from_role_files(state.files)
    except ValueError:
        return cautions
    baseline_ids = {
        member.member_id
        for member in manifest.members_for(
            PackageSide.BASELINE,
            PackageArtifact.EXCEL,
        )
    }
    current_ids = {
        member.member_id
        for member in manifest.members_for(
            PackageSide.CURRENT,
            PackageArtifact.EXCEL,
        )
    }
    for member_id in sorted((baseline_ids & current_ids) - {"primary"}):
        baseline = f"baseline_excel:{member_id}"
        current = f"current_excel:{member_id}"
        if (
            state.files[baseline].name == state.files[current].name
            and state.file_sizes.get(baseline) == state.file_sizes.get(current)
            and state.file_hashes.get(baseline)
            and state.file_hashes.get(current)
            and state.file_hashes.get(baseline) != state.file_hashes.get(current)
        ):
            cautions.append(
                f"baseline and current member {member_id!r} look like the same "
                f"file ({state.files[baseline].name})"
            )
    return cautions


def persist_confirmed_mapping(
    profiles_dir: Path,
    profile_name: str,
    suggestion: MappingSuggestion,
    candidate: SuggestedSource,
) -> CrosscheckMapping:
    """Save one analyst-confirmed suggestion into a named profile."""
    if profile_name == "default":
        raise ValueError("create or select a named profile before confirming mappings")
    profile = load_profile_by_name(profiles_dir, profile_name)
    mapping = CrosscheckMapping(
        slide=suggestion.slide,
        line_skeleton=suggestion.line_skeleton,
        figure_index=suggestion.figure_index,
        label=(
            " ".join(
                part for part in (candidate.row_label, candidate.column_label) if part
            )
            or suggestion.line_skeleton
        ),
        source_sheet=candidate.sheet,
        source_cell=candidate.cell,
        source_member=candidate.source_member,
    )
    identity = (mapping.slide, mapping.line_skeleton, mapping.figure_index)
    profile.crosscheck.mappings = [
        existing
        for existing in profile.crosscheck.mappings
        if (existing.slide, existing.line_skeleton, existing.figure_index) != identity
    ]
    profile.crosscheck.mappings.append(mapping)
    save_profile(profile, _profile_path(profiles_dir, profile_name))
    return mapping


def _queue_status_line(record: RunStateRecord) -> str:
    """One privacy-safe line describing an active or queued request."""
    label = QUEUE_STATUS_LABELS.get(record.status, record.status.value.title())
    parts = [f"#{record.request_id[:8]} · {label}"]
    if record.status is RunStatus.QUEUED and record.queue_position:
        parts.append(f"position {record.queue_position}")
    if record.phase:
        try:
            phase = PHASE_LABELS[RunPhase(record.phase)]
        except ValueError:
            phase = record.phase.replace("_", " ")
        counts = f" ({record.processed}/{record.total})" if record.total else ""
        parts.append(f"{phase}{counts}")
    parts.append(f"{record.elapsed_seconds():.0f}s elapsed")
    names = " · ".join(sorted(record.files.values()))
    if names:
        parts.append(names)
    return "  —  ".join(parts)


def _cancel_handler(manager: RunQueueManager, request_id: str):
    """Zero-arg click handler (NiceGUI passes event args to 1-arg lambdas)."""

    def handler() -> None:
        if manager.cancel(request_id):
            ui.notify("Cancellation requested", type="warning")

    return handler


def _capability_summary(coverage: list[CoverageItem]) -> tuple[str, str, str]:
    """Status tone, title, and detail for what the run could and could not check."""
    unavailable = sum(
        1 for item in coverage if item.state is CoverageState.UNAVAILABLE
    )
    degraded = sum(1 for item in coverage if item.state is CoverageState.DEGRADED)
    if unavailable:
        return (
            "limited",
            "Capability limited",
            f"{unavailable} unavailable, {degraded} degraded — a low finding "
            "count is not a clean result",
        )
    if degraded:
        return ("neutral", "Partial capability", f"{degraded} degraded checks")
    return ("ok", "All checks ran", "")


def _mapping_stats(mapping: MappingCoverage) -> tuple[tuple[str, int], ...]:
    """The complete readable/unavailable claim population shown in review."""
    return (
        ("readable", mapping.eligible),
        ("unavailable", mapping.unavailable),
        ("total surfaces", mapping.eligible + mapping.unavailable),
        ("mapped", mapping.mapped),
        ("verified", mapping.verified),
        ("mismatched", mapping.mismatched),
        ("unresolved", mapping.unresolved),
        ("unmapped", mapping.unmapped),
    )


def _lint_profile_for_record(
    profile: DeliverableProfile,
    record: RunRecord,
) -> list:
    workbook = None
    deck = None
    excel_path = record.file_paths.get("current_excel")
    ppt_path = record.file_paths.get("current_ppt")
    if excel_path:
        workbook = load_workbook_snapshot(Path(excel_path))
    if ppt_path:
        deck = load_deck_snapshot(Path(ppt_path))
    return lint_profile(profile, workbook=workbook, deck=deck)


# --- pages -------------------------------------------------------------------


def _finding_row(finding: Finding, *, mutable: bool = True) -> dict[str, object]:
    where = finding.sheet or finding.slide or ""
    if finding.artifact_member != "primary":
        where = f"{finding.artifact_member} · {where}" if where else finding.artifact_member
    return {
        "id": finding.finding_id,
        "severity": (finding.severity or Severity.WARNING).value,
        "class": finding.finding_class.value,
        "where": where,
        "location": finding.location or finding.baseline_location or "",
        "message": finding.message,
        "baseline": finding.baseline_value or "",
        "current": finding.current_value or "",
        "element": finding.element or "",
        "impacts": "; ".join(finding.impacts),
        "artifact": finding.artifact,
        "comment": finding.analyst_comment,
        "overridden": finding.severity_overridden,
        "mutable": mutable,
        "root": finding.root_cause_key,
        "provenance": (
            finding.provenance.value if finding.provenance is not None else ""
        ),
        "subtype": finding.subtype.value if finding.subtype is not None else "",
        "materiality": (
            finding.materiality.value if finding.materiality is not None else ""
        ),
        "temporal_context": (
            finding.temporal_context.value
            if finding.temporal_context is not None
            else ""
        ),
        "expected_reason": (
            finding.expected_reason.value
            if finding.expected_reason is not None
            else ""
        ),
        "evidence_tags": "; ".join(
            sorted(tag.value for tag in finding.evidence_tags)
        ),
        "waiver": (
            f"{finding.waiver_reason} (expires {finding.waiver_expires})"
            if finding.waiver_reason
            else ""
        ),
        "bx": _excerpt_payload(finding.baseline_excerpt),
        "cx": _excerpt_payload(finding.current_excerpt),
    }


def _excerpt_payload(excerpt: GridExcerpt | None) -> dict[str, object] | None:
    """Excerpt for the Vue grid slots, with repr noise tamed for display."""
    if excerpt is None:
        return None
    payload = excerpt.model_dump()
    payload["cells"] = [
        [_display_cell_text(str(value))[0] for value in row] for row in excerpt.cells
    ]
    return payload


def _findings_rows(
    result: QCRunResult, *, mutable: bool = True
) -> list[dict[str, object]]:
    return [_finding_row(finding, mutable=mutable) for finding in result.findings]


#: Above this, the atomic tab pages server-side instead of shipping every row
#: to the browser; it also matches the findings-store page cache (4 blocks of
#: 5,000), below which a lazy sequence hands out stable cached objects.
_ATOMIC_INLINE_THRESHOLD = 20_000
_ATOMIC_PAGE_SIZE = 100

#: Ceiling for the materialized interactive workbench (16 GB target machines:
#: budget + server baseline + desktop Office must coexist). Estimated as
#: sampled per-finding JSON bytes x count x the measured materialization
#: multiplier (948k-finding run: 1.37 GB raw JSON -> ~13.7 GB workbench).
_WORKBENCH_BUDGET_BYTES = 5 * 1024**3
_WORKBENCH_MATERIALIZE_MULTIPLIER = 8
_WORKBENCH_SAMPLE_FINDINGS = 200


def _estimated_workbench_bytes(findings: Sequence[Finding]) -> int:
    """Projected server memory for the fully interactive run view."""
    total = len(findings)
    if total <= _ATOMIC_INLINE_THRESHOLD:
        return 0
    sample = list(findings[:_WORKBENCH_SAMPLE_FINDINGS])
    if not sample:
        return 0
    sampled_bytes = sum(
        len(json.dumps(finding.model_dump(mode="json"), ensure_ascii=False))
        for finding in sample
    )
    per_finding = sampled_bytes // len(sample)
    return per_finding * total * _WORKBENCH_MATERIALIZE_MULTIPLIER


def _workbench_within_budget(findings: Sequence[Finding]) -> bool:
    return _estimated_workbench_bytes(findings) <= _WORKBENCH_BUDGET_BYTES


def _atomic_page_rows(
    result: QCRunResult, page_index: int, *, mutable: bool = True
) -> list[dict[str, object]]:
    """One server-side page of atomic rows for very large runs."""
    total = len(result.findings)
    start = max(0, min(page_index * _ATOMIC_PAGE_SIZE, max(total - 1, 0)))
    end = min(start + _ATOMIC_PAGE_SIZE, total)
    return [
        _finding_row(finding, mutable=mutable)
        for finding in result.findings[start:end]
    ]


class _FindingIndex:
    """Dict-like finding lookup over a lazy sequence, id-position keyed."""

    def __init__(self, findings: Sequence[Finding]) -> None:
        self._findings = findings

    def get(self, finding_id: str) -> Finding | None:
        return finding_by_id(self._findings, finding_id)


class _MemberSequence(Sequence[Finding]):
    """A group's members hydrated per access from the lazy run sequence."""

    def __init__(
        self, finding_ids: tuple[str, ...], findings: Sequence[Finding]
    ) -> None:
        self._ids = finding_ids
        self._findings = findings

    def __len__(self) -> int:
        return len(self._ids)

    @overload
    def __getitem__(self, index: int) -> Finding: ...

    @overload
    def __getitem__(self, index: slice) -> Sequence[Finding]: ...

    def __getitem__(self, index: int | slice) -> Finding | Sequence[Finding]:
        if isinstance(index, slice):
            return tuple(
                hydrated
                for finding_id in self._ids[index]
                if (hydrated := finding_by_id(self._findings, finding_id))
                is not None
            )
        finding = finding_by_id(self._findings, self._ids[index])
        if finding is None:
            raise IndexError(index)
        return finding

    def __iter__(self) -> Iterator[Finding]:
        for finding_id in self._ids:
            finding = finding_by_id(self._findings, finding_id)
            if finding is not None:
                yield finding


class _SummaryGroup:
    """``ReviewGroup`` twin over a stored summary; members hydrate lazily.

    Detail panels, member dialogs, and group decisions touch one group's
    members at a time, so hydration stays bounded by what the analyst
    actually opens instead of the whole run.
    """

    __slots__ = (
        "_findings",
        "_summary",
        "artifact_member",
        "baseline_bounding_range",
        "baseline_mixed",
        "baseline_ranges",
        "bounding_range",
        "element",
        "expected_growth",
        "finding_class",
        "group_id",
        "ranges",
        "severity",
        "sheet",
        "slide",
        "spatial",
    )

    def __init__(self, summary: GroupSummary, findings: Sequence[Finding]) -> None:
        self._summary = summary
        self._findings = findings
        self.group_id = summary.group_id
        self.finding_class = summary.finding_class
        self.severity = summary.severity
        self.sheet = summary.sheet
        self.slide = summary.slide
        self.element = summary.element
        self.expected_growth = summary.expected_growth
        self.ranges = summary.ranges
        self.bounding_range = summary.bounding_range
        self.baseline_ranges = summary.baseline_ranges
        self.baseline_bounding_range = summary.baseline_bounding_range
        self.baseline_mixed = summary.baseline_mixed
        self.spatial = summary.spatial
        self.artifact_member = summary.artifact_member

    @property
    def artifact(self) -> str:
        return self._summary.artifact

    @property
    def member_count(self) -> int:
        return self._summary.member_count

    @property
    def member_finding_ids(self) -> tuple[str, ...]:
        return self._summary.member_finding_ids

    @property
    def members(self) -> Sequence[Finding]:
        return _MemberSequence(self._summary.member_finding_ids, self._findings)


def _summary_where(summary: GroupSummary) -> str:
    where = summary.sheet or summary.slide or ""
    if summary.artifact_member != "primary":
        where = (
            f"{summary.artifact_member} · {where}"
            if where
            else summary.artifact_member
        )
    return where


def _summary_review_row(
    summary: GroupSummary,
    rationale: str,
    annotated_ids: frozenset[str] | set[str],
    singleton_message: str,
) -> dict[str, object]:
    """``_review_row`` twin that never iterates members.

    Pattern groups are class-uniform (the pattern key includes class and
    severity), so cap detection and the reviewable denominator need no
    member reads; reviewed counts come from the persisted annotations.
    """
    capped = summary.finding_class is FindingClass.FINDINGS_CAPPED
    total_reviewable = 0 if capped else summary.member_count
    reviewed_count = (
        0
        if capped
        else sum(
            1
            for finding_id in summary.member_finding_ids
            if finding_id in annotated_ids
        )
    )
    if total_reviewable == 0:
        review_state = "all"
    elif reviewed_count == 0 or reviewed_count < total_reviewable:
        review_state = "needs_review"
    else:
        review_state = "reviewed"
    where = _summary_where(summary)
    return {
        "id": summary.group_id,
        "kind": "row",
        "parent": "",
        "group_id": summary.group_id,
        "severity": summary.severity.value,
        "sevmix": "",
        "class": summary.finding_class.value,
        "where": where,
        "location": format_ranges(summary.ranges, summary.bounding_range),
        "bounds": summary.bounding_range,
        "baseline": "; ".join(summary.baseline_ranges),
        "members": summary.member_count,
        "message": (
            singleton_message
            if summary.member_count == 1
            else f"{summary.member_count:,} findings · one decision"
        ),
        "element": summary.element,
        "why": rationale,
        "reviewed": reviewed_count,
        "reviewable_members": total_reviewable,
        "review_state": review_state,
        "sel": False,
        "expanded": False,
        "children": 0,
        "children_total": 0,
        "child_page": "",
        "hidden_children": 0,
        "cap_degraded": capped,
    }


#: Top-level lens pagination counts parents and fallbacks only; expanded child
#: slices are paged separately so an expansion can never grow the DOM without
#: bound.
_TOP_PAGE_SIZE = 25
_TOP_PAGE_SIZES = (10, 25, 50, 100)
_CHILD_PAGE_SIZE = 10

_SEVERITY_ORDER: tuple[Severity, ...] = (
    Severity.CRITICAL,
    Severity.WARNING,
    Severity.INFO,
    Severity.EXPECTED,
)


def _review_row(
    group: ReviewGroup,
    members: tuple[Finding, ...],
    row_id: str,
    rationale: str,
    *,
    kind: str = "row",
    parent: str = "",
) -> dict[str, object]:
    """One canonical decision row, or one derived slice of it."""
    # count reviewable members (exclude FINDINGS_CAPPED from denominator)
    total_reviewable = sum(
        1
        for member in members
        if member.finding_class is not FindingClass.FINDINGS_CAPPED
    )
    reviewed_count = sum(
        1
        for member in members
        if (member.finding_class is not FindingClass.FINDINGS_CAPPED)
        and (member.severity_overridden or member.analyst_comment)
    )
    if total_reviewable == 0:
        # cap-only groups remain in All (neither needs_review nor reviewed)
        review_state = "all"
    elif reviewed_count == 0 or reviewed_count < total_reviewable:
        review_state = "needs_review"
    else:
        review_state = "reviewed"

    where = group.sheet or group.slide or ""
    if group.artifact_member != "primary":
        where = (
            f"{group.artifact_member} · {where}" if where else group.artifact_member
        )
    member_count = len(members)
    coordinates = {
        coordinate
        for member in members
        if (coordinate := _cell_coordinate(member.location)) is not None
    }
    location = (
        format_group_ranges(group)
        if member_count == group.member_count
        else "; ".join(
            sorted(
                f"{get_column_letter(column)}{row}" for row, column in coordinates
            )[:3]
        )
        or format_group_ranges(group)
    )
    return {
        "id": row_id,
        "kind": kind,
        "parent": parent,
        "group_id": group.group_id,
        "severity": group.severity.value,
        "sevmix": "",
        "class": group.finding_class.value,
        "where": where,
        "location": location,
        "bounds": group.bounding_range,
        "baseline": "; ".join(group.baseline_ranges),
        "members": member_count,
        # single findings speak for themselves; a group states only what the
        # class and # columns do not already say
        "message": (
            members[0].message
            if member_count == 1
            else f"{member_count:,} findings · one decision"
        ),
        "element": group.element,
        "why": rationale,
        # numeric reviewed counter (ignores FINDINGS_CAPPED in denominator)
        "reviewed": reviewed_count,
        "reviewable_members": total_reviewable,
        "review_state": review_state,
        "sel": False,
        "expanded": False,
        "children": 0,
        "children_total": 0,
        "child_page": "",
        "hidden_children": 0,
        "cap_degraded": any(
            member.finding_class is FindingClass.FINDINGS_CAPPED for member in members
        ),
    }


def _review_group_rows(
    groups: list[ReviewGroup],
    rationales: dict[str, str] | None = None,
) -> list[dict[str, object]]:
    reasons = rationales or {}
    return [
        _review_row(
            group,
            group.members,
            group.group_id,
            reasons.get(group.group_id, ""),
        )
        for group in groups
    ]


def _filter_review_rows(
    rows: list[dict[str, object]],
    *,
    severities: set[str],
    finding_classes: set[str],
    review_state: str,
    needle: str = "",
    story_scope: set[str] | None = None,
    sheets: set[str] | None = None,
) -> list[dict[str, object]]:
    """Apply all Review queue filters with deterministic AND semantics."""
    normalized_needle = needle.strip().casefold()
    pinned = story_scope or set()
    filtered = [
        row
        for row in rows
        if row["severity"] in severities
        and row["class"] in finding_classes
        and (sheets is None or str(row["where"]) in sheets)
        # story pins name canonical groups, so every derived slice of a pinned
        # group stays visible
        and (not pinned or row.get("group_id", row["id"]) in pinned)
        and (
            review_state == "all" or row.get("review_state") == review_state
        )
    ]
    if not normalized_needle:
        return filtered
    return [
        row
        for row in filtered
        if normalized_needle
        in (
            f"{row['where']} {row['class']} {row['location']} {row['message']}"
        ).casefold()
    ]


@dataclass(frozen=True, slots=True)
class LensEntry:
    """One top-level lens row plus its surviving and hidden child slices."""

    row: dict[str, object]
    children: tuple[dict[str, object], ...]
    hidden: int


_QUEUE_ORDERS = ("priority", "severity", "location", "findings")


def _sort_lens_entries(
    entries: list[LensEntry], order: str, *, descending: bool = False
) -> list[LensEntry]:
    """Reorder top-level entries only; children always stay under their parent.

    ``priority`` keeps the evidence ordering; ties in every other key fall
    back to it because the sort is stable.
    """
    if order == "severity":
        rank = {severity.value: index for index, severity in enumerate(_SEVERITY_ORDER)}
        return sorted(
            entries,
            key=lambda entry: rank.get(str(entry.row["severity"]), len(rank)),
            reverse=descending,
        )
    if order == "location":
        return sorted(
            entries,
            key=lambda entry: (
                _natural_key(str(entry.row["where"])),
                _natural_key(str(entry.row["location"])),
            ),
            reverse=descending,
        )
    if order == "findings":
        return sorted(
            entries,
            key=lambda entry: int(str(entry.row["members"])),
            reverse=descending,
        )
    return list(reversed(entries)) if descending else list(entries)


def _cluster_row(
    cluster: SeriesCluster,
    children: Sequence[dict[str, object]],
    *,
    hidden: int,
    rationale: str,
) -> dict[str, object]:
    """A collapsed related-series parent, labelled with structural data only."""
    counts = dict.fromkeys(_SEVERITY_ORDER, 0)
    reviewable = 0
    reviewed = 0
    total_members = 0
    for child in children:
        total_members += int(str(child["members"]))
        reviewable += int(str(child["reviewable_members"]))
        reviewed += int(str(child["reviewed"]))
        severity = Severity(str(child["severity"]))
        counts[severity] += int(str(child["members"]))
    dominant = next(
        (severity for severity in _SEVERITY_ORDER if counts[severity]),
        Severity.WARNING,
    )
    if reviewable == 0:
        review_state = "all"
    elif reviewed == 0 or reviewed < reviewable:
        review_state = "needs_review"
    else:
        review_state = "reviewed"
    # The row keeps each column to one fact: WHERE = sheet, LOCATION = axis
    # position, # = member count. The period count stays in the detail label.
    where = cluster.sheet
    if cluster.artifact_member != "primary":
        where = f"{cluster.artifact_member} · {where}"
    position_label = series_position(cluster.key)
    shown = len(children)
    return {
        "id": cluster.cluster_id,
        "kind": "cluster",
        "parent": "",
        "group_id": "",
        "severity": dominant.value,
        "sevmix": " · ".join(
            f"{counts[severity]} {severity.value}"
            for severity in _SEVERITY_ORDER
            if counts[severity]
        ),
        "mix": [
            {"k": severity.value, "n": counts[severity]}
            for severity in _SEVERITY_ORDER
            if counts[severity]
        ],
        "class": "related_series",
        "where": where,
        "location": position_label,
        "bounds": "",
        "baseline": "",
        "members": total_members,
        "message": f"one series · {shown} decision{'s' if shown != 1 else ''}",
        "element": "",
        "why": rationale,
        "reviewed": reviewed,
        "reviewable_members": reviewable,
        "review_state": review_state,
        "sel": False,
        "expanded": False,
        "children": len(children),
        "children_total": len(cluster.slices),
        "child_page": "",
        "hidden_children": hidden,
        "cap_degraded": any(bool(child["cap_degraded"]) for child in children),
    }


def _lens_entries(
    lens: SeriesReviewLens,
    groups_by_id: Mapping[str, ReviewGroup],
    rationales: Mapping[str, str],
    *,
    severities: set[str],
    finding_classes: set[str],
    review_state: str,
    needle: str = "",
    story_scope: set[str] | None = None,
    sheets: set[str] | None = None,
) -> list[LensEntry]:
    """Filter at child level, then keep every parent with a surviving child."""

    def survives(row: dict[str, object]) -> bool:
        return bool(
            _filter_review_rows(
                [row],
                severities=severities,
                finding_classes=finding_classes,
                review_state=review_state,
                needle=needle,
                story_scope=story_scope,
                sheets=sheets,
            )
        )

    entries: list[LensEntry] = []
    for item in lens.rows:
        if isinstance(item, SeriesCluster):
            children = [
                _review_row(
                    groups_by_id[child.group_id],
                    child.members,
                    child.slice_id,
                    rationales.get(child.group_id, ""),
                    kind="child",
                    parent=item.cluster_id,
                )
                for child in item.slices
                if child.group_id in groups_by_id
            ]
            visible = [row for row in children if survives(row)]
            if not visible:
                continue
            entries.append(
                LensEntry(
                    row=_cluster_row(
                        item,
                        visible,
                        hidden=len(children) - len(visible),
                        rationale=rationales.get(item.slices[0].group_id, ""),
                    ),
                    children=tuple(visible),
                    hidden=len(children) - len(visible),
                )
            )
            continue
        group = groups_by_id.get(item.group_id)
        if group is None:
            continue
        row = _review_row(
            group,
            item.members,
            item.slice_id,
            rationales.get(item.group_id, ""),
        )
        if survives(row):
            entries.append(LensEntry(row=row, children=(), hidden=0))
    return entries


def _flatten_lens_rows(
    entries: Sequence[LensEntry],
    *,
    expanded: set[str],
    child_pages: Mapping[str, int],
    page: int = 0,
    page_size: int = _TOP_PAGE_SIZE,
) -> list[dict[str, object]]:
    """Rows for one top-level page; children only for expanded parents."""
    start = max(page, 0) * page_size
    rows: list[dict[str, object]] = []
    for entry in entries[start : start + page_size]:
        row = dict(entry.row)
        identifier = str(row["id"])
        is_expanded = identifier in expanded
        row["expanded"] = is_expanded
        rows.append(row)
        if not entry.children or not is_expanded:
            continue
        total = len(entry.children)
        pages = max(1, -(-total // _CHILD_PAGE_SIZE))
        index = min(max(child_pages.get(identifier, 0), 0), pages - 1)
        window = entry.children[
            index * _CHILD_PAGE_SIZE : index * _CHILD_PAGE_SIZE + _CHILD_PAGE_SIZE
        ]
        row["child_page"] = (
            f"{index * _CHILD_PAGE_SIZE + 1}-"
            f"{index * _CHILD_PAGE_SIZE + len(window)} of {total}"
            if pages > 1
            else ""
        )
        rows.extend(dict(child) for child in window)
    return rows


_FINDINGS_COLUMNS = [
    {"name": "expand", "label": "", "field": "expand"},
    {"name": "id", "label": "ID", "field": "id", "sortable": True, "classes": "mono"},
    {"name": "severity", "label": "Severity", "field": "severity", "sortable": True},
    {"name": "class", "label": "Class", "field": "class", "sortable": True, "classes": "mono"},
    {"name": "where", "label": "Sheet / Slide", "field": "where", "sortable": True},
    {"name": "location", "label": "Location", "field": "location", "classes": "mono"},
    {"name": "message", "label": "Message", "field": "message", "align": "left"},
]

# Quasar client-side sorting is deliberately OFF here: parents and children
# share one flat row array, so a header sort would scatter children away
# from their parents. Ordering is owned by _sort_lens_entries instead.
_REVIEW_GROUP_COLUMNS = [
    {
        "name": "severity",
        "label": "Severity",
        "field": "severity",
        "align": "left",
    },
    {
        "name": "class",
        "label": "Class",
        "field": "class",
        "align": "left",
    },
    {
        "name": "where",
        "label": "Location",
        "field": "where",
        "align": "left",
    },
    {
        "name": "members",
        "label": "#",
        "field": "members",
        "align": "center",
    },
    {"name": "message", "label": "Review item", "field": "message", "align": "left"},
]


_HISTORY_COLUMNS = [
    {"name": "select", "label": "", "field": "sel"},
    {"name": "id", "label": "Run ID", "field": "id", "sortable": True},
    {"name": "when", "label": "Started", "field": "started", "sortable": True},
    {"name": "mode", "label": "Mode", "field": "mode", "sortable": True},
    {"name": "profile", "label": "Profile", "field": "profile", "sortable": True},
    {
        "name": "capability",
        "label": "Capability",
        "field": "capability",
        "sortable": True,
    },
    {
        "name": "decisions",
        "label": "Decisions",
        "field": "critical",
        "sortable": True,
        "align": "left",
    },
    {"name": "size", "label": "Size", "field": "size", "sortable": True},
    {"name": "files", "label": "Files", "field": "files", "align": "left"},
    {"name": "actions", "label": "", "field": "actions"},
]


#: History grows without output caps; past this total the UI suggests cleanup.
HISTORY_STORAGE_PROMPT_BYTES = 1_073_741_824


def _format_bytes(size: int | None) -> str:
    if size is None:
        return "\u2014"
    value = float(size)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:,.0f} {unit}" if unit == "B" else f"{value:,.1f} {unit}"
        value /= 1024
    return f"{value:,.1f} GB"  # pragma: no cover - loop always returns


def _storage_prompt_due(
    total_bytes: int,
    *,
    threshold: int = HISTORY_STORAGE_PROMPT_BYTES,
    dismissed_at_bytes: int = 0,
) -> bool:
    """Prompt at the threshold, then again only after ~10% further growth."""
    if total_bytes < threshold:
        return False
    return total_bytes >= dismissed_at_bytes * 1.1


def _relative_time(moment: dt.datetime) -> str:
    """Short relative age; the exact UTC stamp stays available on hover."""
    seconds = (dt.datetime.now(dt.UTC) - moment).total_seconds()
    if seconds < 90:
        return "just now"
    for limit, size, unit in (
        (5400, 60, "min"),
        (172800, 3600, "h"),
        (2592000, 86400, "d"),
    ):
        if seconds < limit:
            return f"{int(seconds // size)} {unit} ago"
    return moment.date().isoformat()


def _history_row(record: RunRecord) -> dict[str, object]:
    primary, metric = (
        (record.pattern_review_counts, "pattern review items")
        if record.pattern_review_counts
        else (record.review_counts, "spatial review items")
        if record.review_counts
        else (record.counts, "atomic findings")
    )
    return {
        "id": record.run_id,
        "sel": False,
        "archived": record.archived,
        "finalized": record.signoff is not None,
        "started": record.started_at.isoformat(timespec="seconds"),
        "when": _relative_time(record.started_at),
        "mode": MODE_LABELS[record.mode],
        "mode_key": record.mode.value,
        "profile": record.profile,
        "capability": "limited" if capability_limited(record.coverage) else "complete",
        "decisions": [{"k": kind, "n": count} for kind, count in primary.items()],
        "metric": metric,
        "critical": primary.get("critical", 0),
        "atomic": sum(record.counts.values()),
        "size": record.storage_bytes or 0,
        "size_label": _format_bytes(record.storage_bytes),
        "files": " · ".join(record.files.values()),
        "exports": [
            kind
            for kind, path in (
                record.signoff.report_paths
                if record.signoff is not None
                else record.report_paths
            ).items()
            if Path(path).exists()
        ],
    }


def _history_trend_row(record: RunRecord, history: RunHistory) -> dict[str, object]:
    exact_reusable: int | str = 0
    if record.rerun_of is not None:
        # Carry-forward preview hydrates BOTH runs finding-by-finding; on
        # monster re-QCs that took the history page (and the browser) down.
        if sum(record.counts.values()) > _DELTA_BUDGET_FINDINGS:
            exact_reusable = "—"
        else:
            try:
                exact_reusable = len(
                    preview_carry_forward(history, record.run_id).exact
                )
            except (KeyError, ValueError):
                exact_reusable = 0
    review_seconds = history.review_seconds(record.run_id)
    mapping = record.mapping_coverage
    return {
        "run": record.run_id,
        "profile": record.profile,
        "atomics": sum(record.counts.values()),
        "decisions": sum(
            (record.pattern_review_counts or record.review_counts or record.counts).values()
        ),
        "limited": sum(
            item.state is not CoverageState.CHECKED for item in record.coverage
        ),
        "mapped": mapping.mapped if mapping is not None else 0,
        "verified": mapping.verified if mapping is not None else 0,
        "exact_reusable": exact_reusable,
        "carried": history.carried_annotation_count(record.run_id),
        "finalized": "yes" if record.signoff is not None else "no",
        "review_minutes": (
            "unknown" if review_seconds is None else f"{review_seconds / 60:.1f}"
        ),
    }


_MEMBER_COLUMNS = [
    {"name": "id", "label": "ID", "field": "id", "sortable": True},
    {"name": "severity", "label": "Severity", "field": "severity", "sortable": True},
    {"name": "class", "label": "Class", "field": "class", "sortable": True},
    {"name": "location", "label": "Location", "field": "location"},
    {"name": "message", "label": "Finding", "field": "message", "align": "left"},
]


def _outcome_summary(review_items: Mapping[Severity, int]) -> tuple[str, str, str]:
    """Whether this run needs analyst review, and how much."""
    critical = review_items.get(Severity.CRITICAL, 0)
    warning = review_items.get(Severity.WARNING, 0)
    if critical:
        return (
            "attention",
            "Review required",
            f"{critical:,} critical and {warning:,} warning decisions",
        )
    if warning:
        return ("limited", "Review required", f"{warning:,} warning decisions")
    return ("ok", "No blocking findings", "no critical or warning decisions")


def _limitation_consequence(item: CoverageItem) -> str:
    """Plain language for what a degraded or unavailable check costs the run."""
    if item.state is CoverageState.UNAVAILABLE:
        return f"this run cannot conclude anything about {item.label.lower()}"
    return f"{item.label.lower()} was only partially checked"


def _evidence_axes(finding: Finding) -> list[tuple[str, str]]:
    """Typed evidence carried by one finding, in display order, empties dropped."""
    axes = (
        ("class", finding.finding_class.value),
        ("artifact", finding.artifact),
        ("location", finding.location or finding.baseline_location or ""),
        ("baseline", finding.baseline_value or ""),
        ("current", finding.current_value or ""),
        ("element", finding.element or ""),
        ("provenance", finding.provenance.value if finding.provenance else ""),
        ("subtype", finding.subtype.value if finding.subtype else ""),
        ("materiality", finding.materiality.value if finding.materiality else ""),
        (
            "temporal context",
            finding.temporal_context.value if finding.temporal_context else "",
        ),
        (
            "expected reason",
            finding.expected_reason.value if finding.expected_reason else "",
        ),
        ("evidence", "; ".join(sorted(tag.value for tag in finding.evidence_tags))),
        ("impacts", "; ".join(finding.impacts)),
        ("root cause", finding.root_cause_key or ""),
        (
            "waiver",
            f"{finding.waiver_reason} (expires {finding.waiver_expires})"
            if finding.waiver_reason
            else "",
        ),
        ("analyst comment", finding.analyst_comment or ""),
    )
    return [(key, value) for key, value in axes if value]


def _context_grid_html(excerpt: GridExcerpt, label: str) -> str:
    """One neighbourhood grid; every source value is escaped before rendering."""
    esc = html.escape
    cols = "".join(f"<th>{esc(str(col))}</th>" for col in excerpt.cols)
    rows: list[str] = []
    for row_index, row in enumerate(excerpt.cells):
        cells: list[str] = []
        for col_index, value in enumerate(row):
            short, full = _display_cell_text(str(value))
            title = f' title="{esc(full)}"' if full else ""
            hit = row_index == excerpt.hit_row and col_index == excerpt.hit_col
            attribute_class = ' class="hit"' if hit else ""
            cells.append(f"<td{attribute_class}{title}>{esc(short)}</td>")
        rows.append(
            f"<tr><th>{esc(str(excerpt.rows[row_index]))}</th>" + "".join(cells) + "</tr>"
        )
    return (
        f'<div class="ctxblock"><div class="ctxlabel">{esc(label)}</div>'
        f'<table class="ctxgrid"><tbody><tr><th></th>{cols}</tr>'
        + "".join(rows)
        + "</tbody></table></div>"
    )


#: Bounded lens context: panels never grow past this, and there are never more
#: than `_CONTEXT_MAX_PANELS` of them regardless of cluster size.
_CONTEXT_MAX_ROWS = 25
_CONTEXT_MAX_COLS = 15
_CONTEXT_MAX_PANELS = 40


@dataclass(frozen=True, slots=True)
class ContextPanel:
    """One bounded, deterministic window merged from persisted excerpts."""

    rows: tuple[int, ...]
    cols: tuple[str, ...]
    cells: tuple[tuple[str, ...], ...]
    #: (row index, column index) -> severity value of a related finding cell.
    marks: dict[tuple[int, int], str]
    selected: tuple[int, int] | None


@dataclass(frozen=True, slots=True)
class ClusterContext:
    """Merged lens context plus its honest coverage disclosure."""

    panels: tuple[ContextPanel, ...]
    shown: int
    total: int
    merged: bool

    @property
    def complete(self) -> bool:
        return self.shown == self.total

    @property
    def coverage_note(self) -> str:
        return f"{self.shown} of {self.total} related finding cells shown"


_ISO_MIDNIGHT_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})[T ]00:00:00(?:\.0+)?$")


def _display_cell_text(raw: str) -> tuple[str, str | None]:
    """Readable grid text plus the exact stored value when they differ.

    Evidence axes always keep the verbatim value; this only tames repr noise
    (`3428.8569518319678`) and midnight timestamps inside context grids.
    """
    midnight = _ISO_MIDNIGHT_RE.match(raw)
    if midnight:
        return midnight.group(1), raw
    if len(raw) <= 10:
        return raw, None
    try:
        number = float(raw)
    except ValueError:
        return raw, None
    if not math.isfinite(number):
        return raw, None
    short = f"{number:.7g}"
    return (short, raw) if short != raw else (raw, None)


def _facet_filter_label(selected: int, total: int, noun: str) -> str:
    """Constant-height summary for a checkbox-facet button."""
    if total == 0 or selected >= total:
        return f"All {noun}"
    if selected == 0:
        return f"No {noun}"
    return f"{selected} of {total} {noun}"


def _class_filter_label(selected: int, total: int) -> str:
    """Constant-height summary for the class filter button."""
    return _facet_filter_label(selected, total, "classes")


def _where_noun(groups: Sequence[ReviewGroup]) -> str:
    """What the location facet filters, from the rows' own fields — slide
    labels are titles when the deck has them, so strings cannot tell."""
    has_sheets = any(group.sheet for group in groups)
    has_slides = any(group.slide and not group.sheet for group in groups)
    if has_slides and not has_sheets:
        return "slides"
    if has_slides and has_sheets:
        return "sheets/slides"
    return "sheets"


def _natural_key(value: str) -> tuple[tuple[int, int | str], ...]:
    """Sort key that keeps slide 2 ahead of slide 10."""
    return tuple(
        (0, int(part)) if part.isdigit() else (1, part)
        for part in re.split(r"(\d+)", value.casefold())
        if part
    )


def _cell_coordinate(location: str | None) -> tuple[int, int] | None:
    """Single-cell A1 reference as ``(row, column)``; ranges and axes return None."""
    if not location or ":" in location:
        return None
    try:
        return coordinate_to_tuple(location.replace("$", "").upper())
    except ValueError:
        return None


def _excerpt_rectangle(excerpt: GridExcerpt) -> tuple[int, int, int, int] | None:
    if not excerpt.rows or not excerpt.cols:
        return None
    try:
        columns = [column_index_from_string(str(col)) for col in excerpt.cols]
    except ValueError:
        return None
    rows = [int(row) for row in excerpt.rows]
    return (min(rows), max(rows), min(columns), max(columns))


def _rectangles_touch(
    left: tuple[int, int, int, int], right: tuple[int, int, int, int]
) -> bool:
    return (
        left[0] <= right[1] + 1
        and right[0] <= left[1] + 1
        and left[2] <= right[3] + 1
        and right[2] <= left[3] + 1
    )


def _excerpt_values(excerpt: GridExcerpt) -> dict[tuple[int, int], str]:
    """Stored ``(row, column) -> display value`` pairs of one persisted excerpt."""
    values: dict[tuple[int, int], str] = {}
    for row_index, row in enumerate(excerpt.cells):
        if row_index >= len(excerpt.rows):
            continue
        row_number = int(excerpt.rows[row_index])
        for col_index, value in enumerate(row):
            if col_index >= len(excerpt.cols):
                continue
            try:
                column = column_index_from_string(str(excerpt.cols[col_index]))
            except ValueError:
                continue
            values[(row_number, column)] = str(value)
    return values


def build_cluster_context(
    members: Sequence[Finding],
    *,
    side: str,
    selected_finding_id: str = "",
) -> ClusterContext:
    """Merge persisted 7x7 excerpts into bounded panels; never reload a source.

    Coverage is limited to what the run stored, so unstored intervening cells
    are reported as missing rather than drawn as blanks.
    """
    baseline = side == "baseline"
    windows: list[tuple[tuple[int, int, int, int], GridExcerpt]] = []
    values: dict[tuple[int, int], str] = {}
    marks: dict[tuple[int, int], str] = {}
    selected: tuple[int, int] | None = None
    total = 0
    conflict = False

    for finding in members:
        excerpt = finding.baseline_excerpt if baseline else finding.current_excerpt
        raw = finding.baseline_location if baseline else finding.location
        coordinate = _cell_coordinate(raw)
        if coordinate is not None:
            total += 1
            severity = (finding.severity or Severity.WARNING).value
            marks.setdefault(coordinate, severity)
            if finding.finding_id and finding.finding_id == selected_finding_id:
                selected = coordinate
        if excerpt is None:
            continue
        rectangle = _excerpt_rectangle(excerpt)
        if rectangle is None:
            continue
        windows.append((rectangle, excerpt))
        for coordinate, text in _excerpt_values(excerpt).items():
            previous = values.get(coordinate)
            if previous is not None and previous != text:
                conflict = True
            values[coordinate] = text

    if not windows:
        return ClusterContext(panels=(), shown=0, total=total, merged=True)

    sources: list[tuple[tuple[int, int, int, int], dict[tuple[int, int], str]]] = []
    if conflict:
        # Fail closed: one panel per stored excerpt rather than a merged view
        # that would have to pick a winner for a contradicting coordinate.
        for rectangle, excerpt in windows:
            sources.append((rectangle, _excerpt_values(excerpt)))
    else:
        groups: list[list[tuple[int, int, int, int]]] = []
        for rectangle, _excerpt in windows:
            merged_into = [
                group
                for group in groups
                if any(_rectangles_touch(rectangle, item) for item in group)
            ]
            if not merged_into:
                groups.append([rectangle])
                continue
            head = merged_into[0]
            head.append(rectangle)
            for extra in merged_into[1:]:
                head.extend(extra)
                groups.remove(extra)
        for bound in sorted(
            {
                (
                    min(item[0] for item in group),
                    max(item[1] for item in group),
                    min(item[2] for item in group),
                    max(item[3] for item in group),
                )
                for group in groups
            }
        ):
            sources.append((bound, values))

    panels: list[ContextPanel] = []
    covered: set[tuple[int, int]] = set()
    for (min_row, max_row, min_col, max_col), source in sources:
        for row_start in range(min_row, max_row + 1, _CONTEXT_MAX_ROWS):
            row_numbers = tuple(
                range(row_start, min(row_start + _CONTEXT_MAX_ROWS, max_row + 1))
            )
            for col_start in range(min_col, max_col + 1, _CONTEXT_MAX_COLS):
                columns = tuple(
                    range(col_start, min(col_start + _CONTEXT_MAX_COLS, max_col + 1))
                )
                if not any(
                    (row, column) in source
                    for row in row_numbers
                    for column in columns
                ):
                    continue
                if len(panels) >= _CONTEXT_MAX_PANELS:
                    break
                panel_marks: dict[tuple[int, int], str] = {}
                panel_selected: tuple[int, int] | None = None
                cells: list[tuple[str, ...]] = []
                for row_index, row in enumerate(row_numbers):
                    line: list[str] = []
                    for col_index, column in enumerate(columns):
                        line.append(source.get((row, column), ""))
                        if (row, column) not in source:
                            continue
                        severity = marks.get((row, column))
                        if severity is not None:
                            panel_marks[(row_index, col_index)] = severity
                            covered.add((row, column))
                        if selected == (row, column):
                            panel_selected = (row_index, col_index)
                    cells.append(tuple(line))
                panels.append(
                    ContextPanel(
                        rows=row_numbers,
                        cols=tuple(get_column_letter(column) for column in columns),
                        cells=tuple(cells),
                        marks=panel_marks,
                        selected=panel_selected,
                    )
                )

    return ClusterContext(
        panels=tuple(panels),
        shown=len(covered),
        total=total,
        merged=not conflict,
    )


def _cluster_context_html(panel: ContextPanel, label: str) -> str:
    """One bounded lens panel; every stored value is escaped before rendering."""
    esc = html.escape
    cols = "".join(f"<th scope=\"col\">{esc(str(col))}</th>" for col in panel.cols)
    rows: list[str] = []
    for row_index, row in enumerate(panel.cells):
        cells: list[str] = []
        for col_index, value in enumerate(row):
            severity = panel.marks.get((row_index, col_index))
            short, full = _display_cell_text(str(value))
            classes: list[str] = []
            attributes = ""
            if severity is not None:
                classes.append(f"sev-{severity}")
                reference = f"{panel.cols[col_index]}{panel.rows[row_index]}"
                described = f"{severity} finding at {reference}"
                if full:
                    described = f"{described} — exact value {full}"
                attributes = f' title="{esc(described)}" aria-label="{esc(described)}"'
            elif full:
                attributes = f' title="{esc(full)}"'
            if panel.selected == (row_index, col_index):
                classes.append("hit")
            attribute_class = f' class="{" ".join(classes)}"' if classes else ""
            cells.append(f"<td{attribute_class}{attributes}>{esc(short)}</td>")
        rows.append(
            f'<tr><th scope="row">{esc(str(panel.rows[row_index]))}</th>'
            + "".join(cells)
            + "</tr>"
        )
    return (
        f'<div class="ctxblock"><div class="ctxlabel">{esc(label)}</div>'
        f'<table class="ctxgrid" role="table" aria-label="{esc(label)}">'
        f"<tbody><tr><th></th>{cols}</tr>" + "".join(rows) + "</tbody></table></div>"
    )


#: ``_style_key`` field order from the loader; shown as a legend so the
#: pipe-joined style evidence is readable without reading source code.
_STYLE_KEY_FIELDS = (
    "fill pattern",
    "fill color (RGB)",
    "bold",
    "italic",
    "font size",
    "font color (RGB)",
    "border left",
    "border right",
    "border top",
    "border bottom",
)


def _render_style_key_diff(finding: Finding) -> None:
    """Decode pipe-joined style keys into a labelled baseline/current table."""
    baseline = (finding.baseline_value or "").split("|")
    current = (finding.current_value or "").split("|")
    if len(baseline) != len(_STYLE_KEY_FIELDS) or len(current) != len(
        _STYLE_KEY_FIELDS
    ):
        return
    with ui.expansion("Style attributes decoded").props("dense"):
        with ui.element("div").classes("detailgrid"):
            for label, before, after in zip(
                _STYLE_KEY_FIELDS, baseline, current, strict=True
            ):
                changed = before != after
                ui.label(label + (" ●" if changed else "")).classes("dk")
                ui.label(
                    f"{before} → {after}" if changed else before
                ).classes("dv" + (" mono" if changed else ""))
        ui.label(
            "● marks the attributes that changed; None means the attribute "
            "is not set."
        ).classes("hint")


def _render_evidence_body(finding: Finding) -> None:
    """Typed evidence axes and context grids for one finding, in the caller's slot."""
    with ui.element("div").classes("detailgrid"):
        for key, value in _evidence_axes(finding):
            ui.label(key).classes("dk")
            ui.label(value).classes("dv")
    if finding.finding_class is FindingClass.STYLE_CHANGED:
        _render_style_key_diff(finding)
    # For formula logic changes, render a token-level diff expansion when available
    if (
        finding.finding_class is FindingClass.FORMULA_LOGIC_CHANGED
        and finding.baseline_value
        and finding.current_value
    ):
        try:
            segments = formula_token_diff(
                finding.baseline_value,
                finding.current_value,
                finding.baseline_location or "",
                finding.location or "",
            )
        except Exception:
            segments = ()
        if segments:
            with (
                ui.expansion("Formula token diff"),
                ui.element("div").classes("formula-diff").props(
                    'role=group aria-label="Formula token diff"'
                ),
            ):
                rendered = "".join(
                    f'<span class="fdiff-{segment.kind.value}">'
                    f"{html.escape(segment.text)}</span>"
                    for segment in segments
                )
                ui.html(rendered)
    for excerpt, label in (
        (finding.baseline_excerpt, "baseline"),
        (finding.current_excerpt, "current"),
    ):
        if excerpt is not None:
            ui.html(_context_grid_html(excerpt, label))


def _format_review_time(seconds: float | None) -> str:
    """Clock-style review time: 0:00 until an hour, then h:mm:ss."""
    if seconds is None:
        return "0:00"
    whole = int(seconds)
    hours, remainder = divmod(whole, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours}:{minutes:02d}:{secs:02d}"
    return f"{minutes}:{secs:02d}"


def _run_series_anchors(
    result: QCRunResult,
    history: RunHistory | None,
    run_id: int | None,
) -> dict[str, SeriesAnchor]:
    """Validated anchors for the lens; any failure yields the canonical queue."""
    if history is not None and run_id is not None:
        try:
            return history.get_series_anchors(run_id)
        except (KeyError, sqlite3.Error, ValueError):
            logger.warning("series anchor sidecar unavailable for this run")
            return {}
    return {
        finding.finding_id: finding.series_anchor
        for finding in result.findings
        if finding.series_anchor is not None
        and anchor_matches_finding(finding, finding.series_anchor)
    }


#: Ceiling for the re-QC delta banner: comparing two runs walks every
#: finding of both, so monster pairs skip it (the page says so) instead of
#: blocking the server for minutes.
_DELTA_BUDGET_FINDINGS = 50_000


def _rerun_delta(
    history: "RunHistory", record: "RunRecord"
) -> tuple[FindingsDelta | None, str]:
    """Delta against the re-QC baseline, or a note when it is out of budget."""
    if record.rerun_of is None:
        return None, ""
    try:
        previous = history.get_run(record.rerun_of)
    except KeyError:
        return None, ""
    if (
        len(previous.findings) > _DELTA_BUDGET_FINDINGS
        or len(record.findings) > _DELTA_BUDGET_FINDINGS
    ):
        return None, (
            f"change summary vs run #{record.rerun_of} is skipped for very "
            "large runs — counts and the review queue are complete"
        )
    return compare_findings(previous.findings, record.findings), ""


def _on_demand_export_button(
    result: QCRunResult,
    kind: str,
    history: "RunHistory",
    run_id: int,
    export_root: Path,
    current_paths: dict[str, Path],
) -> None:
    """Generate a deferred report in a worker thread, store it, download it."""
    writer = write_excel_report if kind == "excel" else write_html_report
    extension = "xlsx" if kind == "excel" else "html"
    button = ui.button(f"Generate {kind.capitalize()} report").classes(
        "ghostbtn"
    ).props("no-caps flat dense")

    async def generate() -> None:
        button.set_enabled(False)
        ui.notify(
            f"Generating the {kind.capitalize()} report — a very large run "
            "takes several minutes; reviewing can continue meanwhile."
        )
        try:
            private_directory(export_root)
            stamp = dt.datetime.now(dt.UTC).strftime("%Y%m%dT%H%M%S.%fZ-")
            directory = Path(tempfile.mkdtemp(prefix=stamp, dir=export_root))
            private_directory(directory)
            target = directory / f"qc_report.{extension}"
            await asyncio.to_thread(writer, result, target)
            current_paths[kind] = target
            history.set_report_paths(
                run_id, {name: str(path) for name, path in current_paths.items()}
            )
            ui.download(str(target))
            ui.notify(f"{kind.capitalize()} report ready", type="positive")
        except Exception:
            logger.exception("on-demand %s report generation failed", kind)
            ui.notify(
                f"The {kind.capitalize()} report could not be generated; "
                "see the server log.",
                type="negative",
            )
            button.set_enabled(True)

    button.on_click(generate)


def _render_budget_run_view(
    result: QCRunResult,
    report_paths: dict[str, Path],
    heading: str,
    *,
    history: "RunHistory | None" = None,
    run_id: int | None = None,
    export_root: Path | None = None,
) -> None:
    """Bounded summary for runs whose workbench would exceed the memory budget.

    Renders only from cached counts and lazy page slices — never a whole-run
    materialization — so a monster run can be inspected and its reports
    downloaded without risking the UI server on a 16 GB machine.
    """
    with ui.element("div").classes("runheader"):
        ui.label(heading).classes("runtitle")
        ui.label(
            f"{MODE_LABELS[result.mode]} · profile {result.profile_name!r} · "
            + _result_scope_line(result)
        ).classes("runmeta")
        if result.files:
            ui.label(
                " | ".join(f"{role}: {name}" for role, name in result.files.items())
            ).classes("runmeta")
        with ui.element("div").classes("statstrip"):
            for severity, count in result.counts.items():
                with ui.column().classes(f"stat stat-{severity.value}"):
                    ui.label(f"{count:,}").classes("n")
                    ui.label(f"{severity.value} atomic findings").classes("l")
        ui.label(
            f"This run's {len(result.findings):,} findings exceed the "
            "interactive review budget, so the workbench is summarized "
            "here. The Excel, HTML, and JSON exports carry every finding "
            "and every decision surface."
        ).classes("notecard")
        for disclosure in result.disclosures:
            ui.label(disclosure).classes("notecard")
        with ui.row().classes("headeractions"):
            budget_paths = dict(report_paths)
            for kind, label in (("excel", "Download Excel report"),
                                ("html", "Download HTML report")):
                path = budget_paths.get(kind)
                if path is not None and Path(path).exists():

                    def download(path=path) -> None:
                        ui.download(str(path))

                    ui.button(label, on_click=download).classes("ghostbtn").props(
                        "no-caps flat dense"
                    )
                elif (
                    history is not None
                    and run_id is not None
                    and export_root is not None
                ):
                    _on_demand_export_button(
                        result, kind, history, run_id, export_root, budget_paths
                    )

    ui.label("Atomic evidence").classes("paneltitle")
    findings_table = (
        ui.table(
            columns=[c for c in _FINDINGS_COLUMNS if c["name"] != "expand"],
            rows=[],
            row_key="id",
            pagination=_ATOMIC_PAGE_SIZE,
        )
        .classes("findings-table")
        .props("flat dense")
    )
    page = {"index": 0}
    total = len(result.findings)
    pages = max(1, -(-total // _ATOMIC_PAGE_SIZE))

    def refresh() -> None:
        page["index"] = min(max(page["index"], 0), pages - 1)
        findings_table.rows = _atomic_page_rows(
            result, page["index"], mutable=False
        )
        start = page["index"] * _ATOMIC_PAGE_SIZE
        end = min(start + _ATOMIC_PAGE_SIZE, total)
        page_label.set_text(f"Findings {start + 1:,}-{end:,} of {total:,}")
        previous.set_enabled(page["index"] > 0)
        forward.set_enabled(end < total)

    def turn(delta: int) -> None:
        page["index"] += delta
        refresh()

    with ui.element("div").classes("queuepager"):
        previous = ui.button(
            icon="chevron_left", on_click=lambda: turn(-1)
        ).props("flat dense")
        page_label = ui.label("").classes("meta")
        forward = ui.button(
            icon="chevron_right", on_click=lambda: turn(1)
        ).props("flat dense")
    refresh()


def _render_result_view(
    result: QCRunResult,
    report_paths: dict[str, Path],
    heading: str,
    *,
    delta: FindingsDelta | None = None,
    rerun_of: int | None = None,
    history: RunHistory | None = None,
    run_id: int | None = None,
    profiles_dir: Path | None = None,
    focus_actions: Callable[[Finding], None] | None = None,
    export_root: Path | None = None,
) -> None:
    """Results workbench: a compact run header, then Review queue (default),
    Stories, Coverage, Atomic evidence, and Mapping review. Only the active
    view is mounted, so a high-volume run does not pay for hidden tables."""
    summary_view: (
        tuple[
            list[GroupSummary],
            dict[str, GroupPriorityAggregate],
            list[ChangeStory],
        ]
        | None
    ) = None
    if not _workbench_within_budget(result.findings):
        summary_view = (
            history.get_view_summaries(run_id)
            if history is not None and run_id is not None
            else None
        )
        if summary_view is None:
            _render_budget_run_view(
                result,
                report_paths,
                heading,
                history=history,
                run_id=run_id,
                export_root=export_root,
            )
            return
    if summary_view is None and (
        len(result.findings) > _ATOMIC_INLINE_THRESHOLD
        and not isinstance(result.findings, list)
    ):
        # Groups, stories, the id index, and the pager each iterate the
        # findings; one validation pass beats four over a lazy sequence.
        result = dataclasses.replace(result, findings=list(result.findings))
    findings_by_id = _FindingIndex(result.findings)
    signoff = (
        history.get_signoff(run_id)
        if history is not None and run_id is not None
        else None
    )
    mutable = signoff is None
    if signoff is not None:
        report_paths = {
            kind: Path(path) for kind, path in signoff.report_paths.items()
        }
    if summary_view is not None:
        summary_list, summary_aggregates, stories = summary_view
        # ReviewGroup-shaped adapters: queue building never touches members;
        # detail panels and decisions hydrate one group's members at a time.
        review_groups = [
            cast(ReviewGroup, _SummaryGroup(item, result.findings))
            for item in summary_list
        ]
    else:
        summary_list, summary_aggregates = [], {}
        review_groups = build_pattern_groups(result.findings)
        stories = build_stories(result.findings)

    def _annotated_ids() -> frozenset[str]:
        if history is None or run_id is None:
            return frozenset()
        return frozenset(history.get_annotations(run_id))

    series_anchors = _run_series_anchors(result, history, run_id)
    all_rows: list[dict[str, object]] = []  # atomic rows are built on demand
    atomic_paged = len(result.findings) > _ATOMIC_INLINE_THRESHOLD
    atomic_page = {"index": 0}
    atomic_loaded = [False]
    atomic_controls: dict[str, Any] = {}
    #: Live analyst decisions overlaid on server-built pages: a lazy page
    #: rebuild reads stored payloads whose annotations predate this view.
    atomic_overlay: dict[str, dict[str, object]] = {}
    story_scope: set[str] = set()  # group ids pinned by a story, empty = no pin
    expanded_clusters: set[str] = set()  # parents start collapsed
    child_pages: dict[str, int] = {}
    try:
        stored_size = int(app.storage.general.get("review_page_size", _TOP_PAGE_SIZE))
    except (TypeError, ValueError):
        stored_size = _TOP_PAGE_SIZE
    top_page = {
        "index": 0,
        "size": stored_size if stored_size in _TOP_PAGE_SIZES else _TOP_PAGE_SIZE,
    }
    queue_order = {"key": "priority", "descending": False}
    lens_entries: list[LensEntry] = []
    active_lens: list[SeriesReviewLens | None] = [None]

    with ui.element("div").classes("runheader"):
        ui.label(heading).classes("runtitle")
        ui.label(
            f"{MODE_LABELS[result.mode]} · profile {result.profile_name!r} · "
            + _result_scope_line(result)
        ).classes("runmeta")
        if result.files:
            ui.label(
                " | ".join(f"{role}: {name}" for role, name in result.files.items())
            ).classes("runmeta")
        status_box = ui.element("div").classes("statusrow")
        stats_box = ui.element("div").classes("statstrip")
        if history is not None and run_id is not None:
            timer_history = history
            timer_run = run_id
            timer_box = ui.element("div").classes("stat timerkpi")

            def render_timer() -> None:
                active = timer_history.active_review_run() == timer_run
                seconds = timer_history.review_seconds(timer_run)
                timer_box.clear()
                with timer_box:
                    ui.label(_format_review_time(seconds)).classes("n")
                    ui.label("review time").classes("l")
                    if not mutable:
                        ui.label("finalized").classes("a")
                    elif active:

                        def pause_review() -> None:
                            timer_history.pause_review_sessions()
                            render_timer()

                        ui.button("Pause", on_click=pause_review).classes(
                            "timerbtn"
                        ).props("no-caps flat dense aria-label='Pause review timer'")
                    else:

                        def start_review() -> None:
                            timer_history.start_review_session(timer_run)
                            render_timer()

                        ui.button("Start", on_click=start_review).classes(
                            "timerbtn"
                        ).props("no-caps flat dense aria-label='Start review timer'")

            def tick_timer() -> None:
                # repaint only while counting; an idle box stays untouched
                if timer_history.active_review_run() == timer_run:
                    render_timer()

            render_timer()
            ui.timer(1.0, tick_timer)

        def render_stats() -> None:
            counts = (
                counts_from_summaries(summary_list)
                if summary_view is not None
                else count_pattern_groups(review_groups)
            )
            status_box.clear()
            with status_box:
                status_chip(*_outcome_summary(counts.review_items))
                status_chip(*_capability_summary(result.coverage))
            stats_box.clear()
            with stats_box:
                for severity, count in counts.review_items.items():
                    with ui.column().classes(f"stat stat-{severity.value}"):
                        ui.label(str(count)).classes("n")
                        ui.label(f"{severity.value} pattern review items").classes("l")
                        ui.label(
                            f"{counts.atomic_findings[severity]:,} atomic findings"
                        ).classes("a")
                with ui.column().classes("stat"):
                    ui.label(str(result.verified_crosschecks)).classes("n")
                    ui.label("cross-checks ok").classes("l")

        render_stats()
        if delta is not None and rerun_of is not None:
            with ui.element("div").classes("deltaline"):
                ui.html(
                    f"vs run #{rerun_of} — "
                    f'<span class="good">{delta.resolved} resolved</span> · '
                    f'<span class="bad">{delta.new} new</span> · '
                    f"{delta.persisting} persisting"
                )
        for disclosure in result.disclosures:
            ui.label(disclosure).classes("notecard")
        if stories:
            headline = " · ".join(
                f"{story.title[:60]} ({story.member_count})"
                for story in stories[:3]
            )
            if len(stories) > 3:
                headline += f" · +{len(stories) - 3} more"
            ui.label(f"stories — {headline}").classes("storyline")

        with ui.row().classes("headeractions"):
            header_paths = dict(report_paths)
            for kind, label, writer in (
                ("excel", "Export Excel (.xlsx)", write_excel_report),
                ("html", "Export HTML", write_html_report),
            ):
                path = header_paths.get(kind)
                if path is not None:

                    def export(writer=writer, path=path) -> None:
                        # Regenerate from the current (annotated) state before download.
                        writer(result, path)
                        ui.download(str(path))

                    ui.button(label, on_click=export).classes("ghostbtn").props(
                        "no-caps flat dense"
                    )
                elif (
                    history is not None
                    and run_id is not None
                    and export_root is not None
                ):
                    _on_demand_export_button(
                        result, kind, history, run_id, export_root, header_paths
                    )
            if run_id is not None:
                ui.button(
                    "Re-QC this run",
                    on_click=lambda: ui.navigate.to(f"/?rerun={run_id}"),
                ).classes("ghostbtn").props("no-caps flat dense")
            if (
                mutable
                and history is not None
                and run_id is not None
                and rerun_of is not None
            ):

                def open_carry_forward() -> None:
                    if len(result.findings) > _DELTA_BUDGET_FINDINGS:
                        ui.notify(
                            "Carry-forward preview walks both runs finding by "
                            "finding and is unavailable for very large runs.",
                            type="warning",
                        )
                        return
                    preview = preview_carry_forward(history, run_id)
                    selected: set[str] = set()
                    with ui.dialog() as dialog, ui.card().classes(
                        "w-[46rem] max-w-full"
                    ):
                        ui.label("Review prior decisions").classes("runhead")
                        ui.label(
                            f"{len(preview.exact)} exact reusable · "
                            f"{len(preview.changed_evidence)} changed evidence · "
                            f"{len(preview.resolved)} resolved · "
                            f"{len(preview.ambiguous)} ambiguous"
                        ).classes("note")
                        checkboxes: list[ui.checkbox] = []
                        if not preview.source_finalized:
                            ui.label(
                                "The source run was not finalized. Select each exact "
                                "decision explicitly before applying it."
                            ).classes("notecard")
                        elif preview.exact:

                            def select_all() -> None:
                                selected.update(
                                    candidate.finding_id for candidate in preview.exact
                                )
                                for checkbox in checkboxes:
                                    checkbox.value = True
                                    checkbox.update()

                            ui.button(
                                "Select all exact", on_click=select_all
                            ).classes("ghostbtn").props("flat no-caps dense")
                        for candidate in preview.exact:

                            def select_candidate(
                                event: events.ValueChangeEventArguments,
                                finding_id: str = candidate.finding_id,
                            ) -> None:
                                if event.value:
                                    selected.add(finding_id)
                                else:
                                    selected.discard(finding_id)

                            checkbox = ui.checkbox(
                                f"{candidate.finding_id} · "
                                f"{candidate.severity or 'note only'}",
                                on_change=select_candidate,
                            )
                            checkboxes.append(checkbox)
                        if not preview.exact:
                            ui.label("No prior decisions have identical evidence.").classes(
                                "lede"
                            )

                        def apply_selected() -> None:
                            if not selected:
                                ui.notify("Select at least one exact decision", type="warning")
                                return
                            try:
                                count = apply_carry_forward(history, run_id, selected)
                            except Exception as exc:
                                ui.notify(str(exc), type="negative")
                                return
                            dialog.close()
                            ui.notify(f"Applied {count} prior decisions")
                            ui.navigate.reload()

                        with ui.row().classes("items-center gap-2"):
                            ui.button(
                                "Apply selected", on_click=apply_selected
                            ).classes("runbtn").props("no-caps")
                            ui.button("Cancel", on_click=dialog.close).props(
                                "flat no-caps"
                            )
                    dialog.open()

                ui.button(
                    "Review prior decisions", on_click=open_carry_forward
                ).classes("ghostbtn").props("no-caps flat dense")
            if signoff is not None:
                ui.label(
                    f"Finalized {signoff.finalized_at}"
                ).classes("reviewed")
                if Path(signoff.attestation_path).is_file():
                    ui.button(
                        "Download attestation",
                        on_click=lambda: ui.download(signoff.attestation_path),
                    ).classes("ghostbtn").props("no-caps flat dense")
            elif history is not None and run_id is not None and profiles_dir is not None:

                async def open_signoff_dialog() -> None:
                    # Assessment walks every finding; threaded so a monster
                    # run cannot block the event loop ("trying to reconnect").
                    big_run = len(result.findings) > _DELTA_BUDGET_FINDINGS
                    if big_run:
                        ui.notify(
                            "Assessing review state — a very large run takes "
                            "a few minutes…"
                        )
                    record = await asyncio.to_thread(history.get_run, run_id)
                    initial = await asyncio.to_thread(assess_signoff, record)
                    accepted: set[str] = set()
                    with ui.dialog() as dialog, ui.card().classes(
                        "w-[42rem] max-w-full"
                    ):
                        ui.label("Finalize reviewed run").classes("runhead")
                        ui.label(
                            "Finalization locks this run's review state, regenerates "
                            "reports, and creates a signed private attestation. Use "
                            "Re-QC for later corrections."
                        ).classes("note")
                        if big_run:
                            ui.label(
                                "This run is very large: finalization writes "
                                "full reports and an attestation and can take "
                                "tens of minutes. It runs in the background — "
                                "keep the server running."
                            ).classes("notecard")
                        if record.profile_snapshot is None:
                            ui.label(
                                "This legacy run has no exact profile snapshot. "
                                "Submit a Re-QC run before sign-off."
                            ).classes("notecard")
                        if initial.undecided_finding_ids:
                            ui.label(
                                f"{len(initial.undecided_finding_ids)} Critical/Warning "
                                "findings still need an analyst decision."
                            ).classes("notecard")
                        for code in initial.required_acknowledgements:

                            def acknowledge(
                                event: events.ValueChangeEventArguments,
                                code: str = code,
                            ) -> None:
                                if event.value:
                                    accepted.add(code)
                                else:
                                    accepted.discard(code)

                            ui.checkbox(
                                f"Acknowledge {code.replace(':', ': ')}",
                                on_change=acknowledge,
                            )

                        async def finalize() -> None:
                            ui.notify(
                                "Finalizing — review state, reports, and "
                                "attestation are being written…"
                            )
                            current = await asyncio.to_thread(
                                history.get_run, run_id
                            )
                            assessment = await asyncio.to_thread(
                                assess_signoff, current, accepted
                            )
                            if not assessment.ready:
                                ui.notify(
                                    "Resolve remaining decisions and acknowledgements first",
                                    type="warning",
                                )
                                return
                            try:
                                await asyncio.to_thread(
                                    finalize_run,
                                    profiles_dir.parent,
                                    run_id,
                                    accepted,
                                )
                            except Exception as exc:
                                ui.notify(str(exc), type="negative")
                                return
                            dialog.close()
                            ui.notify("Run finalized and attestation verified")
                            ui.navigate.reload()

                        with ui.row().classes("items-center gap-2"):
                            finalize_button = ui.button(
                                "Finalize and attest", on_click=finalize
                            ).classes("runbtn").props("no-caps")
                            if (
                                record.profile_snapshot is None
                                or initial.undecided_finding_ids
                            ):
                                finalize_button.disable()
                            ui.button("Cancel", on_click=dialog.close).props(
                                "flat no-caps"
                            )
                    dialog.open()

                ui.button(
                    "Finalize review", on_click=open_signoff_dialog
                ).classes("ghostbtn").props("no-caps flat dense")
            ui.label(
                "exports are for sharing — every finding stays viewable here"
            ).classes("hint")

    with ui.tabs().classes("resulttabs").props("dense no-caps align=left") as tab_bar:
        ui.tab("review", label="Review queue")
        ui.tab("stories", label=f"Stories ({len(stories)})")
        ui.tab("coverage", label="Coverage")
        ui.tab("atomic", label="Atomic evidence")
        if result.mapping_coverage is not None:
            ui.tab("mapping", label="Mapping review")

    with ui.tab_panels(tab_bar, value="review").classes("resultpanels") as panels:
        with ui.tab_panel("review"):
            # Filter state lives in plain sets; the controls are constant-height
            # so the queue is visible on first paint at any viewport.
            selected_severities: set[str] = {
                s.value for s in Severity if s is not Severity.EXPECTED
            }
            class_options = sorted({g.finding_class.value for g in review_groups})
            selected_classes: set[str] = set(class_options)
            # location facet: every sheet/slide this run's rows can name,
            # including the blank whole-file bucket, stable across filters
            where_options = sorted(
                (
                    {_summary_where(summary) for summary in summary_list}
                    if summary_view is not None
                    else {
                        str(row["where"])
                        for row in _review_group_rows(review_groups)
                    }
                ),
                key=_natural_key,
            )
            selected_wheres: set[str] = set(where_options)
            where_noun = _where_noun(review_groups)
            severity_pills: dict[str, ui.button] = {}

            def _decision_counts() -> dict[str, int]:
                counts = count_pattern_groups(review_groups).review_items
                return {severity.value: counts.get(severity, 0) for severity in Severity}

            def _refresh_severity_pills() -> None:
                counts = _decision_counts()
                for value, pill in severity_pills.items():
                    pill.set_text(f"{value} {counts.get(value, 0)}")
                    active = value in selected_severities
                    pill.classes(
                        add="active" if active else "",
                        remove="" if active else "active",
                    )
                    pill.props(f'aria-pressed="{str(active).lower()}"')

            def _toggle_severity(value: str) -> None:
                if value in selected_severities:
                    selected_severities.discard(value)
                else:
                    selected_severities.add(value)
                _refresh_severity_pills()
                refilter()

            def _severity_handler(value: str) -> Callable[[], None]:
                # zero-arg closure: NiceGUI passes the click event to 1-arg
                # lambdas, so a default-arg lambda would receive the event
                return lambda: _toggle_severity(value)

            with ui.row().classes("reviewtoolbar"):
                with ui.row().classes("sevpills").props(
                    'role=group aria-label="Severity filter"'
                ):
                    for severity in Severity:
                        value = severity.value
                        severity_pills[value] = (
                            ui.button(
                                value,
                                on_click=_severity_handler(value),
                            )
                            .classes(f"sevpill sevpill-{value}")
                            .props("no-caps flat dense")
                        )
                class_button = (
                    ui.button(_class_filter_label(len(selected_classes), len(class_options)))
                    .classes("classbtn")
                    .props("no-caps flat dense icon-right=arrow_drop_down")
                )
                with class_button, ui.menu().classes("classmenu").props(
                    "no-parent-event=false"
                ):
                    class_search = (
                        ui.input(placeholder="filter classes")
                        .classes("classsearch")
                        .props("outlined dense clearable")
                    )
                    class_checks: dict[str, ui.checkbox] = {}

                    def _refresh_class_button() -> None:
                        class_button.set_text(
                            _class_filter_label(
                                len(selected_classes), len(class_options)
                            )
                        )

                    def _set_class(value: str, checked: bool) -> None:
                        if checked:
                            selected_classes.add(value)
                        else:
                            selected_classes.discard(value)
                        _refresh_class_button()
                        refilter()

                    def _all_classes(select: bool) -> None:
                        selected_classes.clear()
                        if select:
                            selected_classes.update(class_options)
                        for value, box in class_checks.items():
                            box.set_value(value in selected_classes)
                        _refresh_class_button()
                        refilter()

                    with ui.row().classes("classquick"):
                        ui.button(
                            "All", on_click=lambda: _all_classes(True)
                        ).props("flat dense no-caps")
                        ui.button(
                            "None", on_click=lambda: _all_classes(False)
                        ).props("flat dense no-caps")
                    with ui.column().classes("classlist"):
                        for value in class_options:
                            class_checks[value] = ui.checkbox(
                                value.replace("_", " "),
                                value=True,
                                on_change=lambda e, value=value: _set_class(
                                    value, bool(e.value)
                                ),
                            ).props("dense")

                    def _filter_class_list(
                        e: events.ValueChangeEventArguments,
                    ) -> None:
                        needle = str(e.value or "").strip().casefold()
                        for value, box in class_checks.items():
                            box.set_visibility(
                                not needle or needle in value.replace("_", " ")
                            )

                    class_search.on_value_change(_filter_class_list)
                where_button = (
                    ui.button(
                        _facet_filter_label(
                            len(selected_wheres), len(where_options), where_noun
                        )
                    )
                    .classes("classbtn")
                    .props("no-caps flat dense icon-right=arrow_drop_down")
                )
                with where_button, ui.menu().classes("classmenu").props(
                    "no-parent-event=false"
                ):
                    where_search = (
                        ui.input(placeholder=f"filter {where_noun}")
                        .classes("classsearch")
                        .props("outlined dense clearable")
                    )
                    where_checks: dict[str, ui.checkbox] = {}

                    def _refresh_where_button() -> None:
                        where_button.set_text(
                            _facet_filter_label(
                                len(selected_wheres), len(where_options), where_noun
                            )
                        )

                    def _set_where(value: str, checked: bool) -> None:
                        if checked:
                            selected_wheres.add(value)
                        else:
                            selected_wheres.discard(value)
                        _refresh_where_button()
                        refilter()

                    def _all_wheres(select: bool) -> None:
                        selected_wheres.clear()
                        if select:
                            selected_wheres.update(where_options)
                        for value, box in where_checks.items():
                            box.set_value(value in selected_wheres)
                        _refresh_where_button()
                        refilter()

                    with ui.row().classes("classquick"):
                        ui.button(
                            "All", on_click=lambda: _all_wheres(True)
                        ).props("flat dense no-caps")
                        ui.button(
                            "None", on_click=lambda: _all_wheres(False)
                        ).props("flat dense no-caps")
                    with ui.column().classes("classlist"):
                        for value in where_options:
                            where_checks[value] = ui.checkbox(
                                value or "whole file",
                                value=True,
                                on_change=lambda e, value=value: _set_where(
                                    value, bool(e.value)
                                ),
                            ).props("dense")

                    def _filter_where_list(
                        e: events.ValueChangeEventArguments,
                    ) -> None:
                        needle = str(e.value or "").strip().casefold()
                        for value, box in where_checks.items():
                            box.set_visibility(
                                not needle
                                or needle in (value or "whole file").casefold()
                            )

                    where_search.on_value_change(_filter_where_list)
                # Review state segmented control: All / Needs review / Reviewed
                review_state = ui.toggle(
                    {
                        "all": "All",
                        "needs_review": "Needs review",
                        "reviewed": "Reviewed",
                    },
                    value="all",
                ).classes("review-toggle").props("no-caps")
                text_filter = (
                    ui.input(placeholder="filter by sheet, class, or range")
                    .classes("flex-1 min-w-48")
                    .props("outlined dense clearable")
                )
                story_pin = (
                    ui.button("Clear story filter")
                    .classes("ghostbtn")
                    .props("no-caps flat dense")
                )
                story_pin.visible = False
                # What-if preview for private counterfactual bases
                if (
                    result.mode is QCRunMode.CYCLE_COMPARISON
                    and history is not None
                    and run_id is not None
                ):

                    def open_preview_dialog() -> None:
                        with ui.dialog() as dialog, ui.card().classes(
                            "w-[54rem] max-w-full"
                        ):
                            ui.label("What-if preview").classes("runhead")
                            ui.label(
                                "A pure engine-policy view over typed evidence. "
                                "Analyst overrides are ignored; nothing is saved."
                            ).classes("note")
                            ui.label(
                                "Counts show only decisions that would CHANGE: "
                                "tiny differences the run's materiality policy "
                                "already accepted sit in the Info rows below and "
                                "do not move again."
                            ).classes("note")
                            with ui.row().classes("gap-2 w-full"):
                                absolute_input = ui.number(
                                    "Absolute tolerance", value=0, min=0
                                ).classes("flex-1").props("outlined dense")
                                relative_input = ui.number(
                                    "Relative tolerance (%)", value=0, min=0
                                ).classes("flex-1").props("outlined dense")
                                floor_select = ui.select(
                                    {
                                        floor.value: floor.value.replace("_", " ")
                                        for floor in PreviewReviewFloor
                                    },
                                    value=PreviewReviewFloor.NOISE.value,
                                    label="Review floor",
                                ).classes("flex-1").props("outlined dense")

                            result_box = ui.element("div").classes("w-full")

                            def clear_preview() -> None:
                                absolute_input.value = 0
                                relative_input.value = 0
                                floor_select.value = PreviewReviewFloor.NOISE.value
                                absolute_input.update()
                                relative_input.update()
                                floor_select.update()
                                result_box.clear()

                            async def do_preview() -> None:
                                result_box.clear()
                                if (
                                    len(result.findings)
                                    > POLICY_PREVIEW_MAX_FINDINGS
                                ):
                                    with result_box:
                                        ui.label(
                                            "What-if preview is unavailable for "
                                            "this very large run because rebuilding "
                                            "two complete policy views would exceed "
                                            "the UI memory budget. The recorded run "
                                            "is unchanged."
                                        ).classes("notecard")
                                    return
                                try:
                                    policy = CounterfactualPolicy(
                                        acceptance=NumericTolerance(
                                            absolute=float(
                                                absolute_input.value or 0.0
                                            ),
                                            relative=float(
                                                relative_input.value or 0.0
                                            )
                                            / 100.0,
                                        ),
                                        review_floor=PreviewReviewFloor(
                                            str(floor_select.value)
                                        ),
                                    )
                                except (TypeError, ValueError) as exc:
                                    ui.notify(str(exc), type="warning")
                                    return
                                try:
                                    record = await asyncio.to_thread(
                                        history.get_raw_run, run_id
                                    )
                                    bases = await asyncio.to_thread(
                                        history.get_counterfactual_bases, run_id
                                    )
                                except KeyError:
                                    ui.notify("no such run", type="negative")
                                    return
                                except ValueError:
                                    with result_box:
                                        ui.label(
                                            "Preview unavailable because the private "
                                            "typed evidence is invalid; use Re-QC."
                                        ).classes("notecard")
                                    return
                                if not bases:
                                    with result_box:
                                        if not record.counterfactual_digest:
                                            ui.label(
                                                "Preview unavailable for this legacy "
                                                "run; use Re-QC."
                                            ).classes("notecard")
                                        else:
                                            ui.label(
                                                "This run has no numeric replacement "
                                                "findings eligible for preview."
                                            ).classes("notecard")
                                    return
                                try:
                                    preview = await asyncio.to_thread(
                                        preview_policy,
                                        record.findings,
                                        bases,
                                        record.profile_snapshot,
                                        policy,
                                    )
                                except PolicyPreviewUnavailable as exc:
                                    with result_box:
                                        ui.label(str(exc)).classes("notecard")
                                    return
                                except ValueError as exc:
                                    ui.notify(str(exc), type="negative")
                                    return
                                with result_box:
                                    ui.label(
                                        f"{preview.affected_atomics:,} atomics in "
                                        f"{len(preview.affected_decisions):,} decisions "
                                        "would change classification."
                                    ).classes("lede")
                                    if preview.accepted_atomics:
                                        ui.label(
                                            f"{preview.accepted_atomics:,} accepted "
                                            "atomics would remain visible as Info and "
                                            "move below the action queue."
                                        ).classes("note")
                                    if preview.review_floor_atomics:
                                        ui.label(
                                            f"{preview.review_floor_atomics:,} atomics "
                                            "would move below the action queue under "
                                            "the selected materiality floor."
                                        ).classes("note")
                                    summary_rows = [
                                        {
                                            "severity": severity.value,
                                            "atomic_before": preview.before_atomic[
                                                severity.value
                                            ],
                                            "atomic_after": preview.after_atomic[
                                                severity.value
                                            ],
                                            "decision_before": preview.before_decision[
                                                severity.value
                                            ],
                                            "decision_after": preview.after_decision[
                                                severity.value
                                            ],
                                        }
                                        for severity in Severity
                                    ]
                                    ui.table(
                                        columns=[
                                            {
                                                "name": "severity",
                                                "label": "Severity",
                                                "field": "severity",
                                            },
                                            {
                                                "name": "atomic_before",
                                                "label": "Atomics before",
                                                "field": "atomic_before",
                                            },
                                            {
                                                "name": "atomic_after",
                                                "label": "Atomics after",
                                                "field": "atomic_after",
                                            },
                                            {
                                                "name": "decision_before",
                                                "label": "Decisions before",
                                                "field": "decision_before",
                                            },
                                            {
                                                "name": "decision_after",
                                                "label": "Decisions after",
                                                "field": "decision_after",
                                            },
                                        ],
                                        rows=summary_rows,
                                        row_key="severity",
                                    ).classes("coverage-table").props(
                                        "flat dense hide-bottom"
                                    )
                                    if preview.affected_decisions:
                                        affected_rows = [
                                            {
                                                "group_id": item.before_group_id,
                                                "class": item.finding_class,
                                                "where": item.where,
                                                "location": item.location,
                                                "members": item.affected_members,
                                                "before": item.before_severity,
                                                "after": ", ".join(
                                                    item.after_severities
                                                ),
                                                "reason": ", ".join(item.reasons),
                                            }
                                            for item in preview.affected_decisions[:200]
                                        ]
                                        with ui.element("div").classes(
                                            "ctxblock w-full"
                                        ):
                                            ui.table(
                                                columns=[
                                                    {
                                                        "name": "class",
                                                        "label": "Class",
                                                        "field": "class",
                                                    },
                                                    {
                                                        "name": "where",
                                                        "label": "Sheet",
                                                        "field": "where",
                                                    },
                                                    {
                                                        "name": "location",
                                                        "label": "Location",
                                                        "field": "location",
                                                    },
                                                    {
                                                        "name": "members",
                                                        "label": "#",
                                                        "field": "members",
                                                    },
                                                    {
                                                        "name": "before",
                                                        "label": "Before",
                                                        "field": "before",
                                                    },
                                                    {
                                                        "name": "after",
                                                        "label": "After",
                                                        "field": "after",
                                                    },
                                                    {
                                                        "name": "reason",
                                                        "label": "Reason",
                                                        "field": "reason",
                                                    },
                                                ],
                                                rows=affected_rows,
                                                row_key="group_id",
                                                pagination=25,
                                            ).classes("coverage-table").props(
                                                "flat dense"
                                            )

                            with ui.row().classes("items-center gap-2"):
                                ui.button(
                                    "Preview", on_click=do_preview
                                ).classes("ghostbtn").props("no-caps flat dense")
                                ui.button(
                                    "Reset", on_click=clear_preview
                                ).props("flat no-caps")
                                ui.button("Close", on_click=dialog.close).props(
                                    "flat no-caps"
                                )
                        dialog.open()

                    ui.button(
                        "What-if preview", on_click=open_preview_dialog
                    ).classes("ghostbtn").props("no-caps flat dense")
            with ui.element("div").classes("reviewsplit"):
                with ui.element("div").classes("reviewqueue"):
                    if summary_view is not None:
                        ui.label(
                            "Very large run: the queue reads the recorded "
                            "summaries — decisions update reviewed counters "
                            "and ordering but do not regroup rows until "
                            "re-QC, and series clustering is disabled."
                        ).classes("hint")
                    group_table = (
                        ui.table(
                            columns=_REVIEW_GROUP_COLUMNS,
                            rows=[],
                            row_key="id",
                            pagination=0,
                        )
                        .classes("findings-table review-groups-table")
                        .props("flat dense hide-bottom")
                    )
                    group_table.add_slot("body", REVIEW_GROUPS_BODY_SLOT)
                    ui.add_body_html(COL_RESIZE_JS)
                    # the pager belongs to the list it turns, like the members
                    # dialog, not to the filter toolbar
                    with ui.row().classes("items-center queuepager") as top_pager:
                        top_previous = ui.button(
                            icon="chevron_left",
                            on_click=lambda: _turn_top_page(-1),
                        ).props("flat round dense aria-label='Previous review page'")
                        top_page_label = ui.label().classes("hint")
                        top_next = ui.button(
                            icon="chevron_right",
                            on_click=lambda: _turn_top_page(1),
                        ).props("flat round dense aria-label='Next review page'")
                        ui.select(
                            {size: f"{size} per page" for size in _TOP_PAGE_SIZES},
                            value=top_page["size"],
                            # late-bound: _set_page_size is defined below
                            on_change=lambda e: _set_page_size(e),
                        ).classes("pagesize").props(
                            "dense options-dense borderless "
                            "aria-label='Review rows per page'"
                        )
                        # ordering is Python-owned so a sort can never split a
                        # series from its child decisions (header sort did)
                        ui.select(
                            {
                                "priority": "order: priority",
                                "severity": "order: severity",
                                "location": "order: location",
                                "findings": "order: findings",
                            },
                            value="priority",
                            on_change=lambda e: _set_queue_order(e),
                        ).classes("pagesize").props(
                            "dense options-dense borderless "
                            "aria-label='Review queue order'"
                        )
                        direction_button = ui.button(
                            icon="south",
                            on_click=lambda: _flip_queue_direction(),
                        ).props(
                            "flat round dense aria-label='Reverse queue order'"
                        )
                    top_pager.set_visibility(False)
                detail_box = ui.element("div").classes("detailpanel")

        with ui.tab_panel("stories"):
            if not stories:
                ui.label(
                    "No cross-class change stories were proved for this run."
                ).classes("lede")
            else:
                ui.label(
                    "Narrative synthesis across finding classes — a lens over the "
                    "same findings; severities and counts are unchanged."
                ).classes("note")
            for story in stories:
                mix = " · ".join(
                    f"{name}: {count}"
                    for name, count in story.severity_counts.items()
                    if count
                )
                with ui.expansion(
                    f"{story.title} — {story.member_count} finding"
                    f"{'s' if story.member_count != 1 else ''} ({mix})"
                ).classes("story-item").props("dense"):
                    ui.label("proved evidence").classes("dk")
                    for line in story.evidence:
                        ui.label(line).classes("mono")
                    ui.label("interpretation — not proved by the evidence above").classes(
                        "dk"
                    )
                    ui.label(story.description).classes("note")
                    ui.label(
                        f"story {story.story_id} · kind {story.kind.value}"
                    ).classes("a")

                    def pin_story(story=story) -> None:
                        members = set(story.finding_ids)
                        story_scope.clear()
                        if summary_view is not None:
                            story_scope.update(
                                summary.group_id
                                for summary in summary_list
                                if any(
                                    finding_id in members
                                    for finding_id in summary.member_finding_ids
                                )
                            )
                        else:
                            story_scope.update(
                                group.group_id
                                for group in review_groups
                                if any(
                                    member.finding_id in members
                                    for member in group.members
                                )
                            )
                        story_pin.visible = True
                        panels.set_value("review")
                        refilter()

                    ui.button(
                        "Show member review groups", on_click=pin_story
                    ).classes("ghostbtn").props("no-caps flat dense")

        with ui.tab_panel("coverage"):
            section("Check coverage")
            ui.link("Coverage guide →", "/guide#coverage").classes("guide-jump")
            unavailable = [
                item
                for item in result.coverage
                if item.state is CoverageState.UNAVAILABLE
            ]
            degraded = [
                item for item in result.coverage if item.state is CoverageState.DEGRADED
            ]
            checked = [
                item for item in result.coverage if item.state is CoverageState.CHECKED
            ]
            with ui.element("div").classes("statstrip"):
                for tone, label, items in (
                    ("critical", "unavailable", unavailable),
                    ("warning", "degraded", degraded),
                    ("expected", "checked", checked),
                ):
                    with ui.column().classes(f"stat stat-{tone}"):
                        ui.label(str(len(items))).classes("n")
                        ui.label(f"{label} checks").classes("l")
            for item in (*unavailable, *degraded):
                with ui.element("div").classes("covlimit"):
                    ui.label(f"{item.artifact} · {item.label}").classes("t")
                    ui.label(_limitation_consequence(item)).classes("d")
                    if item.detail:
                        ui.label(item.detail).classes("d")
            with ui.expansion(
                "Full coverage matrix",
                value=not (unavailable or degraded),
            ).classes("w-full mt-2"):
                ui.table(
                    columns=[
                        {"name": "artifact", "label": "Artifact", "field": "artifact"},
                        {"name": "check", "label": "Check", "field": "check"},
                        {"name": "status", "label": "Status", "field": "status"},
                        {"name": "findings", "label": "Findings", "field": "findings"},
                        {"name": "detail", "label": "Detail", "field": "detail"},
                    ],
                    rows=[
                        {
                            "artifact": item.artifact,
                            "check": item.label,
                            "status": item.state.value,
                            "findings": item.findings,
                            "detail": item.detail,
                        }
                        for item in result.coverage
                    ],
                    row_key="check",
                    pagination=20,
                ).classes("coverage-table").props("flat dense hide-bottom")

                # Alignment trust per-region (compact factual table)
                if result.alignment_trust is not None:
                    manifest = result.alignment_trust
                    with ui.expansion("Alignment trust per region").classes("w-full mt-2"):
                        rows = []
                        for r in manifest.regions:
                            rows.append(
                                {
                                    "key": (
                                        f"{r.artifact_member}|{r.sheet}|"
                                        f"{r.region_id}|{r.current_range}"
                                    ),
                                    "member": r.artifact_member,
                                    "sheet": r.sheet,
                                    "region": r.region_id,
                                    "baseline_range": r.baseline_range,
                                    "current_range": r.current_range,
                                    "row_method": r.row.method,
                                    "row_fallback": r.row.low_confidence_fallback,
                                    "col_method": r.column.method,
                                    "col_fallback": r.column.low_confidence_fallback,
                                    "paired_rows": r.row.paired,
                                    "paired_cols": r.column.paired,
                                    "cell_pairs": r.comparable_cell_pairs,
                                    "skipped_cells": r.skipped_low_confidence_cells,
                                    "deleted_rows": r.row.deleted,
                                    "inserted_rows": r.row.inserted,
                                    "growth_rows": r.row.growth,
                                    "deleted_cols": r.column.deleted,
                                    "inserted_cols": r.column.inserted,
                                    "growth_cols": r.column.growth,
                                    "low_confidence": r.low_confidence,
                                }
                            )
                        if rows:
                            with ui.element("div").classes("ctxblock w-full"):
                                ui.table(
                                    columns=[
                                        {
                                            "name": name,
                                            "label": label,
                                            "field": name,
                                        }
                                        for name, label in (
                                            ("member", "Member"),
                                            ("sheet", "Sheet"),
                                            ("region", "Region"),
                                            ("baseline_range", "Baseline range"),
                                            ("current_range", "Current range"),
                                            ("row_method", "Row method"),
                                            ("row_fallback", "Row fallback"),
                                            ("col_method", "Col method"),
                                            ("col_fallback", "Col fallback"),
                                            ("paired_rows", "Paired rows"),
                                            ("paired_cols", "Paired cols"),
                                            ("cell_pairs", "Cell pairs"),
                                            ("skipped_cells", "Skipped cells"),
                                            ("deleted_rows", "Del rows"),
                                            ("inserted_rows", "Ins rows"),
                                            ("growth_rows", "Growth rows"),
                                            ("deleted_cols", "Del cols"),
                                            ("inserted_cols", "Ins cols"),
                                            ("growth_cols", "Growth cols"),
                                            ("low_confidence", "Low confidence"),
                                        )
                                    ],
                                    rows=rows,
                                    row_key="key",
                                    pagination=25,
                                ).classes("coverage-table").props("flat dense")
                        # unpaired regions
                        if manifest.unpaired:
                            unpaired_rows = [
                                {
                                    "key": (
                                        f"{u.side}|{u.artifact_member}|{u.sheet}|"
                                        f"{u.region_id}|{u.cell_range}"
                                    ),
                                    "side": u.side,
                                    "member": u.artifact_member,
                                    "sheet": u.sheet,
                                    "region": u.region_id,
                                    "range": u.cell_range,
                                    "orientation": u.orientation,
                                }
                                for u in manifest.unpaired
                            ]
                            ui.label("Unpaired regions").classes("dk")
                            with ui.element("div").classes("ctxblock w-full"):
                                ui.table(
                                    columns=[
                                        {
                                            "name": name,
                                            "label": label,
                                            "field": name,
                                        }
                                        for name, label in (
                                            ("side", "Side"),
                                            ("member", "Member"),
                                            ("sheet", "Sheet"),
                                            ("region", "Region"),
                                            ("range", "Range"),
                                            ("orientation", "Orientation"),
                                        )
                                    ],
                                    rows=unpaired_rows,
                                    row_key="key",
                                ).classes("coverage-table").props("flat dense")

        with ui.tab_panel("atomic"):
            ui.label(
                "Every atomic finding behind the review queue, with analyst "
                "severity overrides and comments."
            ).classes("note")
            if atomic_paged:
                ui.label(
                    f"{len(result.findings):,} atomic findings — this tab pages "
                    "server-side; severity pills filter the review queue, and "
                    "the Excel/JSON exports carry every finding."
                ).classes("hint")
            findings_table = (
                ui.table(
                    columns=_FINDINGS_COLUMNS,
                    rows=[],
                    row_key="id",
                    pagination=25,
                )
                .classes("findings-table")
                .props("flat dense")
            )
            findings_table.add_slot("body", FINDINGS_BODY_SLOT)
            if atomic_paged:
                with ui.element("div").classes("queuepager"):
                    atomic_controls["previous"] = ui.button(
                        icon="chevron_left",
                        on_click=lambda: _turn_atomic_page(-1),
                    ).props("flat dense")
                    atomic_controls["label"] = ui.label("").classes("meta")
                    atomic_controls["next"] = ui.button(
                        icon="chevron_right",
                        on_click=lambda: _turn_atomic_page(1),
                    ).props("flat dense")

        if result.mapping_coverage is not None:
            with ui.tab_panel("mapping"):
                _render_mapping_review(
                    result,
                    render_stats,
                    history=history,
                    run_id=run_id,
                    profiles_dir=profiles_dir,
                    mutable=mutable,
                )

    def _visible_group_rows() -> list[dict[str, object]]:
        nonlocal lens_entries
        severities = set(selected_severities)
        classes = set(selected_classes)
        state_value = str(review_state.value or "all")
        if summary_view is not None:
            annotated = _annotated_ids()
            prioritized_summaries = prioritize_review_summaries(
                summary_list,
                summary_aggregates,
                stories,
                reviewed_ids=annotated,
            )
            summary_rows: list[dict[str, object]] = []
            for item in prioritized_summaries:
                message = ""
                if item.summary.member_count == 1:
                    first = findings_by_id.get(item.summary.member_finding_ids[0])
                    message = first.message if first is not None else ""
                summary_rows.append(
                    _summary_review_row(
                        item.summary, item.rationale, annotated, message
                    )
                )
            # Series clustering needs member hydration; on very large runs the
            # canonical queue keeps everything reachable without it.
            active_lens[0] = None
            lens_entries = [
                LensEntry(row=row, children=(), hidden=0)
                for row in _filter_review_rows(
                    summary_rows,
                    severities=severities,
                    finding_classes=classes,
                    review_state=state_value,
                    needle=str(text_filter.value or ""),
                    story_scope=story_scope,
                    sheets=set(selected_wheres),
                )
            ]
        else:
            prioritized = prioritize_review(review_groups, stories)
            groups = [item.group for item in prioritized]
            rationales = {
                item.group.group_id: item.rationale for item in prioritized
            }
            try:
                lens = build_series_review_lens(groups, series_anchors)
            except SeriesLensError:
                logger.warning(
                    "series lens rejected; falling back to canonical queue"
                )
                lens = None
            active_lens[0] = lens if lens is not None and lens.clusters else None
            if lens is None or not lens.clusters:
                lens_entries = [
                    LensEntry(row=row, children=(), hidden=0)
                    for row in _filter_review_rows(
                        _review_group_rows(groups, rationales),
                        severities=severities,
                        finding_classes=classes,
                        review_state=state_value,
                        needle=str(text_filter.value or ""),
                        story_scope=story_scope,
                        sheets=set(selected_wheres),
                    )
                ]
            else:
                lens_entries = _lens_entries(
                    lens,
                    {group.group_id: group for group in groups},
                    rationales,
                    severities=severities,
                    finding_classes=classes,
                    review_state=state_value,
                    needle=str(text_filter.value or ""),
                    story_scope=story_scope,
                    sheets=set(selected_wheres),
                )
        pages = max(1, -(-len(lens_entries) // top_page["size"]))
        top_page["index"] = min(max(top_page["index"], 0), pages - 1)
        lens_entries = _sort_lens_entries(
            lens_entries,
            str(queue_order["key"]),
            descending=bool(queue_order["descending"]),
        )
        return _flatten_lens_rows(
            lens_entries,
            expanded=expanded_clusters,
            child_pages=child_pages,
            page=top_page["index"],
            page_size=top_page["size"],
        )

    def refilter() -> None:
        group_table.rows = _visible_group_rows()
        _render_top_pager()
        _refresh_severity_pills()
        if atomic_paged:
            _refresh_atomic_page()
        elif all_rows:
            findings_table.rows = [
                row for row in all_rows if row["severity"] in selected_severities
            ]

    def _refresh_atomic_page() -> None:
        if not atomic_paged or not atomic_loaded[0]:
            return
        total = len(result.findings)
        pages = max(1, -(-total // _ATOMIC_PAGE_SIZE))
        atomic_page["index"] = min(max(atomic_page["index"], 0), pages - 1)
        rows = _atomic_page_rows(result, atomic_page["index"], mutable=mutable)
        for row in rows:
            patch = atomic_overlay.get(str(row["id"]))
            if patch:
                row.update(patch)
        findings_table.rows = rows
        start = atomic_page["index"] * _ATOMIC_PAGE_SIZE
        end = min(start + _ATOMIC_PAGE_SIZE, total)
        atomic_controls["label"].set_text(
            f"Findings {start + 1:,}-{end:,} of {total:,}"
        )
        atomic_controls["previous"].set_enabled(atomic_page["index"] > 0)
        atomic_controls["next"].set_enabled(end < total)

    def _overlay_finding(finding: Finding) -> None:
        if not atomic_paged:
            return
        atomic_overlay[finding.finding_id] = {
            "severity": finding.severity.value if finding.severity else "warning",
            "overridden": finding.severity_overridden,
            "comment": finding.analyst_comment,
        }

    def _overlay_updates(updates: Iterable[ReviewAnnotation]) -> None:
        """Mirror persisted bulk decisions onto server-built pages."""
        if not atomic_paged:
            return
        for update in updates:
            entry = atomic_overlay.setdefault(update.finding_id, {})
            if update.severity is not None:
                entry["severity"] = update.severity
                entry["overridden"] = True
            entry["comment"] = update.comment

    def _turn_atomic_page(delta: int) -> None:
        atomic_page["index"] += delta
        _refresh_atomic_page()

    def _render_top_pager() -> None:
        pages = max(1, -(-len(lens_entries) // top_page["size"]))
        # the footer carries the order and page-size controls, so it stays
        # visible whenever there are rows; only multi-page runs need arrows
        top_pager.set_visibility(bool(lens_entries))
        top_previous.set_visibility(pages > 1)
        top_next.set_visibility(pages > 1)
        start = top_page["index"] * top_page["size"]
        end = min(start + top_page["size"], len(lens_entries))
        if lens_entries:
            top_page_label.set_text(
                f"Review rows {start + 1:,}-{end:,} of {len(lens_entries):,}"
            )
        else:
            top_page_label.set_text("No review rows match the filters")
        top_previous.set_enabled(top_page["index"] > 0)
        top_next.set_enabled(end < len(lens_entries))

    def _turn_top_page(delta: int) -> None:
        top_page["index"] += delta
        refilter()
        # a fresh page starts at its first row, not mid-scroll
        ui.run_javascript(
            "document.querySelector('.review-groups-table .q-table__middle')"
            "?.scrollTo(0, 0)"
        )

    def _set_page_size(event: events.ValueChangeEventArguments) -> None:
        size = int(event.value)
        if size not in _TOP_PAGE_SIZES:
            return
        start = top_page["index"] * top_page["size"]
        top_page["size"] = size
        top_page["index"] = start // size  # keep the first visible row stable
        app.storage.general["review_page_size"] = size
        refilter()

    def _set_queue_order(event: events.ValueChangeEventArguments) -> None:
        key = str(event.value)
        if key not in _QUEUE_ORDERS:
            return
        queue_order["key"] = key
        top_page["index"] = 0
        refilter()

    def _flip_queue_direction() -> None:
        queue_order["descending"] = not queue_order["descending"]
        direction_button.props(
            f'icon={"north" if queue_order["descending"] else "south"}'
        )
        top_page["index"] = 0
        refilter()

    def on_toggle_cluster(e: events.GenericEventArguments) -> None:
        cluster_id = str(e.args.get("id", ""))
        if cluster_id in expanded_clusters:
            expanded_clusters.discard(cluster_id)
        else:
            expanded_clusters.add(cluster_id)
        refilter()
        for row in group_table.rows:
            row["sel"] = row["id"] == cluster_id
        group_table.update()
        render_detail(cluster_id)
        ui.run_javascript(
            f'document.querySelector(\'[data-review-id="{cluster_id}"]\')?.focus()'
        )

    def on_child_page(e: events.GenericEventArguments) -> None:
        cluster_id = str(e.args.get("id", ""))
        delta = int(e.args.get("delta", 0))
        child_pages[cluster_id] = max(0, child_pages.get(cluster_id, 0) + delta)
        refilter()

    def refresh_review_groups() -> None:
        nonlocal review_groups
        if summary_view is None:
            review_groups = build_pattern_groups(result.findings)
        # Summary mode keeps stored groups static: decisions update the
        # reviewed counters and deferral (recomputed in _visible_group_rows)
        # but never regroup rows until re-QC.
        refilter()

    def clear_story_scope() -> None:
        story_scope.clear()
        story_pin.visible = False
        refilter()

    story_pin.on("click", clear_story_scope)
    text_filter.on_value_change(refilter)
    review_state.on_value_change(refilter)

    def on_view_change(e: events.ValueChangeEventArguments) -> None:
        # Atomic rows cost real DOM, so they are built the first time they matter.
        if e.value != "atomic":
            return
        if atomic_paged:
            if not atomic_loaded[0]:
                atomic_loaded[0] = True
                _refresh_atomic_page()
        elif not all_rows:
            all_rows[:] = _findings_rows(result, mutable=mutable)
            refilter()

    panels.on_value_change(on_view_change)

    def render_longitudinal(finding: Finding) -> None:
        if history is None or run_id is None:
            return
        active_history = history
        active_run_id = run_id
        loaded = False
        with ui.expansion("History across runs").classes("w-full") as expansion:
            content = ui.element("div").classes("w-full")
            with content:
                ui.label("Open to load bounded local history.").classes("hint")

        async def load_longitudinal(
            event: events.ValueChangeEventArguments,
        ) -> None:
            nonlocal loaded
            if not event.value or loaded:
                return
            loaded = True
            content.clear()
            try:
                dossier = await asyncio.to_thread(
                    active_history.get_dossier,
                    active_run_id,
                    finding.finding_id,
                )
                recurrence = await asyncio.to_thread(
                    active_history.recurrence_eligibility_from_dossier,
                    dossier,
                )
            except (KeyError, RuntimeError, ValueError):
                with content:
                    ui.label(
                        "Longitudinal evidence is unavailable for this stored run."
                    ).classes("hint")
                return
            with content:
                ui.label(
                    "Values and formulas appear only when that run stored them in "
                    "a finding. No stored observation is not a clean or unchanged "
                    "verdict."
                ).classes("note")
                rows = [
                    {
                        "run": entry.run_id,
                        "date": entry.started_at[:10],
                        "status": entry.status.value.replace("_", " "),
                        "finalized": "yes" if entry.finalized else "no",
                        "engine": entry.engine_severity or "",
                        "analyst": entry.analyst_severity or "",
                        "origin": (
                            entry.origin.value if entry.origin is not None else ""
                        ),
                        "baseline": entry.baseline_value or "",
                        "current": entry.current_value or "",
                        "comment": entry.analyst_comment,
                        "carry": (
                            f"run #{entry.carry_source_run_id}"
                            if entry.carry_source_run_id is not None
                            else ""
                        ),
                        "waiver": (
                            f"{entry.waiver_reason} ({entry.waiver_expires})"
                            if entry.waiver_reason
                            else ""
                        ),
                        "promotion": entry.promotion_kind or "",
                    }
                    for entry in dossier.entries
                ]
                with ui.element("div").classes("ctxblock w-full"):
                    ui.table(
                        columns=[
                            {"name": "run", "label": "Run", "field": "run"},
                            {"name": "date", "label": "Date", "field": "date"},
                            {
                                "name": "status",
                                "label": "Evidence",
                                "field": "status",
                            },
                            {
                                "name": "finalized",
                                "label": "Finalized",
                                "field": "finalized",
                            },
                            {
                                "name": "engine",
                                "label": "Engine",
                                "field": "engine",
                            },
                            {
                                "name": "analyst",
                                "label": "Analyst",
                                "field": "analyst",
                            },
                            {
                                "name": "origin",
                                "label": "Origin",
                                "field": "origin",
                            },
                            {
                                "name": "baseline",
                                "label": "Stored baseline",
                                "field": "baseline",
                            },
                            {
                                "name": "current",
                                "label": "Stored current",
                                "field": "current",
                            },
                            {
                                "name": "comment",
                                "label": "Comment",
                                "field": "comment",
                            },
                            {
                                "name": "carry",
                                "label": "Carried from",
                                "field": "carry",
                            },
                            {
                                "name": "waiver",
                                "label": "Waiver",
                                "field": "waiver",
                            },
                            {
                                "name": "promotion",
                                "label": "Promotion",
                                "field": "promotion",
                            },
                        ],
                        rows=rows,
                        row_key="run",
                        pagination=50,
                    ).classes("coverage-table").props("flat dense hide-bottom")
                if (
                    recurrence is not None
                    and profiles_dir is not None
                    and result.profile_name != "default"
                    and available_promotions(finding)
                ):
                    with ui.element("div").classes("notecard") as nudge:
                        ui.label(
                            "This exact decision recurred across "
                            f"{len(recurrence.run_ids)} finalized runs with "
                            f"{recurrence.manual_count} manual confirmations."
                        )
                        with ui.row().classes("items-center gap-2"):
                            ui.button(
                                "Preview contract promotion",
                                on_click=lambda: open_promotion_dialog(finding),
                            ).classes("ghostbtn").props("flat no-caps dense")
                            ui.button(
                                "Dismiss",
                                on_click=lambda: nudge.set_visibility(False),
                            ).props("flat no-caps dense")

        expansion.on_value_change(load_longitudinal)

    def _render_cluster_context(members: tuple[Finding, ...], selected_id: str) -> None:
        # One shared page index: the pager sits between the two grids and
        # turns BOTH sides so the analyst always compares the same panel.
        shared_index = {"value": 0}
        sides: list[tuple[str, ClusterContext, ui.element]] = []
        pager_slot: ui.element | None = None
        for side, label in (("current", "current"), ("baseline", "baseline")):
            context = build_cluster_context(
                members, side=side, selected_finding_id=selected_id
            )
            if not context.panels:
                continue
            if sides and pager_slot is None:
                pager_slot = ui.element("div").classes("w-full")
            ui.label(f"{label} context").classes("dk")
            holder = ui.element("div").classes("w-full")
            sides.append((label, context, holder))
            if not context.merged:
                ui.label(
                    "Stored excerpts disagree for a shared cell, so each excerpt "
                    "is shown separately."
                ).classes("hint")
            if not context.complete:
                ui.label(context.coverage_note).classes("hint")

        def draw_sides() -> None:
            for label, context, holder in sides:
                panel = min(shared_index["value"], len(context.panels) - 1)
                holder.clear()
                with holder:
                    ui.html(_cluster_context_html(context.panels[panel], label))

        draw_sides()
        total_panels = max((len(c.panels) for _l, c, _h in sides), default=0)
        if total_panels > 1:
            pager_host = pager_slot if pager_slot is not None else ui.element("div")

            def turn(delta: int) -> None:
                shared_index["value"] = max(
                    0, min(shared_index["value"] + delta, total_panels - 1)
                )
                draw_sides()
                index_label.set_text(
                    f"Panel {shared_index['value'] + 1} of {total_panels}"
                    " · both sides"
                )

            def _turn_handler(
                step: int, turn: Callable[[int], None] = turn
            ) -> Callable[[], None]:
                # zero-arg closure; a bare lambda would receive the click event
                return lambda: turn(step)

            with pager_host, ui.row().classes("items-center gap-2 ctxpagerrow"):
                ui.button(
                    icon="chevron_left", on_click=_turn_handler(-1)
                ).props("flat round dense aria-label='Previous context panel'")
                index_label = ui.label().classes("hint")
                ui.button(
                    icon="chevron_right", on_click=_turn_handler(1)
                ).props("flat round dense aria-label='Next context panel'")
            index_label.set_text(f"Panel 1 of {total_panels} · both sides")

    def _severity_legend_html() -> str:
        return (
            '<div class="ctxlegend" role="note" aria-label="Severity legend">'
            + "".join(
                f'<span class="key"><span class="swatch sev-{severity.value}"></span>'
                f"{severity.value}</span>"
                for severity in _SEVERITY_ORDER
            )
            + "</div>"
        )

    def _render_cluster_detail(cluster: SeriesCluster) -> None:
        row = next(
            (
                entry.row
                for entry in lens_entries
                if str(entry.row["id"]) == cluster.cluster_id
            ),
            None,
        )
        hidden = int(str(row["hidden_children"])) if row is not None else 0
        with detail_box:
            with ui.element("div").classes("detailhead"):
                ui.label(cluster.cluster_id).classes("detail-id")
                ui.label(cluster.label).classes("detail-msg")
                ui.label(
                    "related-series lens · not a new decision · "
                    f"{len(cluster.slices):,} canonical decisions · "
                    f"{cluster.member_count:,} finding"
                    f"{'s' if cluster.member_count != 1 else ''}"
                ).classes("runmeta")
                with ui.row().classes("detailactions"):
                    ui.button(
                        "View affected findings",
                        on_click=lambda: open_members(cluster.cluster_id),
                    ).classes("ghostbtn").props("no-caps flat dense")
                    if mutable:
                        ui.button(
                            "Confirm visible",
                            on_click=lambda: open_cluster_confirmation(
                                cluster.cluster_id
                            ),
                        ).classes("ghostbtn").props("no-caps flat dense")
                # pinned above the scroll so the outline key stays readable
                # while the grids below are paged and scrolled
                ui.html(_severity_legend_html())
            with ui.element("div").classes("detailbody"):
                if row is not None and row["sevmix"]:
                    ui.label(str(row["sevmix"])).classes("note")
                if hidden:
                    ui.label(
                        f"{hidden} child segment{'s' if hidden != 1 else ''} hidden by "
                        "the current filters."
                    ).classes("hint")
                if any(
                    member.finding_class is FindingClass.FINDINGS_CAPPED
                    for member in cluster.members
                ):
                    ui.label(
                        "Output budgets capped this run, so this parent lists retained "
                        "evidence, not the complete population."
                    ).classes("hint")
                for child in cluster.slices:
                    member = child.members[0]
                    anchor = series_anchors.get(member.finding_id)
                    segment = anchor_segment(anchor) if anchor is not None else ""
                    if segment == "new_period":
                        # A first-time period has no baseline number, so it has no
                        # temporal tier to show; a cleared period has no current one.
                        band = "new period"
                    elif segment == "cleared_period":
                        band = "cleared period"
                    else:
                        band = (
                            member.temporal_context.value
                            if member.temporal_context is not None
                            else "no temporal context"
                        )
                    severity = (member.severity or Severity.WARNING).value
                    ui.label(
                        f"{child.group_id} · {severity} · {band} · "
                        f"{child.member_count:,} finding"
                        f"{'s' if child.member_count != 1 else ''}"
                    ).classes("mono")
                _render_cluster_context(cluster.members, "")

    def _render_slice_detail(item: ReviewSlice, group: ReviewGroup) -> None:
        member = item.members[0]
        with detail_box:
            with ui.element("div").classes("detailhead"):
                ui.label(f"{item.group_id} · segment").classes("detail-id")
                ui.label(
                    member.message
                    if item.member_count == 1
                    else f"{item.member_count:,} findings — {member.message}"
                ).classes("detail-msg")
                temporal = (
                    member.temporal_context.value
                    if member.temporal_context is not None
                    else ""
                )
                ui.label(
                    f"{group.sheet or group.slide or ''} · "
                    f"{(member.severity or Severity.WARNING).value}"
                    + (f" · {temporal}" if temporal else "")
                ).classes("runmeta")
                with ui.row().classes("detailactions"):
                    ui.button(
                        "View affected findings",
                        on_click=lambda: open_members(item.slice_id),
                    ).classes("ghostbtn").props("no-caps flat dense")
                    if mutable:
                        ui.button(
                            "Review group",
                            on_click=lambda: open_group_review(item.group_id),
                        ).classes("ghostbtn").props("no-caps flat dense")
            with ui.element("div").classes("detailbody"):
                _render_evidence_body(member)
                if item.member_count == 1:
                    render_longitudinal(member)
                    if focus_actions is not None:
                        focus_actions(member)

    def render_detail(group_id: str) -> None:
        detail_box.clear()
        lens = active_lens[0]
        if lens is not None:
            cluster = lens.cluster_by_id.get(group_id)
            if cluster is not None:
                _render_cluster_detail(cluster)
                return
            item = lens.slice_by_id.get(group_id)
            if item is not None and not item.whole_group:
                group = next(
                    (
                        candidate
                        for candidate in review_groups
                        if candidate.group_id == item.group_id
                    ),
                    None,
                )
                if group is not None:
                    _render_slice_detail(item, group)
                    return
            if item is not None:
                group_id = item.group_id
        group = next(
            (item for item in review_groups if item.group_id == group_id), None
        )
        if group is None:
            return
        member = group.members[0]
        with detail_box:
            with ui.element("div").classes("detailhead"):
                ui.label(group.group_id).classes("detail-id")
                ui.label(
                    member.message
                    if group.member_count == 1
                    else f"{group.member_count:,} findings — {member.message}"
                ).classes("detail-msg")
                ui.label(
                    f"{group.sheet or group.slide or ''} · {format_group_ranges(group)}"
                ).classes("runmeta")
                with ui.row().classes("detailactions"):
                    ui.button(
                        "View affected findings",
                        on_click=lambda: open_members(group_id),
                    ).classes("ghostbtn").props("no-caps flat dense")
                    if mutable:
                        ui.button(
                            "Review group",
                            on_click=lambda: open_group_review(group_id),
                        ).classes("ghostbtn").props("no-caps flat dense")
            with ui.element("div").classes("detailbody"):
                rationale = next(
                    (
                        str(row.get("why") or "")
                        for row in group_table.rows
                        if row["id"] == group_id
                    ),
                    "",
                )
                if rationale:
                    ui.label(f"prioritized because: {rationale}").classes("note")
                _render_evidence_body(member)
                if group.member_count == 1:
                    render_longitudinal(member)
                # A multi-member group must never focus an arbitrary
                # representative; members are chosen in the affected dialog.
                if focus_actions is not None and group.member_count == 1:
                    focus_actions(member)
                if (
                    group.member_count == 1
                    and history is not None
                    and run_id is not None
                    and profiles_dir is not None
                    and result.profile_name != "default"
                ):
                    render_promotion_action(member)

    def on_select(e: events.GenericEventArguments) -> None:
        group_id = str(e.args.get("id", ""))
        for row in group_table.rows:
            row["sel"] = row["id"] == group_id
        group_table.update()
        render_detail(group_id)

    def _members_for_row(row_id: str) -> tuple[str, str, tuple[Finding, ...]] | None:
        """Row id to (title, location line, exact members) for the members view."""
        lens = active_lens[0]
        if lens is not None:
            cluster = lens.cluster_by_id.get(row_id)
            if cluster is not None:
                return (
                    f"{cluster.cluster_id} — {cluster.member_count:,} "
                    "affected findings",
                    cluster.label,
                    cluster.members,
                )
            item = lens.slice_by_id.get(row_id)
            if item is not None and not item.whole_group:
                group = next(
                    (
                        candidate
                        for candidate in review_groups
                        if candidate.group_id == item.group_id
                    ),
                    None,
                )
                if group is not None:
                    return (
                        f"{item.group_id} — {item.member_count:,} affected findings",
                        f"{group.sheet or group.slide or ''}",
                        item.members,
                    )
            if item is not None:
                row_id = item.group_id
        group = next(
            (item for item in review_groups if item.group_id == row_id), None
        )
        if group is None:
            return None
        return (
            f"{group.group_id} — {group.member_count:,} affected findings",
            f"{group.sheet or group.slide or ''} · {format_group_ranges(group)}",
            group.members,
        )

    def open_members(group_id: str) -> None:
        resolved = _members_for_row(group_id)
        if resolved is None:
            return
        title, location_line, members = resolved
        member_count = len(members)
        page_size = 50
        page = {"index": 0}
        with ui.dialog() as dialog, ui.card().classes("memberscard"):
            with ui.row().classes("items-baseline gap-3 w-full"):
                ui.label(title).classes("runhead")
                ui.label(location_line).classes("runmeta")
            with ui.element("div").classes("reviewsplit memberssplit"):
                member_table = (
                    ui.table(
                        columns=_MEMBER_COLUMNS,
                        rows=[],
                        row_key="id",
                        pagination=page_size,
                    )
                    .classes("findings-table members-table")
                    .props("flat dense hide-bottom")
                )
                member_table.add_slot("body", REVIEW_MEMBER_ROWS_SLOT)
                member_detail = ui.element("div").classes("detailpanel")

            def show_member(finding_id: str) -> None:
                member = findings_by_id.get(finding_id)
                member_detail.clear()
                if member is None:
                    return
                for row in member_table.rows:
                    row["sel"] = row["id"] == finding_id
                member_table.update()
                with member_detail:
                    with ui.element("div").classes("detailhead"):
                        ui.label(member.finding_id).classes("detail-id")
                        ui.label(member.message).classes("detail-msg")
                    with ui.element("div").classes("detailbody"):
                        _render_evidence_body(member)
                        render_longitudinal(member)
                        if focus_actions is not None:
                            focus_actions(member)
                        if (
                            history is not None
                            and run_id is not None
                            and profiles_dir is not None
                            and result.profile_name != "default"
                        ):
                            render_promotion_action(member)

            def on_member_select(e: events.GenericEventArguments) -> None:
                show_member(str(e.args.get("id", "")))

            member_table.on("select", on_member_select)

            def refresh_page() -> None:
                start = page["index"] * page_size
                end = min(start + page_size, member_count)
                member_table.rows = [
                    _finding_row(member, mutable=mutable)
                    for member in members[start:end]
                ]
                page_label.set_text(
                    f"Showing {start + 1:,}-{end:,} of {member_count:,}"
                )
                previous_button.set_enabled(page["index"] > 0)
                next_button.set_enabled(end < member_count)
                show_member(members[start].finding_id)

            def previous_page() -> None:
                page["index"] = max(0, page["index"] - 1)
                refresh_page()

            def next_page() -> None:
                page["index"] += 1
                refresh_page()

            with ui.row().classes("items-center gap-2 w-full"):
                previous_button = ui.button(
                    icon="chevron_left", on_click=previous_page
                ).props("flat round dense aria-label='Previous member page'")
                next_button = ui.button(
                    icon="chevron_right", on_click=next_page
                ).props("flat round dense aria-label='Next member page'")
                page_label = ui.label().classes("hint")
                ui.button("Close", on_click=dialog.close).classes("ml-auto").props(
                    "flat no-caps"
                )
            refresh_page()
            dialog.on(
                "hide",
                lambda: ui.run_javascript(
                    f'document.querySelector(\'[data-review-id="{group_id}"]\')?.focus()'
                ),
            )
        dialog.open()

    def open_promotion_dialog(finding: Finding) -> None:
        if history is None or run_id is None or profiles_dir is None:
            return
        active_history = history
        active_run_id = run_id
        path = _profile_path(profiles_dir, result.profile_name)
        profile = load_profile(path)
        opened_hash = profile_sha256(profile)
        kinds = available_promotions(finding)
        with ui.dialog() as dialog, ui.card().classes("w-[42rem] max-w-full"):
            ui.label("Promote finding to contract").classes("runhead")
            ui.label(
                "This updates the named profile for future runs. The current "
                "run and any finalized evidence remain unchanged."
            ).classes("note")
            kind = ui.select(
                {item.value: item.value.replace("_", " ") for item in kinds},
                value=kinds[0].value,
                label="Rule type",
            ).classes("w-full").props("outlined dense")
            name = ui.input("Control name").classes("w-full").props(
                "outlined dense"
            )
            reason = ui.input("Waiver reason").classes("w-full").props(
                "outlined dense"
            )
            expires = ui.input("Waiver expiry (YYYY-MM-DD)").classes(
                "w-full"
            ).props("outlined dense")
            with ui.row().classes("gap-2 w-full"):
                minimum = ui.number("Minimum").classes("flex-1").props(
                    "outlined dense"
                )
                maximum = ui.number("Maximum").classes("flex-1").props(
                    "outlined dense"
                )
                absolute = ui.number(
                    "Absolute tolerance", value=0, min=0
                ).classes("flex-1").props("outlined dense")
                relative = ui.number(
                    "Relative tolerance", value=0, min=0
                ).classes("flex-1").props("outlined dense")

            async def apply_promotion() -> None:
                try:
                    expiry = (
                        dt.date.fromisoformat(str(expires.value))
                        if expires.value
                        else None
                    )
                    request = PromotionRequest(
                        kind=PromotionKind(str(kind.value)),
                        name=str(name.value or ""),
                        reason=str(reason.value or ""),
                        expires=expiry,
                        minimum=(
                            float(minimum.value)
                            if minimum.value is not None
                            else None
                        ),
                        maximum=(
                            float(maximum.value)
                            if maximum.value is not None
                            else None
                        ),
                        absolute=float(absolute.value or 0),
                        relative=float(relative.value or 0),
                    )
                    draft = promote_finding(profile, finding, request)
                    current = load_profile(path)
                    if profile_sha256(current) != opened_hash:
                        raise ValueError(
                            "profile changed while this dialog was open; reopen it"
                        )
                    record = active_history.get_run(active_run_id)
                    issues = await asyncio.to_thread(
                        _lint_profile_for_record, draft, record
                    )
                    errors = [issue for issue in issues if issue.level == "error"]
                    if errors:
                        raise ValueError(errors[0].message)
                    save_profile(draft, path)
                    try:
                        active_history.record_promotion(
                            active_run_id,
                            finding.finding_id,
                            request.kind.value,
                            profile_sha256(draft),
                        )
                    except Exception:
                        logger.error("promotion-audit-record-failed")
                        dialog.close()
                        ui.notify(
                            "Profile updated, but its promotion audit record "
                            "could not be stored. Re-QC before further review.",
                            type="negative",
                        )
                        return
                except Exception as exc:
                    ui.notify(str(exc), type="negative")
                    return
                dialog.close()
                ui.notify("Profile updated; run Re-QC to apply the new contract")

            with ui.row().classes("items-center gap-2"):
                ui.button(
                    "Preview and save", on_click=apply_promotion
                ).classes("runbtn").props("no-caps")
                ui.button("Cancel", on_click=dialog.close).props("flat no-caps")
        dialog.open()

    def render_promotion_action(
        finding: Finding,
        button_label: str = "Promote to contract",
    ) -> None:
        ui.button(
            button_label,
            on_click=lambda: open_promotion_dialog(finding),
        ).classes("ghostbtn").props("flat no-caps dense")

    def persist_group_review(
        group: ReviewGroup,
        *,
        severity: Severity | None,
        comment: str,
        replace_existing: bool,
    ) -> list[ReviewAnnotation]:
        before = [
            (
                finding,
                finding.severity,
                finding.severity_overridden,
                finding.analyst_comment,
            )
            for finding in group.members
        ]
        updates = apply_group_review(
            group,
            severity=severity,
            comment=comment,
            replace_existing=replace_existing,
        )
        if not updates:
            return []
        try:
            if history is not None and run_id is not None:
                history.set_annotations_bulk(
                    run_id,
                    [
                        (update.finding_id, update.severity, update.comment)
                        for update in updates
                    ],
                )
        except Exception:
            for finding, old_severity, old_overridden, old_comment in before:
                finding.severity = old_severity
                finding.severity_overridden = old_overridden
                finding.analyst_comment = old_comment
            raise
        return updates

    def open_cluster_confirmation(cluster_id: str) -> None:
        """Confirm only the visible child segments, each at its own severity."""
        lens = active_lens[0]
        if lens is None or not mutable:
            return
        cluster = lens.cluster_by_id.get(cluster_id)
        entry = next(
            (item for item in lens_entries if str(item.row["id"]) == cluster_id),
            None,
        )
        if cluster is None or entry is None:
            return
        visible = tuple(
            item
            for item in cluster.slices
            if item.slice_id in {str(child["id"]) for child in entry.children}
        )
        updates = cluster_confirmation_updates(visible, "")
        counts: dict[str, int] = {}
        for update in updates:
            counts[update.severity or ""] = counts.get(update.severity or "", 0) + 1
        hidden_members = cluster.member_count - sum(
            item.member_count for item in visible
        )
        with ui.dialog() as dialog, ui.card().classes("w-[34rem] max-w-full"):
            ui.label(f"Confirm visible · {cluster.label}").classes("runhead")
            ui.label(
                "Each finding keeps the severity shown now. This is a "
                "confirmation, not an override or a replacement."
            ).classes("lede")
            if updates:
                ui.label(
                    "will confirm "
                    + " · ".join(
                        f"{count} {severity}"
                        for severity, count in sorted(counts.items())
                    )
                    + f" ({len(updates):,} finding"
                    + ("s)" if len(updates) != 1 else ")")
                ).classes("note")
            else:
                ui.label(
                    "Every visible finding already carries an analyst decision."
                ).classes("note")
            if entry.hidden:
                ui.label(
                    f"{entry.hidden} child segment"
                    f"{'s' if entry.hidden != 1 else ''} and {hidden_members:,} "
                    "finding"
                    f"{'s' if hidden_members != 1 else ''} are hidden by the "
                    "current filters and stay untouched."
                ).classes("hint")
            comment_input = (
                ui.input("Shared comment (optional)")
                .classes("w-full")
                .props("outlined dense")
            )

            def apply_confirmation() -> None:
                pending = cluster_confirmation_updates(
                    visible, str(comment_input.value or "")
                )
                if not pending:
                    ui.notify("No visible findings need confirmation", type="warning")
                    return
                try:
                    # Commit first: a database failure must leave every
                    # in-memory finding and the current selection untouched.
                    if history is not None and run_id is not None:
                        history.set_annotations_bulk(
                            run_id,
                            [
                                (update.finding_id, update.severity, update.comment)
                                for update in pending
                            ],
                        )
                except Exception as exc:
                    ui.notify(str(exc), type="negative")
                    return
                dialog.close()
                try:
                    for update in pending:
                        finding = findings_by_id.get(update.finding_id)
                        if finding is None:
                            continue
                        finding.severity = Severity(str(update.severity))
                        finding.severity_overridden = True
                        if update.comment:
                            finding.analyst_comment = update.comment
                        _overlay_finding(finding)
                    if all_rows:
                        all_rows[:] = _findings_rows(result, mutable=mutable)
                    refresh_review_groups()
                    render_stats()
                    render_detail(cluster_id)
                except Exception:
                    # The transaction already committed, so the database is
                    # authoritative; never restore the stale pre-commit view.
                    logger.error("cluster-confirmation-rebuild-failed")
                    ui.notify(
                        "Your decisions were saved. Reload this run before "
                        "reviewing anything else.",
                        type="warning",
                    )
                    ui.run_javascript("window.location.reload()")
                    return
                ui.notify(f"Confirmed {len(pending):,} visible findings")

            with ui.row().classes("items-center gap-2"):
                ui.button(
                    "Confirm listed findings", on_click=apply_confirmation
                ).classes("runbtn").props("no-caps")
                ui.button("Cancel", on_click=dialog.close).props("flat no-caps")
            dialog.on(
                "hide",
                lambda: ui.run_javascript(
                    f'document.querySelector(\'[data-review-id="{cluster_id}"]\')'
                    "?.focus()"
                ),
            )
        dialog.open()

    def open_group_review(group_id: str) -> None:
        group = next(
            (item for item in review_groups if item.group_id == group_id), None
        )
        if group is None:
            return
        with ui.dialog() as dialog, ui.card().classes("w-[34rem] max-w-full"):
            ui.label(
                f"Review {group.group_id} · {group.member_count:,} findings"
            ).classes("runhead")
            ui.label(
                "Choose unreviewed only to preserve existing individual decisions, "
                "or replace all explicitly."
            ).classes("lede")
            severity_select = ui.select(
                ["keep", *[severity.value for severity in Severity]],
                value="keep",
                label="Group severity",
            ).classes("w-full").props("outlined dense")
            comment_input = ui.input("Group comment (blank keeps comments)").classes(
                "w-full"
            ).props("outlined dense")
            apply_mode = ui.toggle(
                {
                    "blank": "Unreviewed only",
                    "replace": "Replace all",
                },
                value="blank",
            ).classes("review-toggle").props("no-caps")

            def apply_review() -> None:
                selected = str(severity_select.value or "keep")
                severity = None if selected == "keep" else Severity(selected)
                try:
                    updates = persist_group_review(
                        group,
                        severity=severity,
                        comment=str(comment_input.value or ""),
                        replace_existing=apply_mode.value == "replace",
                    )
                except Exception as exc:
                    ui.notify(str(exc), type="negative")
                    return
                if not updates:
                    ui.notify("No group review changes to apply", type="warning")
                    return
                _overlay_updates(updates)
                if all_rows:
                    all_rows[:] = _findings_rows(result, mutable=mutable)
                # Refreshing the detail panel deregisters this dialog, so close
                # and report before anything is torn down.
                dialog.close()
                ui.notify(f"Updated {len(updates):,} affected findings")
                refresh_review_groups()
                render_stats()
                detail_box.clear()

            with ui.row().classes("items-center gap-2"):
                ui.button("Apply review", on_click=apply_review).classes(
                    "runbtn"
                ).props("no-caps")
                ui.button("Cancel", on_click=dialog.close).props("flat no-caps")
            dialog.on(
                "hide",
                lambda: ui.run_javascript(
                    f'document.querySelector(\'[data-review-id="{group_id}"]\')?.focus()'
                ),
            )
        dialog.open()

    group_table.on("select", on_select)
    group_table.on("toggle", on_toggle_cluster)
    group_table.on("childpage", on_child_page)

    def _reviewable_group(row_id: str) -> ReviewGroup | None:
        """Row id to the exact decision a row-level action may touch.

        A derived child or remainder row reviews only its own members; a parent
        row is a navigation lens and never resolves here.
        """
        lens = active_lens[0]
        if lens is not None:
            if row_id in lens.cluster_by_id:
                return None
            item = lens.slice_by_id.get(row_id)
            if item is not None:
                group = next(
                    (
                        candidate
                        for candidate in review_groups
                        if candidate.group_id == item.group_id
                    ),
                    None,
                )
                if group is None:
                    return None
                return (
                    group
                    if item.whole_group
                    else dataclasses.replace(group, members=item.members)
                )
        return next(
            (item for item in review_groups if item.group_id == row_id), None
        )

    def on_keyboard_triage(e: events.GenericEventArguments) -> None:
        if not mutable:
            return
        group_id = str(e.args.get("id", ""))
        action = str(e.args.get("action", ""))
        group = _reviewable_group(group_id)
        severities = {
            "critical": Severity.CRITICAL,
            "warning": Severity.WARNING,
            "info": Severity.INFO,
            "expected": Severity.EXPECTED,
            "confirm": group.severity if group is not None else None,
        }
        severity = severities.get(action)
        if group is None or severity is None:
            return
        visible_before = [str(row["id"]) for row in group_table.rows]
        current_index = (
            visible_before.index(group_id) if group_id in visible_before else -1
        )
        try:
            updates = persist_group_review(
                group,
                severity=severity,
                comment="",
                replace_existing=False,
            )
        except Exception as exc:
            ui.notify(str(exc), type="negative")
            return
        if not updates:
            ui.notify("This decision is already reviewed", type="warning")
            return
        _overlay_updates(updates)
        if all_rows:
            all_rows[:] = _findings_rows(result, mutable=mutable)
        refresh_review_groups()
        render_stats()

        visible_after = {
            str(row["id"]): row
            for row in group_table.rows
            if row.get("review_state") == "needs_review"
        }
        next_id = next(
            (
                candidate
                for candidate in visible_before[current_index + 1 :]
                if candidate in visible_after
            ),
            None,
        )
        if next_id is None:
            detail_box.clear()
        else:
            for row in group_table.rows:
                row["sel"] = row["id"] == next_id
            group_table.update()
            render_detail(next_id)
            ui.run_javascript(
                f'document.querySelector(\'[data-review-id="{next_id}"]\')?.focus()'
            )
        ui.notify(f"Reviewed {len(updates):,} affected findings")

    group_table.on("triage", on_keyboard_triage)

    def on_severity(e: events.GenericEventArguments) -> None:
        finding_id, value = e.args["id"], e.args["value"]
        finding = findings_by_id.get(finding_id)
        if finding is None:
            return
        finding.severity = Severity(value)
        finding.severity_overridden = True
        _overlay_finding(finding)
        for row in all_rows:
            if row["id"] == finding_id:
                row["severity"] = value
                row["overridden"] = True
        if history is not None and run_id is not None:
            history.set_annotation(
                run_id, finding_id, severity=value, comment=finding.analyst_comment
            )
        refresh_review_groups()
        render_stats()
        ui.notify(f"{finding_id}: severity set to {value} (analyst override)")

    def on_note(e: events.GenericEventArguments) -> None:
        finding_id, value = e.args["id"], str(e.args["value"] or "")
        finding = findings_by_id.get(finding_id)
        if finding is None or finding.analyst_comment == value:
            return
        finding.analyst_comment = value
        _overlay_finding(finding)
        for row in all_rows:
            if row["id"] == finding_id:
                row["comment"] = value
        if history is not None and run_id is not None:
            history.set_annotation(
                run_id,
                finding_id,
                severity=finding.severity.value
                if finding.severity_overridden and finding.severity
                else None,
                comment=value,
            )
        # A comment does not regroup, but it does change the reviewed marker.
        refilter()
        ui.notify(f"{finding_id}: comment saved")

    findings_table.on("sev", on_severity)
    findings_table.on("note", on_note)
    refilter()

    if mutable:
        ui.run_javascript(
            """
;(function () {
    if (window._qcTriageHandler) {
        window.removeEventListener('keydown', window._qcTriageHandler, true);
    }
    function activeElementIsEditable() {
        const a = document.activeElement; if (!a) return false;
        const t = a.tagName || '';
        return (t==='INPUT'||t==='SELECT'||t==='TEXTAREA'||a.isContentEditable);
    }
    function dialogIsOpen() {
        return Array.from(document.querySelectorAll('.q-dialog__inner'))
            .some(e => e.getClientRects().length > 0);
    }
    function visibleIds() {
        return Array.from(document.querySelectorAll('[data-review-id]'))
            .filter(e => e.getClientRects().length > 0)
            .map(e => e.dataset.reviewId);
    }
    function selectedId() {
        const sel = Array.from(document.querySelectorAll('[data-review-id].selrow'))
            .find(e => e.getClientRects().length > 0);
        if (sel) return sel.dataset.reviewId;
        const focused = document.activeElement;
        if (focused?.dataset?.reviewId) return focused.dataset.reviewId;
        return null;
    }
    function triage(action) {
        const selected = selectedId(); if (!selected) return;
        const row = document.querySelector('[data-review-id="'+selected+'"]');
        const control = row?.querySelector('[data-triage-action="'+action+'"]');
        control?.click();
    }
    function activate(id) {
        const el = document.querySelector('[data-review-id="' + id + '"]');
        if (!el) return;
        // Clicking a related-series parent toggles its expansion, so j/k only
        // moves focus there; Enter or Space still expands or collapses it.
        if (el.dataset.reviewKind === 'cluster') { el.focus(); return; }
        el.click();
    }
    window._qcTriageHandler = function (e) {
        if (activeElementIsEditable() || dialogIsOpen()) return;
        if (!visibleIds().length) return;
        const k = e.key;
        if (k==='j' || k==='ArrowDown') {
            e.preventDefault();
            const ids = visibleIds();
            if (!ids.length) return;
            const sel = selectedId();
            const idx = sel ? ids.indexOf(sel) : -1;
            const next = Math.min(ids.length - 1, Math.max(0, idx + 1));
            activate(ids[next]);
        }
        else if (k==='k' || k==='ArrowUp') {
            e.preventDefault();
            const ids = visibleIds();
            if (!ids.length) return;
            const sel = selectedId();
            const idx = sel ? ids.indexOf(sel) : ids.length;
            const previous = Math.max(0, idx - 1);
            activate(ids[previous]);
        }
        else if (k==='1') { e.preventDefault(); triage('critical'); }
        else if (k==='2') { e.preventDefault(); triage('warning'); }
        else if (k==='3') { e.preventDefault(); triage('info'); }
        else if (k==='4') { e.preventDefault(); triage('expected'); }
        else if (k==='c') { e.preventDefault(); triage('confirm'); }
    };
    window.addEventListener('keydown', window._qcTriageHandler, true);
})();
""",
        )


def _result_scope_line(result: QCRunResult) -> str:
    """The comparison scope this run actually applied."""
    scope = result.comparison_scope
    parts: list[str] = []
    if scope.excel_sheets:
        parts.append(f"{len(scope.excel_sheets)} selected sheets")
    if scope.ppt_slide_indices:
        parts.append(f"{len(scope.ppt_slide_indices)} selected slides")
    return "scope: " + (", ".join(parts) if parts else "everything compared")


def _render_mapping_review(
    result: QCRunResult,
    render_stats: Callable[[], None],
    *,
    history: RunHistory | None,
    run_id: int | None,
    profiles_dir: Path | None,
    mutable: bool = True,
) -> None:
    """Excel-to-PowerPoint figure reconciliation and analyst confirmation."""
    section("Mapping review")
    ui.link("Mapping guide →", "/guide#mappings").classes("guide-jump")
    mapping_stats = ui.element("div").classes("mappingstats")

    def render_mapping_stats() -> None:
        mapping_stats.clear()
        mapping = result.mapping_coverage
        if mapping is None:
            return
        with mapping_stats:
            for label, value in _mapping_stats(mapping):
                with ui.column().classes("mappingstat"):
                    ui.label(str(value)).classes("n")
                    ui.label(label).classes("l")

    render_mapping_stats()
    mapping_limitation = next(
        (
            item.detail
            for item in result.coverage
            if item.check_id == "excel-ppt-crosscheck"
            and item.state is not CoverageState.CHECKED
            and item.detail
        ),
        "",
    )
    if mapping_limitation:
        ui.label(mapping_limitation).classes("notecard")
    if result.profile_name == "default":
        ui.label("Named profile required to save confirmed mappings.").classes(
            "notecard"
        )
    suggestions_box = ui.column().classes("w-full gap-2")

    def refresh_mapping_coverage_detail() -> None:
        mapping = result.mapping_coverage
        if mapping is None:
            return
        item = next(
            (
                entry
                for entry in result.coverage
                if entry.check_id == "excel-ppt-crosscheck"
            ),
            None,
        )
        if item is not None:
            item.detail = (
                f"{mapping.eligible} eligible; {mapping.mapped} mapped; "
                f"{mapping.verified} verified; {mapping.mismatched} mismatched; "
                f"{mapping.unresolved} unresolved; {mapping.unmapped} unmapped"
            )

    def render_suggestions() -> None:
        suggestions_box.clear()
        with suggestions_box:
            if not result.mapping_suggestions:
                ui.label("All eligible figures are mapped.").classes("lede")
                return
            for suggestion in result.mapping_suggestions:
                title = (
                    f"{suggestion.slide} · {suggestion.figure_raw} · "
                    f"{suggestion.line_skeleton}"
                )
                with ui.expansion(title).classes("mappingitem"):
                    ui.label(suggestion.line).classes("mappingcontext")
                    if not suggestion.candidates:
                        ui.label("No source candidate found.").classes("lede")
                        continue
                    for candidate in suggestion.candidates:
                        with ui.row().classes("candidate-row"):
                            source_prefix = (
                                f"{candidate.source_member}:"
                                if candidate.source_member != "primary"
                                else ""
                            )
                            ui.label(
                                f"{source_prefix}{candidate.sheet}!{candidate.cell}"
                            ).classes("candidate-ref")
                            ui.label(str(candidate.value)).classes("candidate-value")
                            ui.label(
                                "display match"
                                if candidate.display_match
                                else "near match"
                            ).classes(
                                "candidate-match"
                                if candidate.display_match
                                else "candidate-near"
                            )

                            def confirm_handler(
                                suggestion: MappingSuggestion = suggestion,
                                candidate: SuggestedSource = candidate,
                            ):
                                def confirm() -> None:
                                    if profiles_dir is None:
                                        ui.notify(
                                            "Profile storage is unavailable",
                                            type="negative",
                                        )
                                        return
                                    try:
                                        if history is not None and run_id is not None:
                                            history.assert_mutable(run_id)
                                        persist_confirmed_mapping(
                                            profiles_dir,
                                            result.profile_name,
                                            suggestion,
                                            candidate,
                                        )
                                    except Exception as exc:
                                        ui.notify(str(exc), type="warning")
                                        return
                                    key = (
                                        suggestion.slide,
                                        suggestion.line_skeleton,
                                        suggestion.figure_index,
                                    )
                                    result.mapping_suggestions = [
                                        item
                                        for item in result.mapping_suggestions
                                        if (
                                            item.slide,
                                            item.line_skeleton,
                                            item.figure_index,
                                        )
                                        != key
                                    ]
                                    mapping = result.mapping_coverage
                                    if mapping is not None:
                                        mapping.mapped += 1
                                        mapping.unmapped = max(0, mapping.unmapped - 1)
                                        if candidate.display_match:
                                            mapping.verified += 1
                                        else:
                                            mapping.mismatched += 1
                                        result.verified_crosschecks = mapping.verified
                                        refresh_mapping_coverage_detail()
                                        if history is not None and run_id is not None:
                                            history.update_mapping_review(
                                                run_id,
                                                coverage=mapping,
                                                suggestions=result.mapping_suggestions,
                                                check_coverage=result.coverage,
                                            )
                                    render_stats()
                                    render_mapping_stats()
                                    confirmed_prefix = (
                                        f"{candidate.source_member}:"
                                        if candidate.source_member != "primary"
                                        else ""
                                    )
                                    ui.notify(
                                        f"Mapped {suggestion.figure_raw} to "
                                        f"{confirmed_prefix}{candidate.sheet}!"
                                        f"{candidate.cell}"
                                    )
                                    render_suggestions()

                                return confirm

                            button = ui.button(
                                "Confirm", on_click=confirm_handler()
                            ).classes("ghostbtn").props("no-caps flat dense")
                            if result.profile_name == "default":
                                button.disable()
                            if not mutable:
                                button.disable()

    render_suggestions()


def _result_from_record(record: RunRecord) -> QCRunResult:
    """Rehydrate a stored run so completed work renders like a fresh result."""
    return QCRunResult(
        profile_name=record.profile,
        mode=record.mode,
        files=record.files,
        findings=record.findings,
        disclosures=record.disclosures,
        coverage=record.coverage,
        mapping_coverage=record.mapping_coverage,
        mapping_suggestions=record.mapping_suggestions,
        verified_crosschecks=record.verified_crosschecks,
        comparison_scope=record.comparison_scope,
        package_manifest=record.package_manifest,
        alignment_trust=record.alignment_trust,
        # Stored counts keep `.counts` O(1); a lazy sequence would otherwise
        # revalidate every finding just to draw the header stats.
        severity_counts={
            severity: record.counts.get(severity.value, 0)
            for severity in Severity
        },
    )


def _focus_actions(
    service: "FocusService | None", record: RunRecord
) -> Callable[[Finding], None] | None:
    """Per-finding desktop actions, or ``None`` when focus is not available."""
    if service is None or not service.available:
        return None
    client_id = str(ui.context.client.id)

    def render(finding: Finding) -> None:
        _render_focus_actions(service, record, finding, client_id)

    return render


def _render_focus_actions(
    service: "FocusService",
    record: RunRecord,
    finding: Finding,
    client_id: str,
) -> None:
    """Explicit per-role Bind and Focus actions for one atomic finding."""
    seeds = service.seeds(record, finding.finding_id)
    if not seeds:
        return
    container = ui.element("div").classes("focusrow")

    def paint() -> None:
        container.clear()
        # Rerendering a detail retires every token issued under the old revision.
        service.new_revision(client_id)
        with container, ui.row().classes("items-center gap-2 mt-2 flex-wrap"):
            ui.label("Open in desktop Office").classes("dk")
            for seed in sorted(
                seeds,
                key=lambda item: (item.role.value, item.member_id),
            ):
                _render_role_action(service, record, finding, client_id, seed, paint)

    paint()


def _render_role_action(
    service: "FocusService",
    record: RunRecord,
    finding: Finding,
    client_id: str,
    seed: FocusTargetSeed,
    repaint: Callable[[], None],
) -> None:
    role = seed.role
    label = FOCUS_ROLE_LABELS[role]
    if seed.member_id != "primary":
        label = f"{label} ({seed.member_id})"
    if service.binding(client_id, record.run_id, role, seed.member_id) is None:

        async def bind() -> None:
            report = await service.bind(
                client_id,
                record,
                role,
                seed.member_id,
            )
            if not report.offered:
                ui.notify(
                    f"Cannot bind the {label} document: "
                    f"{report.outcome.value.replace('_', ' ')}",
                    type="warning",
                )
                return
            _confirm_binding(service, record, client_id, report, repaint)

        ui.button(f"Bind {label}", on_click=bind).classes("ghostbtn").props(
            "no-caps flat dense"
        )
        return

    token = service.issue_token(
        client_id,
        record.run_id,
        finding.finding_id,
        role,
        member_id=seed.member_id,
    )

    async def focus() -> None:
        claim = service.consume_token(client_id, token)
        if isinstance(claim, TokenRejection):
            ui.notify(
                "This action expired; reopen the finding and try again.",
                type="warning",
            )
            repaint()
            return
        reply = await service.focus(client_id, record, claim)
        if reply.outcome is FocusOutcome.FOCUSED:
            ui.notify(f"Opened the {label} document", type="positive")
        elif reply.outcome is FocusOutcome.FOCUSED_WITHOUT_FOREGROUND:
            ui.notify(
                f"Opened the {label} document; Windows kept the current window "
                "in front.",
                type="info",
            )
        else:
            ui.notify(
                f"No desktop action was taken: {reply.code.replace('_', ' ')}",
                type="warning",
            )
        repaint()

    ui.button(f"Focus {label}", on_click=focus).classes("primarybtn").props(
        "no-caps flat dense"
    )


def _confirm_binding(
    service: "FocusService",
    record: RunRecord,
    client_id: str,
    report: BindReport,
    repaint: Callable[[], None],
) -> None:
    """Explicit confirmation of the sole byte-identical open document."""
    label = FOCUS_ROLE_LABELS[report.role]
    if report.member_id != "primary":
        label = f"{label} ({report.member_id})"
    needs_acknowledgement = not service.acknowledged(client_id)
    with ui.dialog() as dialog, ui.card().classes("w-[34rem] max-w-full"):
        ui.label("Bind this byte-identical open document").classes("runhead")
        ui.label(
            f"{record.files.get(report.role_key, label)} — {label}"
            + (f" · folder {report.folder_label}" if report.folder_label else "")
        ).classes("mono")
        ui.label(
            "Exactly one open document has saved bytes identical to this run's "
            "input. This is not proof that it is the file you originally "
            "uploaded, only that its saved bytes match."
        ).classes("note")
        if report.unsaved_changes:
            ui.label(
                "That document has unsaved changes. Its saved bytes still match, "
                "but anything edited since the last save is not reflected here."
            ).classes("notecard")
        accepted = {"value": not needs_acknowledgement}
        if needs_acknowledgement:
            def on_ack(event: events.ValueChangeEventArguments) -> None:
                accepted["value"] = bool(event.value)

            ui.checkbox(
                "I understand that changing selection can trigger Office "
                "add-ins or event handlers in this document.",
                on_change=on_ack,
            )
        with ui.row().classes("items-center gap-2"):

            def confirm() -> None:
                if not accepted["value"]:
                    ui.notify(
                        "Acknowledge the Office side-effect note first.",
                        type="warning",
                    )
                    return
                service.acknowledge(client_id)
                outcome = service.confirm(
                    client_id,
                    record,
                    report.role,
                    report.member_id,
                )
                dialog.close()
                if outcome is not BindOutcome.MATCHED:
                    ui.notify("That candidate is no longer available.", type="warning")
                repaint()

            ui.button("Confirm binding", on_click=confirm).classes("primarybtn").props(
                "no-caps flat dense"
            )
            ui.button("Cancel", on_click=dialog.close).classes("ghostbtn").props(
                "no-caps flat dense"
            )
    dialog.open()


#: Shown when a legacy run's locator sidecar exceeds the safe decode limit.
_FOCUS_DEGRADED_NOTE = (
    "Desktop focus is unavailable for this run: its locator sidecar was "
    "recorded before block storage and is too large to decode safely. "
    "Everything else about the run is intact; re-QC the run to restore "
    "desktop focus."
)


def _render_completed_run(
    container: ui.element,
    work_dir: Path,
    run_id: int,
    *,
    focus_service: "FocusService | None" = None,
) -> None:
    history = RunHistory(work_dir / "history.sqlite3")
    try:
        record = history.get_run(run_id)
    except KeyError:
        ui.notify(f"Run #{run_id} is no longer in the history store", type="negative")
        return
    delta: FindingsDelta | None = None
    if record.rerun_of is not None:
        delta, delta_note = _rerun_delta(history, record)
        if delta_note:
            ui.notify(delta_note)
    container.clear()
    with container:
        if record.focus_degraded():
            ui.label(_FOCUS_DEGRADED_NOTE).classes("notecard")
        _render_result_view(
            _result_from_record(record),
            {kind: Path(path) for kind, path in record.report_paths.items()},
            f"Results — run #{run_id}",
            delta=delta,
            rerun_of=record.rerun_of,
            history=history,
            run_id=run_id,
            profiles_dir=work_dir / "profiles",
            focus_actions=_focus_actions(focus_service, record),
            export_root=work_dir / "runs",
        )
    # Results render below the upload/config sections; bring them into view.
    ui.run_javascript(
        f"document.getElementById('{container.html_id}')"
        "?.scrollIntoView({behavior: 'smooth', block: 'start'})"
    )


def create_pages(
    work_dir: Path,
    *,
    port: int = 8080,
    network_mode: NetworkMode = NetworkMode.LOCAL,
    expires_at: dt.datetime | None = None,
    desktop_focus: bool = False,
) -> None:
    secure_managed_tree(work_dir)
    uploads_dir = work_dir / "uploads"
    profiles_dir = work_dir / "profiles"
    private_directory(uploads_dir)
    private_directory(profiles_dir)
    # Windows-only, loopback-only, explicit opt-in; off unless all three hold.
    focus_service = FocusService(
        work_dir, enabled=desktop_focus, network_mode=network_mode
    )

    def open_app_settings() -> None:
        with ui.dialog() as dialog, ui.card().classes("w-[36rem] max-w-[94vw]"):
            ui.label("Local app settings").classes("runhead")
            ui.label("Local storage · active").classes("runhead")
            ui.label(
                "Uploads, profiles, run history, reports, and server configuration "
                "use the QC Tool data directory on Linux and Windows. Choose a "
                "different directory when starting QC Tool with --data-dir."
            ).classes("note")
            if focus_service.platform != "win32":
                ui.label(
                    "Desktop Office focus and Desktop shortcut controls are "
                    "Windows-only. Local storage and the complete browser review "
                    "workflow remain available on Linux."
                ).classes("notecard")
            else:
                persisted_focus = load_server_config(work_dir).desktop_focus
                focus_label = (
                    "On"
                    if persisted_focus
                    else "On for this launch"
                    if focus_service.enabled
                    else "Off"
                )
                ui.label(f"Desktop Office focus · {focus_label}").classes("runhead")
                ui.label(
                    "When enabled, QC Tool can enumerate already-open Excel and "
                    "PowerPoint documents for an exact-byte Bind → Confirm → Focus "
                    "action. It never opens, saves, recalculates, or closes a document."
                ).classes("note")
                if network_mode is NetworkMode.LAN:
                    ui.label(
                        "Focus is unavailable while network access is enabled, even "
                        "when this preference is remembered."
                    ).classes("notecard")

                if focus_service.enabled:

                    def disable_focus() -> None:
                        _set_desktop_focus_preference(work_dir, focus_service, False)
                        dialog.close()
                        ui.notify("Desktop Office focus disabled and bindings cleared")

                    ui.button(
                        "Disable Desktop Office focus",
                        on_click=disable_focus,
                    ).classes("ghostbtn").props("flat no-caps")
                else:

                    def request_focus_consent() -> None:
                        accepted = {"value": False}
                        with ui.dialog() as consent, ui.card().classes(
                            "w-[34rem] max-w-full"
                        ):
                            ui.label("Enable Desktop Office focus?").classes("runhead")
                            ui.label(
                                "QC Tool will inspect the identity and saved bytes of "
                                "already-open Office documents. Changing selection can "
                                "trigger document add-ins or event handlers. Every document "
                                "still requires an explicit byte-identical binding."
                            ).classes("notecard")

                            def acknowledge(event: events.ValueChangeEventArguments) -> None:
                                accepted["value"] = bool(event.value)

                            ui.checkbox(
                                "I understand and want to remember this setting on this "
                                "Windows account.",
                                on_change=acknowledge,
                            )

                            def confirm() -> None:
                                if not accepted["value"]:
                                    ui.notify("Acknowledge the desktop action note first")
                                    return
                                _set_desktop_focus_preference(
                                    work_dir,
                                    focus_service,
                                    True,
                                )
                                consent.close()
                                dialog.close()
                                ui.notify("Desktop Office focus enabled")

                            with ui.row().classes("items-center gap-2"):
                                ui.button("Enable focus", on_click=confirm).classes(
                                    "runbtn"
                                ).props("no-caps")
                                ui.button("Cancel", on_click=consent.close).props(
                                    "flat no-caps"
                                )
                        consent.open()

                    enable_button = ui.button(
                        "Enable Desktop Office focus",
                        on_click=request_focus_consent,
                    ).classes("runbtn").props("no-caps")
                    if network_mode is NetworkMode.LAN:
                        enable_button.disable()

                ui.separator()
                try:
                    shortcut = shortcut_status(work_dir, port=port)
                    ui.label(f"Desktop shortcut · {shortcut.state.value}").classes(
                        "runhead"
                    )

                    def create_shortcut() -> None:
                        try:
                            result = install_shortcut(work_dir, port=port)
                        except Exception:
                            logger.exception("desktop-shortcut-install-failed")
                            ui.notify("Desktop shortcut could not be created", type="negative")
                            return
                        dialog.close()
                        ui.notify(f"Desktop shortcut {result.state.value}")

                    def delete_shortcut() -> None:
                        try:
                            remove_shortcut(work_dir, port=port)
                        except Exception:
                            logger.exception("desktop-shortcut-remove-failed")
                            ui.notify("Desktop shortcut could not be removed", type="negative")
                            return
                        dialog.close()
                        ui.notify("Desktop shortcut removed")

                    if shortcut.state is ShortcutState.INSTALLED:
                        ui.button("Remove desktop shortcut", on_click=delete_shortcut).classes(
                            "ghostbtn"
                        ).props("flat no-caps")
                    else:
                        label = (
                            "Repair desktop shortcut"
                            if shortcut.state is ShortcutState.STALE
                            else "Create desktop shortcut"
                        )
                        ui.button(label, on_click=create_shortcut).classes(
                            "ghostbtn"
                        ).props("flat no-caps")
                    ui.label(
                        f"Quiet-launch diagnostics are written to "
                        f"{launcher_log_path(work_dir).relative_to(work_dir)}."
                    ).classes("note")
                except Exception:
                    logger.exception("desktop-shortcut-status-failed")
                    ui.label("Desktop shortcut status is unavailable.").classes("notecard")
            ui.button("Close", on_click=dialog.close).props("flat no-caps")
        dialog.open()

    def on_disconnect(client) -> None:
        focus_service.forget_client(str(client.id))
        RunHistory(work_dir / "history.sqlite3").pause_review_sessions()

    app.on_disconnect(on_disconnect)
    # One process-global FIFO manager: refreshes reconnect, tabs never duplicate
    # work, and requests left over from a previous server run are orphaned.
    queue_manager = get_manager(work_dir)
    if not queue_manager.shutdown_hook_installed:
        queue_manager.shutdown_hook_installed = True
        app.on_shutdown(queue_manager.shutdown)

    def request_shutdown() -> None:
        """Confirm, then stop the local server exactly as Ctrl+C would."""
        pending = queue_manager.store.pending()
        RunHistory(work_dir / "history.sqlite3").pause_review_sessions()
        with ui.dialog() as dialog, ui.card().classes("w-[32rem] max-w-full"):
            ui.label("Stop the QC Tool server?").classes("runhead")
            ui.label(
                "The local server stops and this page goes offline. Recorded "
                "runs, history, profiles, and exported reports are unaffected."
            ).classes("note")
            if pending:
                names = ", ".join(f"#{record.request_id[:8]}" for record in pending)
                ui.label(
                    f"{len(pending)} queued or running QC job(s) will stop and be "
                    f"recorded as unfinished: {names}. Submit them again after "
                    "restarting."
                ).classes("notecard")
            if network_mode is NetworkMode.LAN:
                ui.label(
                    "Network access is enabled and unauthenticated, so anyone "
                    "who can reach this page can stop the server."
                ).classes("notecard")

            def confirm() -> None:
                dialog.close()
                # Replace the client first so the analyst sees a terminal
                # message instead of NiceGUI's reconnect overlay.
                ui.run_javascript(_STOPPED_PAGE_JS)
                ui.timer(0.4, app.shutdown, once=True)

            with ui.row().classes("items-center gap-2"):
                ui.button("Stop server", on_click=confirm).classes("runbtn").props(
                    "no-caps"
                )
                ui.button("Keep running", on_click=dialog.close).props("flat no-caps")
        dialog.open()

    def maybe_prompt_storage_cleanup(*, offer_history_nav: bool) -> None:
        """One dialog per threshold crossing; silence until ~10% further growth."""
        history = RunHistory(work_dir / "history.sqlite3")
        history.backfill_storage_bytes(limit=50)
        total, archived_bytes, _ = history.storage_summary()
        dismissed = int(app.storage.general.get("storage_prompt_bytes", 0) or 0)
        if not _storage_prompt_due(total, dismissed_at_bytes=dismissed):
            return
        with ui.dialog() as dialog, ui.card().classes("w-[34rem] max-w-full"):
            ui.label("Run history is getting large").classes("runhead")
            ui.label(
                f"Stored runs use {_format_bytes(total)}"
                + (
                    f" ({_format_bytes(archived_bytes)} of that is archived)"
                    if archived_bytes
                    else ""
                )
                + ". Export the runs you want to keep, then archive or delete "
                "the rest from the history list. Deleting removes stored "
                "reports and evidence; source deliverables are never touched."
            ).classes("notecard")

            def dismiss() -> None:
                app.storage.general["storage_prompt_bytes"] = total
                dialog.close()

            with ui.row().classes("items-center gap-2"):
                if offer_history_nav:
                    ui.button(
                        "Open run history",
                        on_click=lambda: ui.navigate.to("/history"),
                    ).classes("runbtn").props("no-caps")
                ui.button("Not now", on_click=dismiss).props("flat no-caps")
        dialog.open()

    @ui.page("/")
    def main_page(rerun: int | None = None) -> None:  # pyright: ignore[reportUnusedFunction]
        state = SessionState(
            mode=_initial_mode(app.storage.general.get("qc_mode")),
        )
        file_states: dict[str, ui.label] = {}
        dynamic_member_boxes: dict[str, ui.element] = {}
        add_member_buttons: dict[str, ui.button] = {}
        maybe_prompt_storage_cleanup(offer_history_nav=True)

        rerun_record = None
        if rerun is not None:
            history = RunHistory(work_dir / "history.sqlite3")
            try:
                rerun_record = history.get_run(rerun)
                state.rerun_of = rerun
                state.mode = _initial_mode(
                    app.storage.general.get("qc_mode"),
                    rerun_mode=rerun_record.mode,
                )
            except KeyError:
                rerun_record = None

        with page_frame(
            "run",
            network_mode=network_mode.value,
            expires_at=expires_at.isoformat(timespec="seconds") if expires_at else None,
            on_settings=open_app_settings,
            on_shutdown=request_shutdown,
        ):
            ui.label("Compare deliverables").classes("pagetitle")
            ui.label(
                "Run intrinsic current-file checks, compare reporting cycles, "
                "or reconcile the final Excel and PowerPoint package. Files "
                "are processed locally and never modified."
            ).classes("lede")
            rerun_banner_actions: ui.row | None = None
            if rerun_record is not None:
                with ui.element("div").classes("rerunbanner"), ui.row().classes(
                    "items-center gap-3 w-full no-wrap"
                ):
                    ui.label(
                        f"Re-QC of run #{rerun_record.run_id} — files are"
                        " re-selected only after verifying they still match that"
                        " run. Anything missing or changed (e.g. renamed) must be"
                        " selected again below."
                    ).classes("flex-1")
                    rerun_banner_actions = ui.row().classes("no-wrap")

            section("1 · QC mode")
            ui.link("How to choose a mode →", "/guide#modes").classes("guide-jump")
            mode_copy = ui.label().classes("modecopy")
            role_group_boxes: dict[str, ui.element] = {}
            role_badges: dict[str, ui.label] = {}
            uploaders: dict[str, ui.upload] = {}

            def update_mode_surface() -> None:
                try:
                    state.package_manifest = PackageManifest.from_role_files(
                        state.files
                    )
                except ValueError:
                    state.package_manifest = None
                descriptions = {
                    QCRunMode.CURRENT_FILE_PREFLIGHT: (
                        "Intrinsic integrity checks on the latest Excel workbook "
                        "and/or PowerPoint deck."
                    ),
                    QCRunMode.CYCLE_COMPARISON: (
                        "Baseline-versus-current checks with expected cadence growth."
                    ),
                    QCRunMode.FINAL_PACKAGE: (
                        "Current Excel and PowerPoint preflight plus figure reconciliation."
                    ),
                }
                mode_copy.text = descriptions[state.mode]
                baseline_box = role_group_boxes.get("baseline")
                if baseline_box is not None:
                    baseline_box.visible = state.mode is QCRunMode.CYCLE_COMPARISON
                for role, badge in role_badges.items():
                    badge.text = _role_requirement(state.mode, role)
                    badge.classes(
                        add="satisfied" if role in state.files else "",
                        remove="" if role in state.files else "satisfied",
                    )

            def on_mode_change(e: events.ValueChangeEventArguments) -> None:
                state.mode = QCRunMode(e.value)
                if state.rerun_of is None:
                    app.storage.general["qc_mode"] = state.mode.value
                update_mode_surface()
                refresh_readiness()

            mode_select = ui.toggle(
                {mode.value: label for mode, label in MODE_LABELS.items()},
                value=state.mode.value,
                on_change=on_mode_change,
            ).classes("mode-select").props("no-caps spread")
            if rerun_record is not None:
                mode_select.disable()
            update_mode_surface()

            section("2 · Inputs")
            ui.label(
                "Baseline is the cycle you compare against. Current is the cycle "
                "you are signing off. Files are never swapped for you."
            ).classes("note")

            async def save_dynamic_upload(
                event: events.UploadEventArguments,
                role: str,
            ) -> None:
                try:
                    filename = _safe_upload_name(event.file.name)
                except ValueError as exc:
                    ui.notify(str(exc), type="negative")
                    return
                target = uploads_dir / role.replace(":", "-") / filename
                private_directory(target.parent)
                generation = state.upload_generations.get(role, 0) + 1
                state.upload_generations[role] = generation
                await event.file.save(target)
                private_file(target)
                size = event.file.size()
                state.files[role] = target
                state.file_sizes[role] = size
                state.file_hashes.pop(role, None)
                file_states[role].text = (
                    f"{filename} · {max(1, size // 1024):,} KB · verifying bytes"
                )
                file_states[role].classes(add="ok", remove="err")
                update_mode_surface()
                refresh_scope()
                refresh_readiness()
                digest = await asyncio.to_thread(sha256_file, target)
                if (
                    state.upload_generations.get(role) != generation
                    or state.files.get(role) != target
                ):
                    return
                state.file_hashes[role] = digest
                file_states[role].text = (
                    f"{filename} · {max(1, size // 1024):,} KB"
                )
                refresh_readiness()

            def clear_dynamic_file(role: str) -> None:
                state.upload_generations[role] = (
                    state.upload_generations.get(role, 0) + 1
                )
                state.files.pop(role, None)
                state.file_sizes.pop(role, None)
                state.file_hashes.pop(role, None)
                state.passwords.pop(role, None)

            def move_button_handler(
                mover: Callable[[int], None],
                offset: int,
            ) -> Callable[[], None]:
                def handle() -> None:
                    mover(offset)

                return handle

            def add_member_handler(side: str) -> Callable[[], None]:
                def handle() -> None:
                    open_add_member(side)

                return handle

            def remove_member_handler(
                role: str,
                member_id: str,
                side: str,
            ) -> Callable[[], None]:
                def handle() -> None:
                    clear_dynamic_file(role)
                    state.member_order[side].remove(member_id)
                    state.selected_member_sheets.pop(member_id, None)
                    state.available_member_sheets.pop(member_id, None)
                    render_dynamic_members(side)
                    update_mode_surface()
                    refresh_scope()
                    refresh_password_inputs()
                    refresh_readiness()

                return handle

            def render_dynamic_members(side: str) -> None:
                container = dynamic_member_boxes.get(side)
                if container is None:
                    return
                container.clear()
                member_ids = state.member_order[side]
                with container:
                    for index, member_id in enumerate(member_ids):
                        role = f"{side}_excel:{member_id}"

                        async def handle_member_upload(
                            event: events.UploadEventArguments,
                            role: str = role,
                        ) -> None:
                            await save_dynamic_upload(event, role)

                        def move_member(
                            offset: int,
                            member_id: str = member_id,
                            side: str = side,
                        ) -> None:
                            order = state.member_order[side]
                            old_index = order.index(member_id)
                            new_index = max(0, min(len(order) - 1, old_index + offset))
                            if old_index == new_index:
                                return
                            order.insert(new_index, order.pop(old_index))
                            render_dynamic_members(side)

                        with ui.element("div").classes("additional-workbook"):
                            with ui.row().classes("items-center gap-2 w-full"):
                                ui.label(member_id).classes("rolelabel mono")
                                ui.label("Excel workbook").classes("rolebadge")
                                up = ui.button(
                                    icon="keyboard_arrow_up",
                                    on_click=move_button_handler(move_member, -1),
                                ).props("flat round dense aria-label='Move workbook up'")
                                up.set_enabled(index > 0)
                                up.tooltip("Move up")
                                down = ui.button(
                                    icon="keyboard_arrow_down",
                                    on_click=move_button_handler(move_member, 1),
                                ).props(
                                    "flat round dense aria-label='Move workbook down'"
                                )
                                down.set_enabled(index < len(member_ids) - 1)
                                down.tooltip("Move down")
                                ui.button(
                                    icon="delete_outline",
                                    on_click=remove_member_handler(
                                        role,
                                        member_id,
                                        side,
                                    ),
                                ).props(
                                    "flat round dense aria-label='Remove workbook member'"
                                ).tooltip("Remove member")
                            uploaders[role] = ui.upload(
                                label=ROLE_ARTIFACT["current_excel"],
                                auto_upload=True,
                                max_file_size=MAX_UPLOAD_BYTES,
                                on_upload=handle_member_upload,
                            ).props(
                                f'accept="{ROLE_ACCEPT["current_excel"]}" flat'
                            )
                            selected = state.files.get(role)
                            file_states[role] = ui.label(
                                selected.name if selected is not None else "no file selected"
                            ).classes("filestate")
                            if selected is not None:
                                file_states[role].classes(add="ok")
                button = add_member_buttons.get(side)
                if button is not None:
                    button.set_enabled(len(member_ids) + 1 < MAX_WORKBOOKS_PER_SIDE)

            def open_add_member(side: str) -> None:
                if len(state.member_order[side]) + 1 >= MAX_WORKBOOKS_PER_SIDE:
                    ui.notify("The eight-workbook limit is reached", type="warning")
                    return
                with ui.dialog() as dialog, ui.card().classes("w-[28rem] max-w-full"):
                    ui.label(f"Add {side} workbook").classes("runhead")
                    member_input = ui.input(
                        "Stable member ID",
                        placeholder="e.g. operations",
                    ).classes("w-full").props("outlined dense")
                    ui.label(
                        "Use the same ID on baseline and current when they are "
                        "the same logical workbook."
                    ).classes("note")

                    def confirm_add() -> None:
                        member_id = str(member_input.value or "").strip()
                        if member_id == "primary":
                            ui.notify("primary uses the main upload above", type="warning")
                            return
                        if re.fullmatch(MEMBER_ID_PATTERN, member_id) is None:
                            ui.notify(
                                "Use a lowercase ID beginning with a letter; "
                                "letters, numbers, dashes and underscores only.",
                                type="warning",
                            )
                            return
                        if member_id in state.member_order[side]:
                            ui.notify("That member ID already exists", type="warning")
                            return
                        state.member_order[side].append(member_id)
                        dialog.close()
                        render_dynamic_members(side)
                        update_mode_surface()
                        refresh_password_inputs()
                        refresh_readiness()

                    with ui.row().classes("items-center gap-2"):
                        ui.button("Add", on_click=confirm_add).classes(
                            "runbtn"
                        ).props("no-caps")
                        ui.button("Cancel", on_click=dialog.close).props(
                            "flat no-caps"
                        )
                dialog.open()

            with ui.element("div").classes("rolegroups"):
                for group, title, subtitle, roles in ROLE_GROUPS:
                    with ui.element("div").classes(
                        f"rolegroup rolegroup-{group}"
                    ) as group_box:
                        role_group_boxes[group] = group_box
                        with ui.row().classes("rolegrouphead"):
                            ui.label(title).classes("rolegrouptitle")
                            ui.label(subtitle).classes("rolegroupnote")
                        with ui.element("div").classes("upgrid"):
                            for role in roles:

                                async def handle_upload(
                                    e: events.UploadEventArguments, role: str = role
                                ) -> None:
                                    try:
                                        filename = _safe_upload_name(e.file.name)
                                    except ValueError as exc:
                                        ui.notify(str(exc), type="negative")
                                        return
                                    target = uploads_dir / role / filename
                                    private_directory(target.parent)
                                    generation = state.upload_generations.get(role, 0) + 1
                                    state.upload_generations[role] = generation
                                    await e.file.save(target)
                                    private_file(target)
                                    state.files[role] = target
                                    size = e.file.size()
                                    state.file_sizes[role] = size
                                    state.file_hashes.pop(role, None)
                                    file_states[role].text = (
                                        f"{e.file.name} · {max(1, size // 1024):,} KB · "
                                        "verifying bytes"
                                    )
                                    file_states[role].classes(add="ok", remove="err")
                                    update_mode_surface()
                                    refresh_scope()
                                    refresh_readiness()
                                    digest = await asyncio.to_thread(sha256_file, target)
                                    if (
                                        state.upload_generations.get(role) != generation
                                        or state.files.get(role) != target
                                    ):
                                        return
                                    state.file_hashes[role] = digest
                                    file_states[role].text = (
                                        f"{e.file.name} · {max(1, size // 1024):,} KB"
                                    )
                                    refresh_readiness()

                                def clear_role(role: str = role) -> None:
                                    state.upload_generations[role] = (
                                        state.upload_generations.get(role, 0) + 1
                                    )
                                    state.files.pop(role, None)
                                    state.file_sizes.pop(role, None)
                                    state.file_hashes.pop(role, None)
                                    uploaders[role].reset()
                                    file_states[role].text = "no file selected"
                                    file_states[role].classes(remove="ok err")
                                    update_mode_surface()
                                    refresh_scope()
                                    refresh_readiness()

                                with ui.column().classes("rolerow"):
                                    with ui.row().classes("rolehead"):
                                        ui.label(ROLE_ARTIFACT[role]).classes("rolelabel")
                                        role_badges[role] = ui.label().classes(
                                            "rolebadge"
                                        )
                                    uploaders[role] = ui.upload(
                                        label=ROLE_ACCEPT[role],
                                        auto_upload=True,
                                        max_file_size=MAX_UPLOAD_BYTES,
                                        on_upload=handle_upload,
                                    ).props(f'accept="{ROLE_ACCEPT[role]}" flat')
                                    with ui.row().classes("rolefoot"):
                                        file_states[role] = ui.label(
                                            "no file selected"
                                        ).classes("filestate")
                                        ui.button(
                                            "Clear", on_click=clear_role
                                        ).classes("linkbtn").props(
                                            f"flat dense no-caps aria-label='Clear {role}'"
                                        )
                        with ui.element("div").classes(
                            "additional-workbooks w-full"
                        ) as member_box:
                            dynamic_member_boxes[group] = member_box
                        add_member_buttons[group] = ui.button(
                            "Add workbook",
                            icon="add",
                            on_click=add_member_handler(group),
                        ).classes("ghostbtn").props("flat no-caps dense")
                        add_member_buttons[group].tooltip(
                            f"Add another {group} Excel workbook"
                        )
                        render_dynamic_members(group)
            update_mode_surface()

            if rerun_record is not None:
                for role in rerun_record.file_paths:
                    prefix, separator, member_id = role.partition(":")
                    if not separator or prefix not in {
                        "baseline_excel",
                        "current_excel",
                    }:
                        continue
                    side = prefix.removesuffix("_excel")
                    if member_id not in state.member_order[side]:
                        state.member_order[side].append(member_id)
                for side in ("baseline", "current"):
                    render_dynamic_members(side)
                state.rerun_required = frozenset(rerun_record.file_paths)
                for role, stored in rerun_record.file_paths.items():
                    if role not in file_states:
                        continue
                    stored_path = Path(stored)
                    display_name = rerun_record.files.get(role, stored_path.name)
                    if not stored_path.exists():
                        file_states[role].text = (
                            f"file not found: {display_name} — select the file again"
                        )
                        file_states[role].classes(add="err")
                        continue
                    expected_hash = rerun_record.file_hashes.get(role)
                    if expected_hash and sha256_file(stored_path) != expected_hash:
                        file_states[role].text = (
                            f"{stored_path.name} changed since run "
                            f"#{rerun_record.run_id} — select the file again"
                        )
                        file_states[role].classes(add="err")
                        continue
                    state.files[role] = stored_path
                    if expected_hash:
                        state.file_hashes[role] = expected_hash
                    file_states[role].text = (
                        f"{stored_path.name} · verified, reused from run "
                        f"#{rerun_record.run_id}"
                    )
                    file_states[role].classes(add="ok")

            section("3 · Run policy")
            ui.link("Profiles and controls guide →", "/guide#profiles").classes(
                "guide-jump"
            )
            profile_options = list_profiles(profiles_dir)
            initial_profile = "default"
            if rerun_record is not None and rerun_record.profile in profile_options:
                initial_profile = rerun_record.profile
                state.profile_name = initial_profile

            def manage_profiles() -> None:
                """Profile authoring lives here so routine setup stays a selector."""
                with ui.dialog() as dialog, ui.card().classes(
                    "w-[72rem] max-w-[96vw] max-h-[92vh] overflow-y-auto"
                ) as card:
                    ui.label("Manage profiles").classes("runhead")
                    ui.label(
                        "Profiles hold controls, waivers, mappings, and cadence "
                        "rules. The built-in default profile cannot be edited."
                    ).classes("note")
                    editor_controller: ProfileEditorController | None = None
                    with ui.row().classes("items-end gap-2 w-full no-wrap"):
                        new_profile_name = (
                            ui.input("New profile name")
                            .classes("flex-1")
                            .props("outlined dense")
                        )

                        def create_profile() -> None:
                            name = (new_profile_name.value or "").strip()
                            if not name or name == "default":
                                ui.notify(
                                    "Choose a non-empty, non-default name",
                                    type="warning",
                                )
                                return
                            try:
                                path = _profile_path(profiles_dir, name)
                            except ValueError as exc:
                                ui.notify(str(exc), type="warning")
                                return
                            if path.exists():
                                ui.notify(
                                    f"Profile {name!r} already exists",
                                    type="warning",
                                )
                                return
                            save_profile(new_profile(name), path)
                            options = list_profiles(profiles_dir)
                            profile_select.options = options
                            profile_select.value = name
                            profile_select.update()
                            state.profile_name = name
                            editing.options = options
                            editing.update()
                            if editor_controller is not None:
                                editor_controller.request_load(name)
                            refresh_readiness()
                            ui.notify(f"Profile {name!r} created")

                        ui.button("Create", on_click=create_profile).classes(
                            "ghostbtn"
                        ).props("no-caps flat")
                    editing = (
                        ui.select(
                            list_profiles(profiles_dir),
                            value=state.profile_name,
                            label="Edit profile",
                        )
                        .classes("w-full")
                        .props("outlined dense")
                    )

                    def profile_saved(old_name: str, new_name: str) -> None:
                        options = list_profiles(profiles_dir)
                        editing.options = options
                        editing.update()
                        profile_select.options = options
                        if state.profile_name == old_name:
                            state.profile_name = new_name
                            profile_select.value = new_name
                        profile_select.update()
                        refresh_readiness()

                    def close_profile_dialog() -> None:
                        dialog.close()

                    editor_controller = open_profile_editor(
                        card,
                        profiles_dir,
                        editing,
                        on_close=close_profile_dialog,
                        selected_files=lambda: dict(state.files),
                        selected_passwords=lambda: dict(state.passwords),
                        on_saved=profile_saved,
                    )
                    dialog.on("hide", editor_controller.dispose)
                dialog.open()

            def on_profile_change(e: events.ValueChangeEventArguments) -> None:
                state.profile_name = str(e.value)
                refresh_readiness()

            with ui.row().classes("policyrow"):
                profile_select = (
                    ui.select(
                        profile_options,
                        value=initial_profile,
                        label="Deliverable profile",
                        on_change=on_profile_change,
                    )
                    .classes("w-64")
                    .props("outlined dense")
                )
                ui.button("Manage profiles", on_click=manage_profiles).classes(
                    "ghostbtn"
                ).props("no-caps flat")

                def on_acceptance_absolute(e: events.ValueChangeEventArguments) -> None:
                    state.acceptance_absolute = float(e.value or 0.0)
                    refresh_readiness()

                def on_acceptance_percent(e: events.ValueChangeEventArguments) -> None:
                    state.acceptance_percent = float(e.value or 0.0)
                    refresh_readiness()

                def on_allow_large(e: events.ValueChangeEventArguments) -> None:
                    state.allow_large_workbooks = bool(e.value)
                    refresh_readiness()

                ui.number(
                    label="Accept ± value",
                    value=0,
                    min=0,
                    step=0.01,
                    on_change=on_acceptance_absolute,
                ).props('dense outlined placeholder="0 = off"').classes("w-32").tooltip(
                    "Numeric differences within this absolute value report as "
                    "within-tolerance Info in cycle comparisons — visible, "
                    "never hidden"
                )
                ui.number(
                    label="Accept ± %",
                    value=0,
                    min=0,
                    step=0.05,
                    on_change=on_acceptance_percent,
                ).props('dense outlined placeholder="e.g. 0.1"').classes("w-32").tooltip(
                    "Cycle-comparison percentage bound (either bound accepts); "
                    "0 keeps the strict default"
                )
                ui.checkbox(
                    "Override workbook workload refusals",
                    value=False,
                    on_change=on_allow_large,
                ).tooltip(
                    "Continue past workbook size or formula-link safety limits "
                    "for every workbook in this run. Use only after reading the "
                    "refusal reason and confirming enough local memory; workload "
                    "coverage will be degraded."
                )

            with (
                ui.expansion(
                    "Comparison scope — optional; everything is compared by default"
                ).classes("w-full mt-4"),
                ui.column().classes("gap-2 p-2"),
            ):
                scope_hint = ui.label(
                    "Upload files to list their sheets and slides."
                ).classes("note")
                scope_box = ui.element("div").classes("w-full")

            def refresh_scope() -> None:
                try:
                    selected_files = _files_for_mode(state.mode, state.files)
                except ValueError:
                    selected_files = {}
                excel_roles = sorted(
                    role
                    for role in selected_files
                    if role == "current_excel" or role.startswith("current_excel:")
                )
                if not excel_roles:
                    excel_roles = sorted(
                        role
                        for role in selected_files
                        if role == "baseline_excel"
                        or role.startswith("baseline_excel:")
                    )
                member_paths: dict[str, Path] = {}
                for role in excel_roles:
                    _prefix, _separator, member_id = role.partition(":")
                    member_paths[member_id or "primary"] = selected_files[role]
                ppt_path = state.files.get("current_ppt")
                state.available_member_sheets = {
                    member_id: peek_sheet_names(path)
                    for member_id, path in member_paths.items()
                }
                if set(member_paths) == {"primary"}:
                    state.available_sheets = state.available_member_sheets["primary"]
                    state.available_member_sheets = {}
                else:
                    state.available_sheets = []
                state.available_slides = (
                    peek_slide_titles(ppt_path) if ppt_path else []
                )
                state.selected_sheets &= set(state.available_sheets)
                for member_id in list(state.selected_member_sheets):
                    if member_id not in state.available_member_sheets:
                        state.selected_member_sheets.pop(member_id, None)
                        continue
                    state.selected_member_sheets[member_id] &= set(
                        state.available_member_sheets[member_id]
                    )
                state.selected_slides &= {
                    index for index, _ in state.available_slides
                }
                scope_hint.visible = not (
                    state.available_sheets
                    or state.available_member_sheets
                    or state.available_slides
                )
                scope_box.clear()
                with scope_box:
                    if state.available_sheets:
                        ui.label(
                            "Excel sheets (none checked = compare all)"
                        ).classes("rolelabel")
                        with ui.row().classes("gap-2 flex-wrap"):
                            for name in state.available_sheets:

                                def toggle_sheet(
                                    e: events.ValueChangeEventArguments,
                                    name: str = name,
                                ) -> None:
                                    if e.value:
                                        state.selected_sheets.add(name)
                                    else:
                                        state.selected_sheets.discard(name)
                                    refresh_readiness()

                                ui.checkbox(
                                    name,
                                    value=name in state.selected_sheets,
                                    on_change=toggle_sheet,
                                ).props("dense")
                    if state.available_member_sheets:
                        ui.label(
                            "Excel sheets by workbook member "
                            "(none checked = compare all)"
                        ).classes("rolelabel")
                        for member_id, sheet_names in sorted(
                            state.available_member_sheets.items()
                        ):
                            selected = state.selected_member_sheets.setdefault(
                                member_id,
                                set(),
                            )
                            display_name = next(
                                (
                                    path.name
                                    for role, path in selected_files.items()
                                    if (
                                        role == "current_excel"
                                        and member_id == "primary"
                                    )
                                    or (
                                        role.startswith("current_excel:")
                                        and role.endswith(f":{member_id}")
                                    )
                                ),
                                member_id,
                            )
                            ui.label(
                                f"{member_id} · {display_name}"
                            ).classes("runmeta")
                            with ui.row().classes("gap-2 flex-wrap"):
                                for name in sheet_names:

                                    def toggle_member_sheet(
                                        event: events.ValueChangeEventArguments,
                                        member_id: str = member_id,
                                        name: str = name,
                                    ) -> None:
                                        member_selection = (
                                            state.selected_member_sheets.setdefault(
                                                member_id,
                                                set(),
                                            )
                                        )
                                        if event.value:
                                            member_selection.add(name)
                                        else:
                                            member_selection.discard(name)
                                        refresh_readiness()

                                    ui.checkbox(
                                        name,
                                        value=name in selected,
                                        on_change=toggle_member_sheet,
                                    ).props("dense")
                    if state.available_slides:
                        ui.label(
                            "PPT slides (none checked = compare all)"
                        ).classes("rolelabel")
                        with ui.row().classes("gap-2 flex-wrap"):
                            for index, title in state.available_slides:

                                def toggle_slide(
                                    e: events.ValueChangeEventArguments,
                                    index: int = index,
                                ) -> None:
                                    if e.value:
                                        state.selected_slides.add(index)
                                    else:
                                        state.selected_slides.discard(index)
                                    refresh_readiness()

                                ui.checkbox(
                                    f"{index}: {title[:32]}",
                                    value=index in state.selected_slides,
                                    on_change=toggle_slide,
                                ).props("dense")

            refresh_scope()

            with (
                ui.expansion("Passwords — only needed for encrypted files").classes(
                    "w-full mt-4"
                ),
                ui.row().classes("gap-4 flex-wrap p-2") as password_box,
            ):
                pass

            def refresh_password_inputs() -> None:
                password_box.clear()
                dynamic_roles = [
                    f"{side}_excel:{member_id}"
                    for side in ("baseline", "current")
                    for member_id in state.member_order[side]
                ]
                with password_box:
                    for role in (*ROLES, *dynamic_roles):
                        def update_password(
                            event: events.ValueChangeEventArguments,
                            role: str = role,
                        ) -> None:
                            state.passwords[role] = str(event.value or "")

                        ui.input(
                            role_label(role),
                            value=state.passwords.get(role, ""),
                            password=True,
                            password_toggle_button=True,
                            on_change=update_password,
                        ).classes("w-72").props("outlined dense")

            refresh_password_inputs()

            # Readiness bar: mode, supplied roles, policy, blockers, and Run QC
            # stay visible while the analyst scrolls the rest of the setup.
            extra_run_buttons: list[ui.button] = []
            with ui.element("div").classes("readybar"):
                with ui.element("div").classes("readysummary"):
                    ready_line = ui.label().classes("r1")
                    ready_detail = ui.label().classes("r2")
                with ui.row().classes("readyactions"):
                    spinner = ui.spinner(size="1.6rem")
                    spinner.visible = False
                    run_button = (
                        # late-bound: start_run is defined below; NiceGUI
                        # schedules the returned coroutine
                        ui.button("Run QC", on_click=lambda: start_run())
                        .classes("runbtn")
                        .props("no-caps")
                    )
                    run_button.tooltip(
                        "Runs execute one at a time; extra submissions queue and "
                        "survive a browser refresh"
                    )
                queue_row = ui.row().classes("readyqueue")
                queue_row.visible = False
                with queue_row:
                    progress_label = ui.label("").classes("hint")
                    queue_actions = ui.row().classes("no-wrap items-center gap-2")
                    completed_links = ui.row().classes("no-wrap items-center gap-2")

                def refresh_readiness() -> None:
                    blockers = _run_blockers(
                        state.mode,
                        state.files,
                        file_hashes=state.file_hashes,
                        rerun_of=state.rerun_of,
                        rerun_required=state.rerun_required,
                    )
                    supplied = [role_label(role) for role in sorted(state.files)]
                    ready_line.set_text(
                        f"{MODE_LABELS[state.mode]} — "
                        + ("cannot run yet" if blockers else "ready to run")
                    )
                    ready_line.classes(
                        add="blocked" if blockers else "",
                        remove="" if blockers else "blocked",
                    )
                    details = [
                        "files: " + (", ".join(supplied) or "none selected"),
                        f"profile: {state.profile_name}",
                        _scope_summary(state),
                        _acceptance_summary(state),
                    ]
                    if state.allow_large_workbooks:
                        details.append("workbook workload refusals overridden")
                    cautions = _input_cautions(state)
                    if cautions:
                        details.insert(0, "check inputs — " + "; ".join(cautions))
                    if blockers:
                        details.insert(0, "blocked — " + "; ".join(blockers))
                    ready_detail.set_text(" · ".join(details))
                    ready_detail.classes(
                        add="caution" if cautions and not blockers else "",
                        remove="" if cautions and not blockers else "caution",
                    )
                    run_button.set_enabled(not blockers)
                    for button in extra_run_buttons:
                        button.set_enabled(not blockers)

                queue_signature: dict[str, tuple[tuple[str, str], ...]] = {"value": ()}
                own_requests: set[str] = set()
                # Requests this page observed while active, including ones a
                # different tab submitted, so a refresh still reports the outcome.
                watched_requests: set[str] = set()

                def announce(record: RunStateRecord) -> None:
                    if record.status is RunStatus.SUCCEEDED and record.run_id:
                        ui.notify(f"Run #{record.run_id} complete")
                        if record.request_id in own_requests:
                            _render_completed_run(
                                results,
                                work_dir,
                                record.run_id,
                                focus_service=focus_service,
                            )
                        else:
                            with completed_links:
                                ui.link(
                                    f"Run #{record.run_id} complete — open results",
                                    f"/runs/{record.run_id}",
                                ).classes("guide-jump")
                    elif record.status is RunStatus.CANCELLED:
                        ui.notify(
                            "Run cancelled; no successful run was recorded",
                            type="warning",
                        )
                    elif record.status is RunStatus.ORPHANED:
                        ui.notify(
                            "The server stopped before this run finished; "
                            "submit it again",
                            type="warning",
                        )
                    else:
                        ui.notify(
                            f"QC run failed: {record.error or 'unknown error'}",
                            type="negative",
                        )

                def refresh_queue() -> None:
                    pending = queue_manager.store.pending()
                    lines = [_queue_status_line(record) for record in pending]
                    progress_label.set_text(" || ".join(lines))
                    queue_row.visible = bool(lines) or bool(completed_links.default_slot.children)
                    signature = tuple(
                        (record.request_id, record.status.value) for record in pending
                    )
                    if signature != queue_signature["value"]:
                        queue_signature["value"] = signature
                        queue_actions.clear()
                        with queue_actions:
                            for record in pending:
                                ui.button(
                                    f"Cancel #{record.request_id[:8]}",
                                    icon="stop_circle",
                                    on_click=_cancel_handler(
                                        queue_manager, record.request_id
                                    ),
                                ).classes("ghostbtn").props("flat no-caps dense")
                    watched_requests.update(record.request_id for record in pending)
                    spinner.visible = any(
                        record.request_id in own_requests for record in pending
                    )
                    for request_id in sorted(watched_requests):
                        record = queue_manager.store.get(request_id)
                        if record is None or record.is_active:
                            continue
                        watched_requests.discard(request_id)
                        announce(record)
                        own_requests.discard(request_id)

                ui.timer(0.5, refresh_queue)

                async def start_run() -> None:
                    blockers = _run_blockers(
                        state.mode,
                        state.files,
                        file_hashes=state.file_hashes,
                        rerun_of=state.rerun_of,
                        rerun_required=state.rerun_required,
                    )
                    if blockers:
                        ui.notify("; ".join(blockers), type="warning")
                        return
                    RunHistory(work_dir / "history.sqlite3").pause_review_sessions()
                    files = _files_for_mode(state.mode, state.files)
                    try:
                        profile = load_profile_by_name(profiles_dir, state.profile_name)
                    except Exception as exc:
                        ui.notify(f"Profile failed to load: {exc}", type="negative")
                        return
                    try:
                        manifest = PackageManifest.from_role_files(files)
                    except ValueError as exc:
                        ui.notify(str(exc), type="negative")
                        return
                    state.package_manifest = manifest

                    # Projected diff volume: warn before a monster comparison
                    # starts and offer scoping to the sheets that changed.
                    if (
                        state.mode is QCRunMode.CYCLE_COMPARISON
                        and "baseline_excel" in files
                        and "current_excel" in files
                        and not state.selected_sheets
                    ):
                        projection = await asyncio.to_thread(
                            project_cycle_volume,
                            files["baseline_excel"],
                            files["current_excel"],
                        )
                        if (
                            projection is not None
                            and projection.projected_max_findings
                            > REPORT_DEFER_FINDINGS
                        ):
                            _open_projection_dialog(projection, files, profile, manifest)
                            return
                    submit_run(files, profile, manifest)

                def _open_projection_dialog(
                    projection: VolumeProjection,
                    files: dict[str, Path],
                    profile: DeliverableProfile,
                    manifest: PackageManifest,
                ) -> None:
                    minutes = max(1, projection.projected_max_findings // 20_000)
                    picked: set[str] = set()
                    boxes: dict[str, ui.checkbox] = {}
                    with ui.dialog() as dialog, ui.card().classes(
                        "w-[42rem] max-w-full"
                    ):
                        ui.label("This looks like a very large comparison").classes(
                            "runhead"
                        )
                        ui.label(
                            f"Up to ~{projection.projected_max_findings:,} findings "
                            f"across {len(projection.changed_sheets)} changed sheet"
                            f"{'s' if len(projection.changed_sheets) != 1 else ''}"
                            + (
                                f" (+{len(projection.added_sheets)} added, "
                                f"{len(projection.removed_sheets)} removed)"
                                if projection.added_sheets
                                or projection.removed_sheets
                                else ""
                            )
                            + f"; {len(projection.identical_sheets)} sheets are "
                            "byte-identical and produce nothing."
                        ).classes("lede")
                        ui.label(
                            f"A full run can take up to roughly {minutes} minute"
                            f"{'s' if minutes != 1 else ''} and its review queue "
                            "reads from recorded summaries. Pick the sheets to "
                            "compare now, or run everything — the full run "
                            "remains the sign-off artifact."
                        ).classes("note")

                        def sync_picked() -> None:
                            run_selected.set_enabled(bool(picked))
                            run_selected.set_text(
                                f"Run {len(picked)} selected sheet"
                                f"{'s' if len(picked) != 1 else ''}"
                                if picked
                                else "Run selected sheets"
                            )

                        def toggle(name: str, value: bool) -> None:
                            if value:
                                picked.add(name)
                            else:
                                picked.discard(name)
                            sync_picked()

                        def set_all(value: bool) -> None:
                            for box in boxes.values():
                                box.set_value(value)

                        with ui.row().classes("items-center gap-2"):
                            ui.label("Sheets that differ").classes("dk")
                            ui.button("All", on_click=lambda: set_all(True)).props(
                                "flat dense no-caps"
                            )
                            ui.button("None", on_click=lambda: set_all(False)).props(
                                "flat dense no-caps"
                            )
                        with ui.element("div").classes(
                            "max-h-64 overflow-y-auto w-full"
                        ):
                            for name, volume in projection.sheet_volumes:
                                boxes[name] = ui.checkbox(
                                    f"{name} — up to ~{volume:,} findings",
                                    value=False,
                                    on_change=lambda e, name=name: toggle(
                                        name, bool(e.value)
                                    ),
                                ).props("dense")
                        if projection.removed_sheets:
                            ui.label(
                                "Removed sheets (reported regardless of scope): "
                                + ", ".join(sorted(projection.removed_sheets))
                            ).classes("hint")

                        def run_picked() -> None:
                            if not picked:
                                return
                            state.selected_sheets = set(picked)
                            dialog.close()
                            submit_run(files, profile, manifest)

                        def run_everything() -> None:
                            dialog.close()
                            submit_run(files, profile, manifest)

                        with ui.row().classes("items-center gap-2"):
                            run_selected = ui.button(
                                "Run selected sheets", on_click=run_picked
                            ).classes("runbtn").props("no-caps")
                            run_selected.set_enabled(False)
                            ui.button(
                                "Run everything", on_click=run_everything
                            ).classes("ghostbtn").props("no-caps flat")
                            ui.button("Cancel", on_click=dialog.close).props(
                                "flat no-caps"
                            )
                    dialog.open()

                def submit_run(
                    files: dict[str, Path],
                    profile: DeliverableProfile,
                    manifest: PackageManifest,
                ) -> None:
                    request = RunRequest(
                        request_id=new_request_id(),
                        work_dir=str(work_dir),
                        mode=state.mode.value,
                        profile_name=profile.name,
                        profile=profile.model_dump(mode="json"),
                        files={role: str(path) for role, path in files.items()},
                        display_files={role: path.name for role, path in files.items()},
                        package_manifest=manifest.model_dump(mode="json"),
                        compare_member_sheets={
                            key: tuple(sorted(value))
                            for key, value in state.selected_member_sheets.items()
                            if value
                        },
                        allow_large_workbooks=state.allow_large_workbooks,
                        acceptance_absolute=max(state.acceptance_absolute, 0.0),
                        acceptance_relative=max(state.acceptance_percent, 0.0) / 100.0,
                        compare_sheets=(
                            tuple(sorted(state.selected_sheets))
                            if state.selected_sheets
                            and set(state.available_sheets) != state.selected_sheets
                            else ()
                        ),
                        compare_slides=(
                            tuple(sorted(state.selected_slides))
                            if state.selected_slides
                            and {i for i, _ in state.available_slides}
                            != state.selected_slides
                            else ()
                        ),
                        rerun_of=state.rerun_of,
                    )
                    credentials = {
                        role: password
                        for role, password in state.passwords.items()
                        if role in files and password
                    }
                    try:
                        record = queue_manager.submit(request, credentials)
                    except QueueBusyError as exc:
                        ui.notify(str(exc), type="warning")
                        return
                    own_requests.add(request.request_id)
                    watched_requests.add(request.request_id)
                    ui.notify(
                        "Run started"
                        if record.queue_position == 0
                        else f"Run queued at position {record.queue_position}"
                    )
                    refresh_queue()

            if rerun_banner_actions is not None:
                with rerun_banner_actions:
                    extra_run_buttons.append(
                        ui.button("Run QC now", on_click=start_run)
                        .classes("runbtn")
                        .props("no-caps dense")
                    )

            update_mode_surface()
            refresh_readiness()

            results = ui.column().classes("w-full")

    @ui.page("/guide")
    def guide_page() -> None:  # pyright: ignore[reportUnusedFunction]
        with page_frame(
            "guide",
            network_mode=network_mode.value,
            expires_at=expires_at.isoformat(timespec="seconds") if expires_at else None,
            on_settings=open_app_settings,
            on_shutdown=request_shutdown,
            colophon="Curated by Himanshu",
        ):
            render_guide()

    @ui.page("/history")
    def history_page() -> None:  # pyright: ignore[reportUnusedFunction]
        with page_frame(
            "history",
            network_mode=network_mode.value,
            expires_at=expires_at.isoformat(timespec="seconds") if expires_at else None,
            on_settings=open_app_settings,
            on_shutdown=request_shutdown,
        ):
            ui.label("Run history").classes("pagetitle")
            history = RunHistory(work_dir / "history.sqlite3")
            history.backfill_storage_bytes()
            runs = history.list_runs()
            if not runs:
                ui.label("No runs recorded yet.").classes("lede")
                return
            records = {record.run_id: record for record in runs}
            rows = [_history_row(record) for record in runs]
            selected: set[int] = set()

            with ui.row().classes("historytoolbar"):
                search = (
                    ui.input(placeholder="search run id, profile, or file")
                    .classes("flex-1 min-w-56")
                    .props("outlined dense clearable")
                )
                mode_filter = (
                    ui.select(
                        {"": "All modes", **{m.value: t for m, t in MODE_LABELS.items()}},
                        value="",
                        label="Mode",
                    )
                    .classes("w-52")
                    .props("outlined dense")
                )
                profile_filter = (
                    ui.select(
                        {"": "All profiles", **{r.profile: r.profile for r in runs}},
                        value="",
                        label="Profile",
                    )
                    .classes("w-44")
                    .props("outlined dense")
                )
                capability_filter = (
                    ui.select(
                        {
                            "": "Any capability",
                            "limited": "Capability limited",
                            "complete": "All checks ran",
                        },
                        value="",
                        label="Capability",
                    )
                    .classes("w-48")
                    .props("outlined dense")
                )
                date_filter = (
                    ui.select(
                        {"": "Any date", "1": "Last 24h", "7": "Last 7d", "30": "Last 30d"},
                        value="",
                        label="Date",
                    )
                    .classes("w-36")
                    .props("outlined dense")
                )
                archive_filter = (
                    ui.select(
                        {
                            "active": "Active runs",
                            "archived": "Archived runs",
                            "all": "Active and archived",
                        },
                        value="active",
                        label="Shelf",
                    )
                    .classes("w-44")
                    .props("outlined dense")
                )
                ui.button(
                    "Select all shown", on_click=lambda: select_all_shown()
                ).classes("ghostbtn").props("no-caps flat dense")

            storage_note = ui.label().classes("note")

            def refresh_storage_note() -> None:
                total, archived_bytes, _ = history.storage_summary()
                text = (
                    f"History uses {_format_bytes(total)} across {len(runs)} stored runs"
                    " · Run IDs are permanent and may have gaps after deletion"
                )
                if archived_bytes:
                    text += f" · {_format_bytes(archived_bytes)} archived"
                storage_note.text = text

            refresh_storage_note()
            maybe_prompt_storage_cleanup(offer_history_nav=False)

            bulk_bar = ui.row().classes("bulkbar")
            bulk_bar.visible = False
            with bulk_bar:
                bulk_label = ui.label().classes("bulkcount")
                ui.button(
                    "Export selected (.zip)", on_click=lambda: export_selected()
                ).classes("ghostbtn").props("no-caps flat dense")
                ui.button("Archive", on_click=lambda: set_archive(True)).classes(
                    "ghostbtn"
                ).props("no-caps flat dense")
                ui.button("Restore", on_click=lambda: set_archive(False)).classes(
                    "ghostbtn"
                ).props("no-caps flat dense")
                ui.button("Delete", on_click=lambda: confirm_delete()).classes(
                    "ghostbtn dangerbtn"
                ).props("no-caps flat dense")
                ui.button("Clear selection", on_click=lambda: clear_selection()).classes(
                    "linkbtn"
                ).props("no-caps flat dense")

            table = (
                ui.table(
                    columns=_HISTORY_COLUMNS,
                    rows=rows,
                    row_key="id",
                    pagination=20,
                )
                .classes("findings-table history-table")
                .props("flat dense")
            )
            table.add_slot("body", HISTORY_BODY_SLOT)
            empty_note = ui.label("No run matches these filters.").classes("note")
            empty_note.visible = False
            with ui.expansion("Assurance trends").classes("w-full mt-4"):
                ui.label(
                    "Fixed local counts only. Review minutes come only from the "
                    "explicit Start/Pause timer; unknown means no timer evidence."
                ).classes("note")
                trend_table = ui.table(
                    columns=[
                        {"name": "run", "label": "Run", "field": "run"},
                        {"name": "profile", "label": "Profile", "field": "profile"},
                        {"name": "atomics", "label": "Atomics", "field": "atomics"},
                        {"name": "decisions", "label": "Decisions", "field": "decisions"},
                        {"name": "limited", "label": "Limits", "field": "limited"},
                        {"name": "mapped", "label": "Mapped", "field": "mapped"},
                        {"name": "verified", "label": "Verified", "field": "verified"},
                        {"name": "exact_reusable", "label": "Reusable", "field": "exact_reusable"},
                        {"name": "carried", "label": "Carried", "field": "carried"},
                        {"name": "finalized", "label": "Finalized", "field": "finalized"},
                        {
                            "name": "review_minutes",
                            "label": "Review min",
                            "field": "review_minutes",
                        },
                    ],
                    rows=[],
                    row_key="run",
                    pagination=20,
                ).classes("findings-table").props("flat dense")

            def refresh_bulk_bar() -> None:
                bulk_bar.visible = bool(selected)
                bulk_label.set_text(
                    f"{len(selected)} run{'s' if len(selected) != 1 else ''} selected"
                )

            def refilter() -> None:
                needle = str(search.value or "").strip().lower()
                days = str(date_filter.value or "")
                shelf = str(archive_filter.value or "active")
                cutoff = (
                    (dt.datetime.now(dt.UTC) - dt.timedelta(days=int(days))).isoformat()
                    if days
                    else None
                )
                visible = [
                    row
                    for row in rows
                    if (
                        not needle
                        or needle
                        in f"#{row['id']} {row['profile']} {row['files']}".lower()
                    )
                    and (not mode_filter.value or row["mode_key"] == mode_filter.value)
                    and (
                        not profile_filter.value
                        or row["profile"] == profile_filter.value
                    )
                    and (
                        not capability_filter.value
                        or row["capability"] == capability_filter.value
                    )
                    and (cutoff is None or str(row["started"]) >= cutoff)
                    and (
                        shelf == "all"
                        or (shelf == "archived") == bool(row["archived"])
                    )
                ]
                for row in rows:
                    row["sel"] = row["id"] in selected
                table.rows = visible
                trend_table.rows = [
                    _history_trend_row(records[int(str(row["id"]))], history)
                    for row in visible
                    if int(str(row["id"])) in records
                ]
                empty_note.visible = not visible

            for control in (
                search,
                mode_filter,
                profile_filter,
                capability_filter,
                date_filter,
                archive_filter,
            ):
                control.on_value_change(refilter)

            def clear_selection() -> None:
                selected.clear()
                refresh_bulk_bar()
                refilter()

            def reload_rows() -> None:
                """Re-read the store so archive and delete results are truthful."""
                nonlocal runs, rows
                runs = history.list_runs()
                records.clear()
                records.update({record.run_id: record for record in runs})
                rows = [_history_row(record) for record in runs]
                selected.intersection_update(records)
                refresh_storage_note()
                refresh_bulk_bar()
                refilter()

            def set_archive(archived: bool) -> None:
                if not selected:
                    return
                changed = history.set_archived(selected, archived)
                reload_rows()
                ui.notify(
                    f"{changed} run{'s' if changed != 1 else ''} "
                    + ("archived" if archived else "restored")
                )

            def confirm_delete() -> None:
                if not selected:
                    return
                targets = sorted(selected)
                with ui.dialog() as dialog, ui.card().classes("w-[32rem] max-w-full"):
                    ui.label(
                        f"Delete {len(targets)} run"
                        f"{'s' if len(targets) != 1 else ''} permanently?"
                    ).classes("runhead")
                    ui.label(
                        "The history records, analyst annotations, and stored "
                        "report files are removed. Source deliverables are never "
                        "touched. Archiving keeps the evidence instead."
                    ).classes("notecard")
                    ui.label(
                        "runs: " + ", ".join(f"#{run_id}" for run_id in targets)
                    ).classes("runmeta")

                    def do_delete() -> None:
                        removed = history.delete_runs(
                            targets, managed_root=work_dir / "runs"
                        )
                        for run_id in targets:
                            focus_service.forget_run(run_id)
                        selected.clear()
                        reload_rows()
                        dialog.close()
                        ui.notify(
                            f"Deleted {removed} run{'s' if removed != 1 else ''}",
                            type="warning",
                        )

                    with ui.row().classes("items-center gap-2"):
                        ui.button("Delete permanently", on_click=do_delete).classes(
                            "runbtn"
                        ).props("no-caps")
                        ui.button("Cancel", on_click=dialog.close).props(
                            "flat no-caps"
                        )
                dialog.open()

            def export_selected() -> None:
                if not selected:
                    return
                chosen = [records[run_id] for run_id in sorted(selected) if run_id in records]
                stamp = dt.datetime.now(dt.UTC).strftime("%Y%m%dT%H%M%SZ")
                bundle = work_dir / "exports" / f"qc-runs-{stamp}.zip"
                try:
                    export_runs_archive(
                        chosen, bundle, managed_root=work_dir / "runs"
                    )
                except OSError as exc:
                    ui.notify(f"Export failed: {exc}", type="negative")
                    return
                ui.download(str(bundle))
                ui.notify(f"Exported {len(chosen)} run(s)")

            def on_select(e: events.GenericEventArguments) -> None:
                run_id = int(e.args["id"])
                if bool(e.args.get("value")):
                    selected.add(run_id)
                else:
                    selected.discard(run_id)
                refresh_bulk_bar()
                refilter()

            def on_select_page(e: events.GenericEventArguments) -> None:
                visible_ids = {int(row["id"]) for row in table.rows}
                if bool(e.args.get("value")):
                    selected.update(visible_ids)
                else:
                    selected.difference_update(visible_ids)
                refresh_bulk_bar()
                refilter()

            def select_all_shown() -> None:
                selected.update(int(row["id"]) for row in table.rows)
                refresh_bulk_bar()
                refilter()

            def on_open(e: events.GenericEventArguments) -> None:
                ui.navigate.to(f"/runs/{int(e.args['id'])}")

            def on_rerun(e: events.GenericEventArguments) -> None:
                ui.navigate.to(f"/?rerun={int(e.args['id'])}")

            def on_export(e: events.GenericEventArguments) -> None:
                record = records.get(int(e.args["id"]))
                paths = (
                    record.signoff.report_paths
                    if record is not None and record.signoff is not None
                    else record.report_paths
                    if record is not None
                    else {}
                )
                path = (
                    paths.get(str(e.args["kind"]))
                )
                if path and Path(path).exists():
                    ui.download(str(path))
                    return
                ui.notify("That export is no longer on disk", type="warning")

            table.on("select", on_select)
            table.on("open", on_open)
            table.on("rerun", on_rerun)
            table.on("export", on_export)
            refilter()

    @ui.page("/runs/{run_id}")
    def run_detail_page(run_id: int) -> None:  # pyright: ignore[reportUnusedFunction]
        with page_frame(
            "history",
            network_mode=network_mode.value,
            expires_at=expires_at.isoformat(timespec="seconds") if expires_at else None,
            on_settings=open_app_settings,
            on_shutdown=request_shutdown,
        ):
            history = RunHistory(work_dir / "history.sqlite3")
            try:
                record = history.get_run(run_id)
            except KeyError:
                section("Run not found")
                ui.label(
                    f"No QC run with id {run_id} exists in the history store."
                ).classes("lede")
                return
            result = _result_from_record(record)
            heading = (
                f"Run #{record.run_id} — "
                f"{record.started_at.isoformat(timespec='seconds')}"
            )
            delta = None
            delta_note = ""
            if record.rerun_of is not None:
                heading += f" — re-QC of #{record.rerun_of}"
                delta, delta_note = _rerun_delta(history, record)
            report_paths = {kind: Path(path) for kind, path in record.report_paths.items()}
            if record.focus_degraded():
                ui.label(_FOCUS_DEGRADED_NOTE).classes("notecard")
            if delta_note:
                ui.label(delta_note).classes("notecard")
            _render_result_view(
                result,
                report_paths,
                heading,
                delta=delta,
                rerun_of=record.rerun_of,
                history=history,
                run_id=record.run_id,
                profiles_dir=work_dir / "profiles",
                focus_actions=_focus_actions(focus_service, record),
                export_root=work_dir / "runs",
            )


def run_app(
    work_dir: Path,
    *,
    port: int = 8080,
    host: str = "127.0.0.1",
    network_mode: NetworkMode = NetworkMode.LOCAL,
    expires_at: dt.datetime | None = None,
    desktop_focus: bool = False,
    show: bool = True,
) -> None:
    authority = (
        claim_local_instance(work_dir, port)
        if network_mode is NetworkMode.LOCAL and host == "127.0.0.1"
        else None
    )
    try:
        create_pages(
            work_dir,
            port=port,
            network_mode=network_mode,
            expires_at=expires_at,
            desktop_focus=desktop_focus,
        )
        if network_mode is NetworkMode.LAN:
            if expires_at is None:
                raise ValueError("LAN mode requires an expiry timestamp")

            async def expire_network_access() -> None:
                while True:
                    remaining = (expires_at - dt.datetime.now(dt.UTC)).total_seconds()
                    if remaining <= 0:
                        _persist_expired_lan_config(work_dir)
                        logger.warning("temporary LAN exposure expired; shutting down server")
                        app.shutdown()
                        return
                    await asyncio.sleep(min(2.0, remaining))
                    if not lan_config_matches(work_dir, expires_at):
                        logger.warning(
                            "LAN exposure config changed or expired; shutting down server"
                        )
                        app.shutdown()
                        return

            app.on_startup(lambda: asyncio.create_task(expire_network_access()))
        ui.run(
            title="QC Tool",
            host=host,
            port=port,
            reload=False,
            show=show,
            storage_secret=_storage_secret(work_dir),
        )
    finally:
        if authority is not None:
            release_local_instance(authority)
