"""NiceGUI local web application for the QC tool.

The UI stays thin: uploads, profile selection, passwords, one button.
All comparison logic lives in `qc_tool.engine`; `qc_tool.run_service.perform_run`
wraps a run with report writing and history recording, and heavy runs execute
in one owned worker process managed by `qc_tool.runqueue`. Everything runs on
localhost; sources are read-only. The visual language lives in
`qc_tool.ui.theme`.
"""

import asyncio
import datetime as dt
import html
import logging
import re
import secrets
from collections.abc import Callable, Mapping
from dataclasses import dataclass, field
from pathlib import Path

import yaml
from nicegui import app, events, ui

from qc_tool.config.profile import (
    CrosscheckMapping,
    DeliverableProfile,
    default_profile,
    load_profile,
    save_profile,
)
from qc_tool.coverage import (
    CoverageItem,
    CoverageState,
    QCRunMode,
    capability_limited,
)
from qc_tool.crosscheck.trace import MappingSuggestion, SuggestedSource
from qc_tool.engine import FindingsDelta, QCRunResult, compare_findings
from qc_tool.findings import Finding, FindingClass, GridExcerpt, Severity
from qc_tool.focus.binding import BindOutcome
from qc_tool.focus.model import FocusTargetSeed
from qc_tool.focus.protocol import FocusOutcome
from qc_tool.focus.service import ROLE_LABELS as FOCUS_ROLE_LABELS
from qc_tool.focus.service import BindReport, FocusService, TokenRejection
from qc_tool.history.run_state import RunStateRecord, RunStatus
from qc_tool.history.store import RunHistory, RunRecord, export_runs_archive, sha256_file
from qc_tool.io.peek import peek_sheet_names, peek_slide_titles
from qc_tool.progress import RunPhase
from qc_tool.report.excel_report import write_excel_report
from qc_tool.report.html_report import write_html_report
from qc_tool.review import (
    ReviewGroup,
    apply_group_review,
    build_pattern_groups,
    count_pattern_groups,
    format_group_ranges,
    prioritize_review,
)
from qc_tool.run_service import RunArtifacts, perform_run
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
    local_config,
    save_server_config,
)
from qc_tool.story import build_stories
from qc_tool.ui.guide import render_guide
from qc_tool.ui.theme import (
    FINDINGS_BODY_SLOT,
    HISTORY_BODY_SLOT,
    REVIEW_GROUPS_BODY_SLOT,
    REVIEW_MEMBER_ROWS_SLOT,
    page_frame,
    section,
    status_chip,
)

logger = logging.getLogger(__name__)

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
_PROFILE_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9 ._-]{0,79}$")

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
    passwords: dict[str, str] = field(default_factory=dict)  # role -> password
    profile_name: str = "default"
    mode: QCRunMode = QCRunMode.CYCLE_COMPARISON
    allow_large_workbooks: bool = False
    acceptance_absolute: float = 0.0
    acceptance_percent: float = 0.0  # analyst-facing percent; engine gets a fraction
    selected_sheets: set[str] = field(default_factory=set)
    available_sheets: list[str] = field(default_factory=list)
    selected_slides: set[int] = field(default_factory=set)
    available_slides: list[tuple[int, str]] = field(default_factory=list)
    rerun_of: int | None = None
    rerun_required: frozenset[str] = frozenset()  # roles the previous run used


# --- pure helpers (unit-testable without a browser) -------------------------


def _safe_upload_name(raw_name: str) -> str:
    """Return a browser-supplied basename that cannot escape managed storage."""
    name = Path(raw_name.replace("\\", "/")).name.strip()
    if not name or name in {".", ".."}:
        raise ValueError("uploaded file has no usable filename")
    return name


def _profile_path(profiles_dir: Path, name: str) -> Path:
    """Resolve a validated profile name inside the managed profile directory."""
    if not _PROFILE_NAME_RE.fullmatch(name) or name in {".", ".."}:
        raise ValueError(
            "profile name must start with a letter or number and contain only "
            "letters, numbers, spaces, dots, dashes, or underscores"
        )
    return profiles_dir / f"{name}.yaml"


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


def list_profiles(profiles_dir: Path) -> list[str]:
    private_directory(profiles_dir)
    return ["default", *sorted(p.stem for p in profiles_dir.glob("*.yaml"))]


def load_profile_by_name(profiles_dir: Path, name: str) -> DeliverableProfile:
    if name == "default":
        return default_profile()
    return load_profile(_profile_path(profiles_dir, name))


def _files_for_mode(
    mode: QCRunMode, files: dict[str, Path]
) -> dict[str, Path]:
    """Validate and select only file roles used by one QC mode."""
    if mode is QCRunMode.CYCLE_COMPARISON:
        has_excel = {"baseline_excel", "current_excel"} <= files.keys()
        has_ppt = {"baseline_ppt", "current_ppt"} <= files.keys()
        if not has_excel and not has_ppt:
            raise ValueError(
                "Upload a baseline and current Excel pair and/or PowerPoint pair"
            )
        return {
            role: path
            for role, path in files.items()
            if (has_excel and "excel" in role) or (has_ppt and "ppt" in role)
        }
    if mode is QCRunMode.CURRENT_FILE_PREFLIGHT:
        selected = {
            role: path
            for role, path in files.items()
            if role in {"current_excel", "current_ppt"}
        }
        if not selected:
            raise ValueError("Upload a current Excel workbook and/or PowerPoint deck")
        return selected
    required = {"current_excel", "current_ppt"}
    if not required <= files.keys():
        raise ValueError("Upload both current Excel and current PowerPoint files")
    return {role: files[role] for role in sorted(required)}


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
    rerun_of: int | None = None,
    rerun_required: frozenset[str] = frozenset(),
) -> list[str]:
    """Everything preventing a run, listed before the analyst presses Run QC."""
    blockers: list[str] = []
    missing = rerun_required - files.keys()
    if rerun_of is not None and missing:
        names = ", ".join(ROLE_LABELS[role] for role in sorted(missing))
        blockers.append(
            f"Re-QC of run #{rerun_of} is blocked until these files are "
            f"selected again: {names}"
        )
    try:
        _files_for_mode(mode, files)
    except ValueError as exc:
        blockers.append(str(exc))
    return blockers


def _scope_summary(state: SessionState) -> str:
    """The comparison scope actually sent to the engine, in one phrase."""
    parts: list[str] = []
    if state.selected_sheets and set(state.available_sheets) != state.selected_sheets:
        parts.append(
            f"{len(state.selected_sheets)}/{len(state.available_sheets)} sheets"
        )
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
        ):
            cautions.append(
                f"baseline and current {artifact} look like the same file "
                f"({state.files[baseline].name}) — a clean result would prove nothing"
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


# --- pages -------------------------------------------------------------------


def _finding_row(finding: Finding) -> dict[str, object]:
    return {
        "id": finding.finding_id,
        "severity": (finding.severity or Severity.WARNING).value,
        "class": finding.finding_class.value,
        "where": finding.sheet or finding.slide or "",
        "location": finding.location or finding.baseline_location or "",
        "message": finding.message,
        "baseline": finding.baseline_value or "",
        "current": finding.current_value or "",
        "element": finding.element or "",
        "impacts": "; ".join(finding.impacts),
        "artifact": finding.artifact,
        "comment": finding.analyst_comment,
        "overridden": finding.severity_overridden,
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
        "bx": finding.baseline_excerpt.model_dump() if finding.baseline_excerpt else None,
        "cx": finding.current_excerpt.model_dump() if finding.current_excerpt else None,
    }


def _findings_rows(result: QCRunResult) -> list[dict[str, object]]:
    return [_finding_row(finding) for finding in result.findings]


def _review_group_rows(
    groups: list[ReviewGroup],
    rationales: dict[str, str] | None = None,
) -> list[dict[str, object]]:
    reasons = rationales or {}
    return [
        {
            "id": group.group_id,
            "severity": group.severity.value,
            "class": group.finding_class.value,
            "where": group.sheet or group.slide or "",
            "location": format_group_ranges(group),
            "bounds": group.bounding_range,
            "baseline": "; ".join(group.baseline_ranges),
            "members": group.member_count,
            "message": (
                group.members[0].message
                if group.member_count == 1
                else (
                    f"{group.member_count:,} contiguous "
                    f"{group.finding_class.value.replace('_', ' ')} findings"
                )
            ),
            "element": group.element,
            "why": reasons.get(group.group_id, ""),
            "sel": False,
            "cap_degraded": any(
                member.finding_class is FindingClass.FINDINGS_CAPPED
                for member in group.members
            ),
        }
        for group in groups
    ]


_FINDINGS_COLUMNS = [
    {"name": "expand", "label": "", "field": "expand"},
    {"name": "id", "label": "ID", "field": "id", "sortable": True, "classes": "mono"},
    {"name": "severity", "label": "Severity", "field": "severity", "sortable": True},
    {"name": "class", "label": "Class", "field": "class", "sortable": True, "classes": "mono"},
    {"name": "where", "label": "Sheet / Slide", "field": "where", "sortable": True},
    {"name": "location", "label": "Location", "field": "location", "classes": "mono"},
    {"name": "message", "label": "Message", "field": "message", "align": "left"},
]

_REVIEW_GROUP_COLUMNS = [
    {
        "name": "severity",
        "label": "Severity",
        "field": "severity",
        "sortable": True,
        "align": "left",
    },
    {
        "name": "class",
        "label": "Class",
        "field": "class",
        "sortable": True,
        "align": "left",
    },
    {
        "name": "where",
        "label": "Location",
        "field": "where",
        "sortable": True,
        "align": "left",
    },
    {"name": "members", "label": "#", "field": "members", "sortable": True},
    {"name": "message", "label": "Review item", "field": "message", "align": "left"},
]


_HISTORY_COLUMNS = [
    {"name": "select", "label": "", "field": "sel"},
    {"name": "id", "label": "Run", "field": "id", "sortable": True},
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
    {"name": "files", "label": "Files", "field": "files", "align": "left"},
    {"name": "actions", "label": "", "field": "actions"},
]


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
        "files": " · ".join(record.files.values()),
        "exports": [
            kind
            for kind, path in record.report_paths.items()
            if Path(path).exists()
        ],
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
        cells = "".join(
            f'<td class="hit">{esc(str(value))}</td>'
            if row_index == excerpt.hit_row and col_index == excerpt.hit_col
            else f"<td>{esc(str(value))}</td>"
            for col_index, value in enumerate(row)
        )
        rows.append(f"<tr><th>{esc(str(excerpt.rows[row_index]))}</th>{cells}</tr>")
    return (
        f'<div class="ctxblock"><div class="ctxlabel">{esc(label)}</div>'
        f'<table class="ctxgrid"><tbody><tr><th></th>{cols}</tr>'
        + "".join(rows)
        + "</tbody></table></div>"
    )


def _render_evidence_body(finding: Finding) -> None:
    """Typed evidence axes and context grids for one finding, in the caller's slot."""
    with ui.element("div").classes("detailgrid"):
        for key, value in _evidence_axes(finding):
            ui.label(key).classes("dk")
            ui.label(value).classes("dv")
    for excerpt, label in (
        (finding.baseline_excerpt, "baseline"),
        (finding.current_excerpt, "current"),
    ):
        if excerpt is not None:
            ui.html(_context_grid_html(excerpt, label))


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
) -> None:
    """Results workbench: a compact run header, then Review queue (default),
    Stories, Coverage, Atomic evidence, and Mapping review. Only the active
    view is mounted, so a high-volume run does not pay for hidden tables."""
    findings_by_id = {f.finding_id: f for f in result.findings}
    review_groups = build_pattern_groups(result.findings)
    stories = build_stories(result.findings)
    all_rows: list[dict[str, object]] = []  # atomic rows are built on demand
    story_scope: set[str] = set()  # group ids pinned by a story, empty = no pin

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

        def render_stats() -> None:
            counts = count_pattern_groups(review_groups)
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
            for kind, label, writer in (
                ("excel", "Export Excel (.xlsx)", write_excel_report),
                ("html", "Export HTML", write_html_report),
            ):
                path = report_paths.get(kind)
                if path is None:
                    continue

                def export(writer=writer, path=path) -> None:
                    # Regenerate from the current (annotated) state before download.
                    writer(result, path)
                    ui.download(str(path))

                ui.button(label, on_click=export).classes("ghostbtn").props(
                    "no-caps flat dense"
                )
            if run_id is not None:
                ui.button(
                    "Re-QC this run",
                    on_click=lambda: ui.navigate.to(f"/?rerun={run_id}"),
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
            with ui.row().classes("reviewtoolbar"):
                severity_filter = (
                    ui.select(
                        [s.value for s in Severity],
                        value=[s.value for s in Severity if s is not Severity.EXPECTED],
                        multiple=True,
                        label="Show severities",
                    )
                    .classes("w-72")
                    .props("outlined dense use-chips")
                )
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
            with ui.element("div").classes("reviewsplit"):
                group_table = (
                    ui.table(
                        columns=_REVIEW_GROUP_COLUMNS,
                        rows=[],
                        row_key="id",
                        pagination=25,
                    )
                    .classes("findings-table review-groups-table")
                    .props("flat dense")
                )
                group_table.add_slot("body", REVIEW_GROUPS_BODY_SLOT)
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

        with ui.tab_panel("atomic"):
            ui.label(
                "Every atomic finding behind the review queue, with analyst "
                "severity overrides and comments."
            ).classes("note")
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

        if result.mapping_coverage is not None:
            with ui.tab_panel("mapping"):
                _render_mapping_review(
                    result,
                    render_stats,
                    history=history,
                    run_id=run_id,
                    profiles_dir=profiles_dir,
                )

    def _visible_group_rows() -> list[dict[str, object]]:
        selected = set(severity_filter.value or [])
        needle = str(text_filter.value or "").strip().lower()
        prioritized = prioritize_review(review_groups, stories)
        rows = [
            row
            for row in _review_group_rows(
                [item.group for item in prioritized],
                {item.group.group_id: item.rationale for item in prioritized},
            )
            if row["severity"] in selected
            and (not story_scope or row["id"] in story_scope)
        ]
        if needle:
            rows = [
                row
                for row in rows
                if needle
                in f"{row['where']} {row['class']} {row['location']} {row['message']}".lower()
            ]
        return rows

    def refilter() -> None:
        group_table.rows = _visible_group_rows()
        if all_rows:
            selected = set(severity_filter.value or [])
            findings_table.rows = [
                row for row in all_rows if row["severity"] in selected
            ]

    def refresh_review_groups() -> None:
        nonlocal review_groups
        review_groups = build_pattern_groups(result.findings)
        refilter()

    def clear_story_scope() -> None:
        story_scope.clear()
        story_pin.visible = False
        refilter()

    story_pin.on("click", clear_story_scope)
    severity_filter.on_value_change(refilter)
    text_filter.on_value_change(refilter)

    def on_view_change(e: events.ValueChangeEventArguments) -> None:
        # Atomic rows cost real DOM, so they are built the first time they matter.
        if e.value == "atomic" and not all_rows:
            all_rows[:] = _findings_rows(result)
            refilter()

    panels.on_value_change(on_view_change)

    def render_detail(group_id: str) -> None:
        group = next(
            (item for item in review_groups if item.group_id == group_id), None
        )
        detail_box.clear()
        if group is None:
            return
        member = group.members[0]
        with detail_box:
            ui.label(group.group_id).classes("detail-id")
            ui.label(
                member.message
                if group.member_count == 1
                else f"{group.member_count:,} findings — {member.message}"
            ).classes("detail-msg")
            ui.label(
                f"{group.sheet or group.slide or ''} · {format_group_ranges(group)}"
            ).classes("runmeta")
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
            # A multi-member group must never focus an arbitrary representative;
            # its members are chosen explicitly in the affected-findings dialog.
            if focus_actions is not None and group.member_count == 1:
                focus_actions(member)
            with ui.row().classes("items-center gap-2 mt-2"):
                ui.button(
                    "View affected findings",
                    on_click=lambda: open_members(group_id),
                ).classes("ghostbtn").props("no-caps flat dense")
                ui.button(
                    "Review group",
                    on_click=lambda: open_group_review(group_id),
                ).classes("ghostbtn").props("no-caps flat dense")

    def on_select(e: events.GenericEventArguments) -> None:
        group_id = str(e.args.get("id", ""))
        for row in group_table.rows:
            row["sel"] = row["id"] == group_id
        group_table.update()
        render_detail(group_id)

    def open_members(group_id: str) -> None:
        group = next(
            (item for item in review_groups if item.group_id == group_id), None
        )
        if group is None:
            return
        page_size = 50
        page = {"index": 0}
        members_by_id = {member.finding_id: member for member in group.members}
        with ui.dialog() as dialog, ui.card().classes("memberscard"):
            with ui.row().classes("items-baseline gap-3 w-full"):
                ui.label(
                    f"{group.group_id} — {group.member_count:,} affected findings"
                ).classes("runhead")
                ui.label(
                    f"{group.sheet or group.slide or ''} · {format_group_ranges(group)}"
                ).classes("runmeta")
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
                member = members_by_id.get(finding_id)
                member_detail.clear()
                if member is None:
                    return
                for row in member_table.rows:
                    row["sel"] = row["id"] == finding_id
                member_table.update()
                with member_detail:
                    ui.label(member.finding_id).classes("detail-id")
                    ui.label(member.message).classes("detail-msg")
                    _render_evidence_body(member)
                    if focus_actions is not None:
                        focus_actions(member)

            def on_member_select(e: events.GenericEventArguments) -> None:
                show_member(str(e.args.get("id", "")))

            member_table.on("select", on_member_select)

            def refresh_page() -> None:
                start = page["index"] * page_size
                end = min(start + page_size, group.member_count)
                member_table.rows = [
                    _finding_row(member) for member in group.members[start:end]
                ]
                page_label.set_text(
                    f"Showing {start + 1:,}-{end:,} of {group.member_count:,}"
                )
                previous_button.set_enabled(page["index"] > 0)
                next_button.set_enabled(end < group.member_count)
                show_member(group.members[start].finding_id)

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
                updates = apply_group_review(
                    group,
                    severity=severity,
                    comment=str(comment_input.value or ""),
                    replace_existing=apply_mode.value == "replace",
                )
                if not updates:
                    ui.notify("No group review changes to apply", type="warning")
                    return
                if history is not None and run_id is not None:
                    history.set_annotations_bulk(
                        run_id,
                        [
                            (update.finding_id, update.severity, update.comment)
                            for update in updates
                        ],
                    )
                if all_rows:
                    all_rows[:] = _findings_rows(result)
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
        dialog.open()

    group_table.on("select", on_select)

    def on_severity(e: events.GenericEventArguments) -> None:
        finding_id, value = e.args["id"], e.args["value"]
        finding = findings_by_id.get(finding_id)
        if finding is None:
            return
        finding.severity = Severity(value)
        finding.severity_overridden = True
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
        ui.notify(f"{finding_id}: comment saved")

    findings_table.on("sev", on_severity)
    findings_table.on("note", on_note)
    refilter()


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
            for label, value in (
                ("eligible", mapping.eligible),
                ("mapped", mapping.mapped),
                ("verified", mapping.verified),
                ("mismatched", mapping.mismatched),
                ("unresolved", mapping.unresolved),
                ("unmapped", mapping.unmapped),
            ):
                with ui.column().classes("mappingstat"):
                    ui.label(str(value)).classes("n")
                    ui.label(label).classes("l")

    render_mapping_stats()
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
                            ui.label(f"{candidate.sheet}!{candidate.cell}").classes(
                                "candidate-ref"
                            )
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
                                    render_suggestions()
                                    ui.notify(
                                        f"Mapped {suggestion.figure_raw} to "
                                        f"{candidate.sheet}!{candidate.cell}"
                                    )

                                return confirm

                            button = ui.button(
                                "Confirm", on_click=confirm_handler()
                            ).classes("ghostbtn").props("no-caps flat dense")
                            if result.profile_name == "default":
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
            for seed in sorted(seeds, key=lambda item: item.role.value):
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
    if service.binding(client_id, record.run_id, role) is None:

        async def bind() -> None:
            report = await service.bind(client_id, record, role)
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

    token = service.issue_token(client_id, record.run_id, finding.finding_id, role)

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
    needs_acknowledgement = not service.acknowledged(client_id)
    with ui.dialog() as dialog, ui.card().classes("w-[34rem] max-w-full"):
        ui.label("Bind this byte-identical open document").classes("runhead")
        ui.label(
            f"{record.files.get(report.role.value, label)} — {label}"
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
                outcome = service.confirm(client_id, record, report.role)
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
        try:
            previous = history.get_run(record.rerun_of)
            delta = compare_findings(previous.findings, record.findings)
        except KeyError:
            delta = None
    container.clear()
    with container:
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
        )
    # Results render below the upload/config sections; bring them into view.
    ui.run_javascript(
        f"document.getElementById('{container.html_id}')"
        "?.scrollIntoView({behavior: 'smooth', block: 'start'})"
    )


def create_pages(
    work_dir: Path,
    *,
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
    app.on_disconnect(lambda client: focus_service.forget_client(str(client.id)))
    # One process-global FIFO manager: refreshes reconnect, tabs never duplicate
    # work, and requests left over from a previous server run are orphaned.
    queue_manager = get_manager(work_dir)
    if not queue_manager.shutdown_hook_installed:
        queue_manager.shutdown_hook_installed = True
        app.on_shutdown(queue_manager.shutdown)

    def request_shutdown() -> None:
        """Confirm, then stop the local server exactly as Ctrl+C would."""
        pending = queue_manager.store.pending()
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

    @ui.page("/")
    def main_page(rerun: int | None = None) -> None:  # pyright: ignore[reportUnusedFunction]
        state = SessionState()
        file_states: dict[str, ui.label] = {}

        rerun_record = None
        if rerun is not None:
            history = RunHistory(work_dir / "history.sqlite3")
            try:
                rerun_record = history.get_run(rerun)
                state.rerun_of = rerun
                state.mode = rerun_record.mode
            except KeyError:
                rerun_record = None

        with page_frame(
            "run",
            network_mode=network_mode.value,
            expires_at=expires_at.isoformat(timespec="seconds") if expires_at else None,
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
                                    await e.file.save(target)
                                    private_file(target)
                                    state.files[role] = target
                                    size = e.file.size()
                                    state.file_sizes[role] = size
                                    file_states[role].text = (
                                        f"{e.file.name} · {max(1, size // 1024):,} KB"
                                    )
                                    file_states[role].classes(add="ok", remove="err")
                                    update_mode_surface()
                                    refresh_scope()
                                    refresh_readiness()

                                def clear_role(role: str = role) -> None:
                                    state.files.pop(role, None)
                                    state.file_sizes.pop(role, None)
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
            update_mode_surface()

            if rerun_record is not None:
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
                with ui.dialog() as dialog, ui.card().classes("w-[42rem] max-w-full"):
                    ui.label("Manage profiles").classes("runhead")
                    ui.label(
                        "Profiles hold controls, waivers, mappings, and cadence "
                        "rules. The built-in default profile cannot be edited."
                    ).classes("note")
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
                            save_profile(DeliverableProfile(name=name), path)
                            options = list_profiles(profiles_dir)
                            profile_select.options = options
                            profile_select.value = name
                            profile_select.update()
                            editing.options = options
                            editing.value = name
                            editing.update()
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
                    editor = (
                        ui.textarea()
                        .classes("w-full")
                        .style("font-family: var(--font-mono); font-size: 0.8rem")
                        .props("rows=16 outlined")
                    )

                    def load_selected() -> None:
                        name = str(editing.value or "default")
                        if name == "default":
                            editor.value = ""
                            editor.disable()
                            return
                        editor.enable()
                        path = _profile_path(profiles_dir, name)
                        editor.value = (
                            path.read_text(encoding="utf-8") if path.exists() else ""
                        )

                    editing.on_value_change(load_selected)
                    load_selected()

                    def save() -> None:
                        name = str(editing.value or "default")
                        if name == "default":
                            ui.notify(
                                "Create a named profile before editing",
                                type="warning",
                            )
                            return
                        try:
                            profile = DeliverableProfile.model_validate(
                                yaml.safe_load(editor.value or "") or {}
                            )
                        except Exception as exc:  # surfaced to the analyst
                            ui.notify(f"Invalid profile: {exc}", type="negative")
                            return
                        save_profile(profile, _profile_path(profiles_dir, name))
                        ui.notify("Profile saved")

                    with ui.row().classes("items-center gap-2"):
                        ui.button("Save", on_click=save).classes("runbtn").props(
                            "no-caps"
                        )
                        ui.button("Close", on_click=dialog.close).props(
                            "flat no-caps"
                        )
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
                    "Override large-workbook refusal",
                    value=False,
                    on_change=on_allow_large,
                ).tooltip(
                    "Process workbooks above local safety limits; workload "
                    "coverage is degraded"
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
                excel_path = state.files.get("current_excel") or state.files.get(
                    "baseline_excel"
                )
                ppt_path = state.files.get("current_ppt")
                state.available_sheets = (
                    peek_sheet_names(excel_path) if excel_path else []
                )
                state.available_slides = (
                    peek_slide_titles(ppt_path) if ppt_path else []
                )
                state.selected_sheets &= set(state.available_sheets)
                state.selected_slides &= {
                    index for index, _ in state.available_slides
                }
                scope_hint.visible = not (
                    state.available_sheets or state.available_slides
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
                ui.row().classes("gap-4 flex-wrap p-2"),
            ):
                for role in ROLES:
                    ui.input(
                        ROLE_LABELS[role],
                        password=True,
                        password_toggle_button=True,
                        on_change=lambda e, role=role: state.passwords.__setitem__(
                            role, e.value or ""
                        ),
                    ).classes("w-72").props("outlined dense")

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
                        rerun_of=state.rerun_of,
                        rerun_required=state.rerun_required,
                    )
                    supplied = [
                        ROLE_SHORT[role] for role in ROLES if role in state.files
                    ]
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
                        details.append("large-workbook refusal overridden")
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

                def start_run() -> None:
                    blockers = _run_blockers(
                        state.mode,
                        state.files,
                        rerun_of=state.rerun_of,
                        rerun_required=state.rerun_required,
                    )
                    if blockers:
                        ui.notify("; ".join(blockers), type="warning")
                        return
                    files = _files_for_mode(state.mode, state.files)
                    try:
                        profile = load_profile_by_name(profiles_dir, state.profile_name)
                    except Exception as exc:
                        ui.notify(f"Profile failed to load: {exc}", type="negative")
                        return
                    request = RunRequest(
                        request_id=new_request_id(),
                        work_dir=str(work_dir),
                        mode=state.mode.value,
                        profile_name=profile.name,
                        profile=profile.model_dump(mode="json"),
                        files={role: str(path) for role, path in files.items()},
                        display_files={role: path.name for role, path in files.items()},
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
            on_shutdown=request_shutdown,
        ):
            ui.label("Run history").classes("pagetitle")
            history = RunHistory(work_dir / "history.sqlite3")
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
                path = (
                    record.report_paths.get(str(e.args["kind"]))
                    if record is not None
                    else None
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
            if record.rerun_of is not None:
                heading += f" — re-QC of #{record.rerun_of}"
                try:
                    previous = history.get_run(record.rerun_of)
                    delta = compare_findings(previous.findings, record.findings)
                except KeyError:
                    delta = None
            report_paths = {kind: Path(path) for kind, path in record.report_paths.items()}
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
            )


def run_app(
    work_dir: Path,
    *,
    port: int = 8080,
    host: str = "127.0.0.1",
    network_mode: NetworkMode = NetworkMode.LOCAL,
    expires_at: dt.datetime | None = None,
    desktop_focus: bool = False,
) -> None:
    create_pages(
        work_dir,
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
                    save_server_config(work_dir, local_config())
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
        show=True,
        storage_secret=_storage_secret(work_dir),
    )
