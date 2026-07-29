"""NiceGUI local web application for the QC tool.

The UI stays thin: uploads, profile selection, passwords, one button.
All comparison logic lives in `qc_tool.engine`; `perform_run` wraps a run
with report writing and history recording and is used by both the UI and
the tests. Everything runs on localhost; sources are read-only. The
visual language lives in `qc_tool.ui.theme`.
"""

import asyncio
import datetime as dt
import logging
import re
import secrets
import shutil
import tempfile
import threading
from dataclasses import dataclass, field
from pathlib import Path

import yaml
from nicegui import app, events, run, ui

from qc_tool.config.profile import (
    CrosscheckMapping,
    DeliverableProfile,
    default_profile,
    load_profile,
    save_profile,
)
from qc_tool.coverage import QCRunMode
from qc_tool.crosscheck.trace import MappingSuggestion, SuggestedSource
from qc_tool.engine import FindingsDelta, QCRunResult, compare_findings, run_qc
from qc_tool.findings import Severity
from qc_tool.history.store import RunHistory, sha256_file
from qc_tool.io.decrypt import InvalidPasswordError, PasswordRequiredError
from qc_tool.progress import (
    CancellationToken,
    ProgressCallback,
    ProgressEvent,
    RunCancelled,
    RunPhase,
    check_cancelled,
    report_progress,
)
from qc_tool.report.excel_report import write_excel_report
from qc_tool.report.html_report import write_html_report
from qc_tool.security import private_directory, private_file, secure_managed_tree
from qc_tool.server_config import (
    NetworkMode,
    lan_config_matches,
    local_config,
    save_server_config,
)
from qc_tool.ui.guide import render_guide
from qc_tool.ui.theme import FINDINGS_BODY_SLOT, page_frame, section

logger = logging.getLogger(__name__)

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
    RunPhase.ANALYZING_POWERPOINT: "Analyzing PowerPoint",
    RunPhase.CROSSCHECKING: "Cross-checking package",
    RunPhase.WRITING_REPORTS: "Writing reports",
    RunPhase.RECORDING_HISTORY: "Recording history",
    RunPhase.COMPLETE: "Complete",
}


def _download_handler(path: str | Path):
    """Zero-arg click handler (NiceGUI passes event args to 1-arg lambdas)."""

    def handler() -> None:
        ui.download(str(path))

    return handler


@dataclass(slots=True)
class RunArtifacts:
    run_id: int
    result: QCRunResult
    report_paths: dict[str, Path]
    rerun_of: int | None = None
    delta: FindingsDelta | None = None


@dataclass(slots=True)
class SessionState:
    files: dict[str, Path] = field(default_factory=dict)
    passwords: dict[str, str] = field(default_factory=dict)  # role -> password
    profile_name: str = "default"
    mode: QCRunMode = QCRunMode.CYCLE_COMPARISON
    allow_large_workbooks: bool = False
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


def perform_run(
    work_dir: Path,
    files: dict[str, Path],
    passwords: dict[str, str],
    profile: DeliverableProfile,
    *,
    mode: QCRunMode = QCRunMode.CYCLE_COMPARISON,
    rerun_of: int | None = None,
    allow_large_workbooks: bool = False,
    cancellation_token: CancellationToken | None = None,
    on_progress: ProgressCallback | None = None,
) -> RunArtifacts:
    """Run QC, write both reports, and record the run in history."""
    password_by_file = {
        files[role].name: password
        for role, password in passwords.items()
        if role in files and password
    }
    result = run_qc(
        baseline_excel=files.get("baseline_excel"),
        current_excel=files.get("current_excel"),
        baseline_ppt=files.get("baseline_ppt"),
        current_ppt=files.get("current_ppt"),
        profile=profile,
        passwords=password_by_file,
        mode=mode,
        allow_large_workbooks=allow_large_workbooks,
        cancellation_token=cancellation_token,
        on_progress=on_progress,
    )
    check_cancelled(cancellation_token)
    runs_dir = private_directory(work_dir / "runs")
    stamp = dt.datetime.now(dt.UTC).strftime("%Y%m%dT%H%M%S.%fZ-")
    run_dir = Path(tempfile.mkdtemp(prefix=stamp, dir=runs_dir))
    private_directory(run_dir)
    report_paths = {
        "excel": run_dir / "qc_report.xlsx",
        "html": run_dir / "qc_report.html",
    }
    recorded = False
    try:
        report_progress(on_progress, RunPhase.WRITING_REPORTS, total=2)
        write_excel_report(result, report_paths["excel"])
        check_cancelled(cancellation_token)
        report_progress(
            on_progress,
            RunPhase.WRITING_REPORTS,
            processed=1,
            total=2,
        )
        write_html_report(result, report_paths["html"])
        check_cancelled(cancellation_token)
        report_progress(
            on_progress,
            RunPhase.WRITING_REPORTS,
            processed=2,
            total=2,
        )

        report_progress(on_progress, RunPhase.RECORDING_HISTORY, total=1)
        history = RunHistory(work_dir / "history.sqlite3")
        delta: FindingsDelta | None = None
        if rerun_of is not None:
            try:
                previous = history.get_run(rerun_of)
                delta = compare_findings(previous.findings, result.findings)
            except KeyError:
                logger.warning("re-QC referenced missing run %s", rerun_of)
                rerun_of = None
        file_hashes: dict[str, str] = {}
        for role, path in files.items():
            check_cancelled(cancellation_token)
            file_hashes[role] = sha256_file(path)
        check_cancelled(cancellation_token)
        run_id = history.record_run(
            result,
            file_hashes=file_hashes,
            report_paths={kind: str(path) for kind, path in report_paths.items()},
            file_paths={role: str(path) for role, path in files.items()},
            rerun_of=rerun_of,
        )
        recorded = True
        report_progress(
            on_progress,
            RunPhase.RECORDING_HISTORY,
            processed=1,
            total=1,
        )
        report_progress(on_progress, RunPhase.COMPLETE, processed=1, total=1)
    except BaseException:
        if not recorded:
            shutil.rmtree(run_dir, ignore_errors=True)
        raise
    return RunArtifacts(
        run_id=run_id,
        result=result,
        report_paths=report_paths,
        rerun_of=rerun_of,
        delta=delta,
    )


# --- pages -------------------------------------------------------------------


def _findings_rows(result: QCRunResult) -> list[dict[str, object]]:
    return [
        {
            "id": f.finding_id,
            "severity": (f.severity or Severity.WARNING).value,
            "class": f.finding_class.value,
            "where": f.sheet or f.slide or "",
            "location": f.location or f.baseline_location or "",
            "message": f.message,
            "baseline": f.baseline_value or "",
            "current": f.current_value or "",
            "element": f.element or "",
            "impacts": "; ".join(f.impacts),
            "artifact": f.artifact,
            "comment": f.analyst_comment,
            "overridden": f.severity_overridden,
            "root": f.root_cause_key,
            "waiver": (
                f"{f.waiver_reason} (expires {f.waiver_expires})"
                if f.waiver_reason
                else ""
            ),
            "bx": f.baseline_excerpt.model_dump() if f.baseline_excerpt else None,
            "cx": f.current_excerpt.model_dump() if f.current_excerpt else None,
        }
        for f in result.findings
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
) -> None:
    """Shared results renderer: stat strip, disclosures, exports, table."""
    all_rows = _findings_rows(result)
    findings_by_id = {f.finding_id: f for f in result.findings}
    section(heading)
    stats_box = ui.element("div").classes("statstrip")

    def render_stats() -> None:
        stats_box.clear()
        with stats_box:
            for severity, count in result.counts.items():
                with ui.column().classes(f"stat stat-{severity.value}"):
                    ui.label(str(count)).classes("n")
                    ui.label(severity.value).classes("l")
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

    section("Check coverage")
    ui.link("Coverage guide →", "/guide#coverage").classes("guide-jump")
    coverage_rows = [
        {
            "artifact": item.artifact,
            "check": item.label,
            "status": item.state.value,
            "findings": item.findings,
            "detail": item.detail,
        }
        for item in result.coverage
    ]
    ui.table(
        columns=[
            {"name": "artifact", "label": "Artifact", "field": "artifact"},
            {"name": "check", "label": "Check", "field": "check"},
            {"name": "status", "label": "Status", "field": "status"},
            {"name": "findings", "label": "Findings", "field": "findings"},
            {"name": "detail", "label": "Detail", "field": "detail"},
        ],
        rows=coverage_rows,
        row_key="check",
        pagination=20,
    ).classes("coverage-table").props("flat dense hide-bottom")

    if result.mapping_coverage is not None:
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

    with ui.row().classes("items-center gap-3 mt-3 w-full"):
        exports = (
            ("excel", "Export Excel (.xlsx)", write_excel_report),
            ("html", "Export HTML", write_html_report),
        )
        for kind, label, writer in exports:
            path = report_paths.get(kind)
            if path is None:
                continue

            def export(writer=writer, path=path) -> None:
                # Regenerate from the current (annotated) state before download.
                writer(result, path)
                ui.download(str(path))

            ui.button(label, on_click=export).classes("ghostbtn").props("no-caps flat")
        ui.label("exports are for sharing — findings are fully viewable below").classes(
            "text-xs text-gray-500"
        )
        severity_filter = (
            ui.select(
                [s.value for s in Severity],
                value=[s.value for s in Severity if s is not Severity.EXPECTED],
                multiple=True,
                label="Show severities",
            )
            .classes("w-72 ml-auto")
            .props("outlined dense use-chips")
        )

    visible = set(severity_filter.value or [])
    table = (
        ui.table(
            columns=_FINDINGS_COLUMNS,
            rows=[r for r in all_rows if r["severity"] in visible],
            row_key="id",
            pagination=25,
        )
        .classes("findings-table")
        .props("flat dense")
    )
    table.add_slot("body", FINDINGS_BODY_SLOT)

    def refilter() -> None:
        selected = set(severity_filter.value or [])
        table.rows = [r for r in all_rows if r["severity"] in selected]

    severity_filter.on_value_change(refilter)

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

    table.on("sev", on_severity)
    table.on("note", on_note)


def _render_results(container: ui.element, artifacts: RunArtifacts, work_dir: Path) -> None:
    container.clear()
    with container:
        _render_result_view(
            artifacts.result,
            artifacts.report_paths,
            f"4 · Results — run #{artifacts.run_id}",
            delta=artifacts.delta,
            rerun_of=artifacts.rerun_of,
            history=RunHistory(work_dir / "history.sqlite3"),
            run_id=artifacts.run_id,
            profiles_dir=work_dir / "profiles",
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
) -> None:
    secure_managed_tree(work_dir)
    uploads_dir = work_dir / "uploads"
    profiles_dir = work_dir / "profiles"
    private_directory(uploads_dir)
    private_directory(profiles_dir)

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
        ):
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
            upload_boxes: dict[str, ui.column] = {}

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
                for role, box in upload_boxes.items():
                    box.visible = (
                        state.mode is QCRunMode.CYCLE_COMPARISON
                        or role in {"current_excel", "current_ppt"}
                    )

            def on_mode_change(e: events.ValueChangeEventArguments) -> None:
                state.mode = QCRunMode(e.value)
                update_mode_surface()

            mode_select = ui.toggle(
                {mode.value: label for mode, label in MODE_LABELS.items()},
                value=state.mode.value,
                on_change=on_mode_change,
            ).classes("mode-select").props("no-caps spread")
            if rerun_record is not None:
                mode_select.disable()
            update_mode_surface()

            section("2 · Files")
            with ui.element("div").classes("upgrid"):
                for role in ROLES:

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
                        size_kb = max(1, e.file.size() // 1024)
                        file_states[role].text = f"{e.file.name} · {size_kb:,} KB"
                        file_states[role].classes(add="ok", remove="err")

                    with ui.column().classes("gap-0") as upload_box:
                        upload_boxes[role] = upload_box
                        ui.label(ROLE_LABELS[role]).classes("rolelabel")
                        ui.upload(
                            label=ROLE_ACCEPT[role],
                            auto_upload=True,
                            max_file_size=MAX_UPLOAD_BYTES,
                            on_upload=handle_upload,
                        ).props(f'accept="{ROLE_ACCEPT[role]}" flat')
                        file_states[role] = ui.label("no file selected").classes(
                            "filestate"
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

            section("3 · Configuration")
            ui.link("Profiles and controls guide →", "/guide#profiles").classes(
                "guide-jump"
            )
            with ui.row().classes("items-end gap-3 w-full"):
                profile_options = list_profiles(profiles_dir)
                initial_profile = "default"
                if rerun_record is not None and rerun_record.profile in profile_options:
                    initial_profile = rerun_record.profile
                    state.profile_name = initial_profile
                profile_select = (
                    ui.select(
                        profile_options,
                        value=initial_profile,
                        label="Deliverable profile",
                        on_change=lambda e: setattr(state, "profile_name", e.value),
                    )
                    .classes("w-64")
                    .props("outlined dense")
                )

                async def edit_profile() -> None:
                    name = state.profile_name
                    if name == "default":
                        ui.notify("Create a named profile first", type="warning")
                        return
                    path = _profile_path(profiles_dir, name)
                    content = path.read_text(encoding="utf-8") if path.exists() else ""
                    with ui.dialog() as dialog, ui.card().classes("w-[42rem]"):
                        ui.label(f"Profile: {name}").classes("font-bold")
                        editor = (
                            ui.textarea(value=content)
                            .classes("w-full")
                            .style("font-family: var(--font-mono); font-size: 0.8rem")
                            .props("rows=20 outlined")
                        )

                        def save() -> None:
                            try:
                                profile = DeliverableProfile.model_validate(
                                    yaml.safe_load(editor.value or "") or {}
                                )
                            except Exception as exc:  # surfaced to the analyst
                                ui.notify(f"Invalid profile: {exc}", type="negative")
                                return
                            save_profile(profile, path)
                            ui.notify("Profile saved")
                            dialog.close()

                        with ui.row():
                            ui.button("Save", on_click=save).classes("runbtn").props(
                                "no-caps"
                            )
                            ui.button("Cancel", on_click=dialog.close).props(
                                "flat no-caps"
                            )
                    dialog.open()

                ui.button("Edit profile", on_click=edit_profile).classes(
                    "ghostbtn"
                ).props("no-caps flat")

                new_profile_name = (
                    ui.input("New profile name").classes("w-48").props("outlined dense")
                )

                def create_profile() -> None:
                    name = (new_profile_name.value or "").strip()
                    if not name or name == "default":
                        ui.notify("Choose a non-empty, non-default name", type="warning")
                        return
                    try:
                        path = _profile_path(profiles_dir, name)
                    except ValueError as exc:
                        ui.notify(str(exc), type="warning")
                        return
                    save_profile(DeliverableProfile(name=name), path)
                    profile_select.options = list_profiles(profiles_dir)
                    profile_select.value = name
                    profile_select.update()
                    ui.notify(f"Profile {name!r} created")

                ui.button("Create", on_click=create_profile).classes("ghostbtn").props(
                    "no-caps flat"
                )

                spinner = ui.spinner(size="1.6rem").classes("ml-auto")
                spinner.visible = False

                ui.checkbox(
                    "Override large-workbook refusal",
                    value=False,
                    on_change=lambda e: setattr(
                        state,
                        "allow_large_workbooks",
                        bool(e.value),
                    ),
                ).tooltip(
                    "Process workbooks above local safety limits; workload coverage is degraded"
                )

                progress_label = ui.label("").classes("text-sm min-w-52")
                progress_label.visible = False
                progress_lock = threading.Lock()
                latest_progress: dict[str, ProgressEvent | None] = {"event": None}
                active_token: dict[str, CancellationToken | None] = {"token": None}

                def cancel_active_run() -> None:
                    token = active_token["token"]
                    if token is None:
                        return
                    token.cancel()
                    cancel_button.disable()
                    progress_label.set_text("Cancelling at the next safe boundary")

                cancel_button = ui.button(
                    "Cancel",
                    icon="stop_circle",
                    on_click=cancel_active_run,
                ).classes("ghostbtn").props("flat no-caps")
                cancel_button.visible = False

                def refresh_progress() -> None:
                    with progress_lock:
                        event = latest_progress["event"]
                    if event is None or active_token["token"] is None:
                        return
                    label = PHASE_LABELS[event.phase]
                    counts = (
                        f" ({event.processed}/{event.total})"
                        if event.total
                        else ""
                    )
                    detail = f" · {event.detail}" if event.detail else ""
                    progress_label.set_text(f"{label}{counts}{detail}")

                ui.timer(0.25, refresh_progress)

                async def start_run() -> None:
                    if state.rerun_of is not None:
                        missing = state.rerun_required - state.files.keys()
                        if missing:
                            names = ", ".join(
                                ROLE_LABELS[role] for role in sorted(missing)
                            )
                            ui.notify(
                                f"Re-QC of run #{state.rerun_of} is blocked — "
                                f"select these files again: {names}",
                                type="negative",
                            )
                            return
                    try:
                        files = _files_for_mode(state.mode, state.files)
                    except ValueError as exc:
                        ui.notify(str(exc), type="warning")
                        return
                    try:
                        profile = load_profile_by_name(profiles_dir, state.profile_name)
                    except Exception as exc:
                        ui.notify(f"Profile failed to load: {exc}", type="negative")
                        return
                    spinner.visible = True
                    run_button.disable()
                    cancellation_token = CancellationToken()
                    active_token["token"] = cancellation_token
                    with progress_lock:
                        latest_progress["event"] = None
                    progress_label.set_text("Preparing run")
                    progress_label.visible = True
                    cancel_button.enable()
                    cancel_button.visible = True

                    def receive_progress(event: ProgressEvent) -> None:
                        with progress_lock:
                            latest_progress["event"] = event

                    try:
                        artifacts = await run.io_bound(
                            perform_run,
                            work_dir,
                            files,
                            dict(state.passwords),
                            profile,
                            mode=state.mode,
                            rerun_of=state.rerun_of,
                            allow_large_workbooks=state.allow_large_workbooks,
                            cancellation_token=cancellation_token,
                            on_progress=receive_progress,
                        )
                    except RunCancelled:
                        ui.notify(
                            "Run cancelled; no successful run was recorded",
                            type="warning",
                        )
                        return
                    except PasswordRequiredError as exc:
                        ui.notify(f"{exc} — set it under Passwords", type="negative")
                        return
                    except InvalidPasswordError as exc:
                        ui.notify(str(exc), type="negative")
                        return
                    except Exception as exc:  # analyst-facing failure, never a crash
                        logger.exception("QC run failed")
                        ui.notify(f"QC run failed: {exc}", type="negative")
                        return
                    finally:
                        active_token["token"] = None
                        spinner.visible = False
                        run_button.enable()
                        progress_label.visible = False
                        cancel_button.visible = False
                    if artifacts is None:  # run.io_bound is typed Optional
                        ui.notify("QC run returned no result", type="negative")
                        return
                    ui.notify(f"Run #{artifacts.run_id} complete")
                    _render_results(results, artifacts, work_dir)

                run_button = ui.button("Run QC", on_click=start_run).classes(
                    "runbtn"
                ).props("no-caps")

            if rerun_banner_actions is not None:
                with rerun_banner_actions:
                    ui.button("Run QC now", on_click=start_run).classes("runbtn").props(
                        "no-caps dense"
                    )

            results = ui.column().classes("w-full")

    @ui.page("/guide")
    def guide_page() -> None:  # pyright: ignore[reportUnusedFunction]
        with page_frame(
            "guide",
            network_mode=network_mode.value,
            expires_at=expires_at.isoformat(timespec="seconds") if expires_at else None,
        ):
            render_guide()

    @ui.page("/history")
    def history_page() -> None:  # pyright: ignore[reportUnusedFunction]
        with page_frame(
            "history",
            network_mode=network_mode.value,
            expires_at=expires_at.isoformat(timespec="seconds") if expires_at else None,
        ):
            section("Run history")
            history = RunHistory(work_dir / "history.sqlite3")
            runs = history.list_runs()
            if not runs:
                ui.label("No runs recorded yet.").classes("lede")
                return
            for record in runs:
                with ui.card().classes("runcard"):
                    ui.label(
                        f"#{record.run_id} — "
                        f"{record.started_at.isoformat(timespec='seconds')}"
                        f" — {MODE_LABELS[record.mode]} — profile {record.profile!r}"
                    ).classes("runhead")
                    ui.label(
                        "  |  ".join(f"{r}: {n}" for r, n in record.files.items())
                    ).classes("runmeta")
                    with ui.element("div").classes("runcounts"):
                        for kind, count in record.counts.items():
                            with ui.row().classes("items-center gap-1 no-wrap"):
                                ui.element("span").classes(f"sevdot sev-{kind}")
                                ui.label(f"{kind} {count}")
                    with ui.row().classes("gap-2 mt-1"):
                        run_id = record.run_id
                        ui.button(
                            "View findings",
                            on_click=lambda run_id=run_id: ui.navigate.to(
                                f"/runs/{run_id}"
                            ),
                        ).classes("runbtn").props("no-caps dense")
                        ui.button(
                            "Re-QC",
                            on_click=lambda run_id=run_id: ui.navigate.to(
                                f"/?rerun={run_id}"
                            ),
                        ).classes("ghostbtn").props("no-caps flat dense")
                        for kind, path in record.report_paths.items():
                            if Path(path).exists():
                                ui.button(
                                    f"Export {kind}",
                                    on_click=_download_handler(path),
                                ).classes("ghostbtn").props("no-caps flat dense")

    @ui.page("/runs/{run_id}")
    def run_detail_page(run_id: int) -> None:  # pyright: ignore[reportUnusedFunction]
        with page_frame(
            "history",
            network_mode=network_mode.value,
            expires_at=expires_at.isoformat(timespec="seconds") if expires_at else None,
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
            result = QCRunResult(
                profile_name=record.profile,
                mode=record.mode,
                files=record.files,
                findings=record.findings,
                disclosures=record.disclosures,
                coverage=record.coverage,
                mapping_coverage=record.mapping_coverage,
                mapping_suggestions=record.mapping_suggestions,
                verified_crosschecks=record.verified_crosschecks,
            )
            heading = (
                f"Run #{record.run_id} — "
                f"{record.started_at.isoformat(timespec='seconds')} — "
                f"{MODE_LABELS[record.mode]} — profile {record.profile!r}"
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
            )
            with ui.row().classes("items-center gap-3 mt-2"):
                ui.label(
                    "  |  ".join(f"{r}: {n}" for r, n in record.files.items())
                ).classes("runmeta")
                ui.button(
                    "Re-QC this run",
                    on_click=lambda: ui.navigate.to(f"/?rerun={record.run_id}"),
                ).classes("ghostbtn").props("no-caps flat dense")


def run_app(
    work_dir: Path,
    *,
    port: int = 8080,
    host: str = "127.0.0.1",
    network_mode: NetworkMode = NetworkMode.LOCAL,
    expires_at: dt.datetime | None = None,
) -> None:
    create_pages(work_dir, network_mode=network_mode, expires_at=expires_at)
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
