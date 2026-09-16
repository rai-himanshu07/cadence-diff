"""Full-page, mode-aware configuration workspace (plan-20260913, Step 7).

The mandatory intermediate step between file intake (still handled by the
existing "/" page's uploads/mode/profile controls) and actual QC
submission. Auto-starts a bounded setup-analysis scan, lets the analyst
review structure and pick a profile only after the scan completes, and
performs the run/save actions through the same `RunRequest` shape the main
page's own retry paths use (`qc_tool.ui.app.build_run_request`).

Scope boundary (disclosed, matches the plan's own step ownership): this
module builds the SHELL -- auto-start/cancel/reconnect, post-analysis
profile selection with a field diff, a current/baseline preview toggle,
per-region mode/anchor/identity-column decisions, warning acknowledgement,
and the five final actions. It does not yet implement every region/column
semantic (blank-key threshold nuance, period bands, selector
prerequisites, PowerPoint/final-package slide review) -- those are Step
8/9's job, extending this same shell and its underlying
`qc_tool.ui.config_review` builder.
"""

from __future__ import annotations

import asyncio
import dataclasses
import json
import uuid
from pathlib import Path
from typing import Any

from nicegui import events, ui
from openpyxl.utils import get_column_letter

from qc_tool.config.profile import (
    DeliverableProfile,
    list_profiles,
    load_profile_by_name,
    new_profile,
    profile_path,
    profile_sha256,
    save_profile,
)
from qc_tool.config.resolved_input import ResolvedInputConfigurationV1
from qc_tool.coverage import QCRunMode
from qc_tool.history.config_session import ConfigSessionConflictError, ConfigSessionStore
from qc_tool.history.store import sha256_file
from qc_tool.package import MEMBER_ID_PATTERN, PackageManifest
from qc_tool.projection import project_cycle_volume
from qc_tool.run_service import REPORT_DEFER_FINDINGS
from qc_tool.runqueue import QueueBusyError, RunQueueManager, run_exclusive
from qc_tool.security import private_directory, private_file
from qc_tool.setup.coordinator import (
    SetupMemberInput,
    SetupStatus,
    get_setup_coordinator,
)
from qc_tool.setup.key_overlap_worker import KeyOverlapRequest, run_key_overlap_worker
from qc_tool.setup.models import MemberSetupProfile, SetupAnalysisResult
from qc_tool.setup.preview_store import SetupScanStore
from qc_tool.setup.preview_worker import (
    PreviewWindowRequest,
    run_preview_window_worker,
)
from qc_tool.ui.config_review import (
    ConfigWorkspaceState,
    DeckReview,
    MemberReview,
    SelectorDecision,
    add_selector,
    apply_anchor_click,
    apply_manual_range,
    apply_region_transform,
    build_input_contract,
    build_resolved_configuration,
    compute_key_overlap_warnings,
    compute_required_slide_warnings,
    compute_sheet_pairing_warnings,
    compute_slide_pairing_warnings,
    compute_warnings,
    compute_workspace_readiness,
    confirm_all_regions,
    deck_review_from_titles,
    diff_profile_against_scan,
    effective_region_data_range,
    effective_sheet_pairing,
    effective_slide_pairing,
    format_a1_range,
    is_clean_profile_diff,
    member_review_from_scan,
    parse_a1_cell,
    parse_a1_range,
    region_key_overlap_query_bounds,
    remove_selector,
    resolve_slide_anchors,
    set_column_baseline_letter,
    set_expected_refresh_columns,
    set_identity_columns,
    set_ignore_columns,
    set_ordinal_columns,
    set_region_data_start,
    set_region_footer_rows,
    set_selector_baseline_cell,
    set_sheet_rename,
    set_slide_included,
    set_slide_rename,
    sheet_pairing,
    update_region_decision,
)
from qc_tool.ui.credential_vault import CredentialVault
from qc_tool.ui.profile_editor import ProfileEditorController, open_profile_editor
from qc_tool.ui.theme import page_frame, section, status_chip

#: Excel role prefixes this workspace's setup scan understands.
_EXCEL_BASELINE_PREFIX = "baseline_excel"
_EXCEL_CURRENT_PREFIX = "current_excel"
#: PowerPoint role prefixes (Step 9's own slide-review scope).
_PPT_BASELINE_PREFIX = "baseline_ppt"
_PPT_CURRENT_PREFIX = "current_ppt"

#: One preview "window" fetched at a time -- comfortably inside
#: `preview_worker.MAX_PREVIEW_ROWS`/`MAX_PREVIEW_COLS`, small enough to stay
#: a lightweight DOM table.
_PREVIEW_WINDOW_ROWS = 20
_PREVIEW_WINDOW_COLS = 24


def _excel_members(files: dict[str, str]) -> dict[str, tuple[str | None, str | None]]:
    """``{member_id: (baseline_path_or_None, current_path_or_None)}`` for
    every Excel role present, parsing the same ``role`` / ``role:member``
    convention the rest of the app already uses.
    """
    members: dict[str, list[str | None]] = {}

    def _member(prefix: str, role: str) -> str:
        _, _, member_id = role.partition(":")
        return member_id or "primary"

    for role, path in files.items():
        if role == _EXCEL_BASELINE_PREFIX or role.startswith(_EXCEL_BASELINE_PREFIX + ":"):
            member_id = _member(_EXCEL_BASELINE_PREFIX, role)
            members.setdefault(member_id, [None, None])[0] = path
        elif role == _EXCEL_CURRENT_PREFIX or role.startswith(_EXCEL_CURRENT_PREFIX + ":"):
            member_id = _member(_EXCEL_CURRENT_PREFIX, role)
            members.setdefault(member_id, [None, None])[1] = path
    return {member_id: (pair[0], pair[1]) for member_id, pair in members.items()}


def _excel_role(prefix: str, member_id: str, roles: dict[str, str]) -> str:
    available = roles.keys()
    member_role = f"{prefix}:{member_id}"
    if member_role in available:
        return member_role
    return prefix


def _config_session_store(work_dir: Path) -> ConfigSessionStore:
    return ConfigSessionStore(work_dir / "history.sqlite3")


def build_session_choices(
    *,
    mode: QCRunMode,
    profile_name: str,
    files: dict[str, str],
    file_hashes: dict[str, str],
    output_mode: str,
    allow_large_workbooks: bool,
    allow_dependency_indexing: bool,
    acceptance_absolute: float,
    acceptance_percent: float,
    rerun_of: int | None,
    profile_snapshot: dict[str, object] | None = None,
    resolved_input_configuration: dict[str, object] | None = None,
) -> dict[str, object]:
    """The primitive-only payload persisted into `ConfigSessionStore` when
    the main intake page hands off to this workspace.
    """
    return {
        "mode": mode.value,
        "profile_name": profile_name,
        "files": dict(files),
        "file_hashes": dict(file_hashes),
        "output_mode": output_mode,
        "allow_large_workbooks": allow_large_workbooks,
        "allow_dependency_indexing": allow_dependency_indexing,
        "acceptance_absolute": acceptance_absolute,
        "acceptance_percent": acceptance_percent,
        "rerun_of": rerun_of,
        "profile_snapshot": profile_snapshot,
        "resolved_input_configuration": resolved_input_configuration,
    }


def _member_profile_from_payload(payload: dict[str, Any]) -> MemberSetupProfile:
    """Rehydrate one member's typed scan result from its own already
    self-produced, already-validated JSON round-trip (see
    ``qc_tool.setup.preview_worker``'s own "asdict() then JSON round-trip"
    contract) -- not external/untrusted input, so plain key access plus a
    permissive ``Any`` shape is appropriate here rather than a defensive
    isinstance wall.
    """
    from qc_tool.excel.complexity import WorkbookComplexity
    from qc_tool.excel.ranked_identity import RankedTableCandidate
    from qc_tool.excel.regions import TableRegion
    from qc_tool.setup.models import DetectedRegion, SheetSetupProfile, XlsbRiskProfile

    def _region(raw: dict[str, Any]) -> DetectedRegion:
        region_raw: dict[str, Any] = raw["region"]
        candidate_raw: dict[str, Any] | None = raw.get("ranked_candidate")
        candidate = None
        if candidate_raw is not None:
            candidate = RankedTableCandidate(
                columns=tuple(candidate_raw["columns"]),
                non_blank_coverage=candidate_raw["non_blank_coverage"],
                unique_ratio=candidate_raw["unique_ratio"],
                key_overlap=candidate_raw["key_overlap"],
                formula_ratio=candidate_raw["formula_ratio"],
                displaced_ratio=candidate_raw["displaced_ratio"],
                mismatch_reduction=candidate_raw["mismatch_reduction"],
                projected_positional_mismatches=candidate_raw[
                    "projected_positional_mismatches"
                ],
                ordinal_columns=tuple(candidate_raw.get("ordinal_columns", ())),
                header_row=candidate_raw.get("header_row"),
                manual_review=bool(candidate_raw.get("manual_review", False)),
            )
        return DetectedRegion(
            region=TableRegion(
                sheet=region_raw["sheet"],
                min_row=region_raw["min_row"],
                min_col=region_raw["min_col"],
                max_row=region_raw["max_row"],
                max_col=region_raw["max_col"],
                orientation=region_raw["orientation"],
                header_row=region_raw.get("header_row"),
                key_col=region_raw.get("key_col"),
                period_axis=region_raw["period_axis"],
            ),
            ranked_candidate=candidate,
        )

    def _sheet(raw: dict[str, Any]) -> SheetSetupProfile:
        return SheetSetupProfile(
            sheet_name=raw["sheet_name"],
            hidden=bool(raw.get("hidden", False)),
            very_hidden=bool(raw.get("very_hidden", False)),
            regions=tuple(_region(r) for r in raw.get("regions", ())),
            failure_detail=raw.get("failure_detail", ""),
        )

    def _risk(raw: dict[str, Any] | None):
        if raw is None:
            return None
        return XlsbRiskProfile(
            risky_features=tuple(raw.get("risky_features", ())),
            passive_features=tuple(raw.get("passive_features", ())),
            blocking_features=tuple(raw.get("blocking_features", ())),
            unknown_external_features=tuple(raw.get("unknown_external_features", ())),
            safe_for_external_engine=bool(raw.get("safe_for_external_engine", False)),
        )

    def _complexity(raw: dict[str, Any] | None):
        if raw is None:
            return None
        return WorkbookComplexity(
            formula_count=raw.get("formula_count", 0),
            lexical_formula_count=raw.get("lexical_formula_count", 0),
            reference_operands=raw.get("reference_operands", 0),
            resolved_range_cells=raw.get("resolved_range_cells", 0),
            projected_concrete_edges=raw.get("projected_concrete_edges", 0),
            interaction_rule_count=raw.get("interaction_rule_count", 0),
            warning_reasons=tuple(raw.get("warning_reasons", ())),
            override_used=bool(raw.get("override_used", False)),
            sampled_formulas=raw.get("sampled_formulas", 0),
            extras=dict(raw.get("extras", {})),
        )

    return MemberSetupProfile(
        member_id=payload["member_id"],
        baseline_hash=payload["baseline_hash"],
        current_hash=payload["current_hash"],
        baseline_sheets=tuple(_sheet(s) for s in payload.get("baseline_sheets", ())),
        current_sheets=tuple(_sheet(s) for s in payload.get("current_sheets", ())),
        baseline_complexity=_complexity(payload.get("baseline_complexity")),
        current_complexity=_complexity(payload.get("current_complexity")),
        baseline_xlsb_risk=_risk(payload.get("baseline_xlsb_risk")),
        current_xlsb_risk=_risk(payload.get("current_xlsb_risk")),
        failure_detail=payload.get("failure_detail", ""),
    )


_MODE_LABELS = {
    QCRunMode.CYCLE_COMPARISON: "Cycle comparison",
    QCRunMode.CURRENT_FILE_PREFLIGHT: "Current-file preflight",
    QCRunMode.FINAL_PACKAGE: "Final-package QC",
}
_REGION_MODE_OPTIONS = {
    "automatic": "Automatic",
    "keyed": "Match rows by key",
    "positional": "Compare by position",
    "excluded": "Exclude from this run",
}


def configuration_export_bytes(
    profile: DeliverableProfile,
    resolved: ResolvedInputConfigurationV1,
) -> bytes:
    """Build the private configuration export entirely in memory."""
    bundle = {
        "profile": profile.model_dump(mode="json", by_alias=True),
        "resolved_input_configuration": resolved.model_dump(mode="json"),
        "resolved_input_digest": resolved.canonical_sha256(),
    }
    return json.dumps(bundle, indent=2).encode("utf-8")


def _safe_upload_name(raw_name: str) -> str:
    name = Path(raw_name.replace("\\", "/")).name.strip()
    if not name or name in {".", ".."}:
        raise ValueError("uploaded file has no usable filename")
    return name


def _workspace_role_blockers(
    mode: QCRunMode, files: dict[str, Path]
) -> tuple[str, ...]:
    baseline_members = {
        role.partition(":")[2] or "primary"
        for role in files
        if role == _EXCEL_BASELINE_PREFIX
        or role.startswith(f"{_EXCEL_BASELINE_PREFIX}:")
    }
    current_members = {
        role.partition(":")[2] or "primary"
        for role in files
        if role == _EXCEL_CURRENT_PREFIX
        or role.startswith(f"{_EXCEL_CURRENT_PREFIX}:")
    }
    has_current_ppt = _PPT_CURRENT_PREFIX in files
    has_baseline_ppt = _PPT_BASELINE_PREFIX in files
    if mode is QCRunMode.CURRENT_FILE_PREFLIGHT:
        return () if current_members or has_current_ppt else (
            "Choose a current Excel workbook and/or PowerPoint deck",
        )
    if mode is QCRunMode.FINAL_PACKAGE:
        blockers = []
        if not current_members:
            blockers.append("Choose at least one current Excel workbook")
        if not has_current_ppt:
            blockers.append("Choose a current PowerPoint deck")
        return tuple(blockers)
    has_excel_pair = bool(baseline_members and baseline_members == current_members)
    has_ppt_pair = has_baseline_ppt and has_current_ppt
    if has_excel_pair or has_ppt_pair:
        return ()
    return (
        "Choose matching baseline/current Excel members and/or a PowerPoint pair",
    )


def _credentials_for_unchanged_sources(
    credentials: dict[str, str],
    previous_hashes: dict[str, str],
    current_hashes: dict[str, str],
) -> dict[str, str]:
    return {
        role: password
        for role, password in credentials.items()
        if password
        and previous_hashes.get(role) == current_hashes.get(role)
        and role in current_hashes
    }


def _coerce_str_dict(value: object) -> dict[str, str]:
    """Narrow a `ConfigSessionRecord.choices` value (stored as plain
    `dict[str, object]` since the session store round-trips arbitrary JSON)
    back to the flat string-keyed/string-valued shape `build_session_choices`
    always writes.
    """
    if not isinstance(value, dict):
        return {}
    return {str(key): str(item) for key, item in value.items()}


def _coerce_str_tuple(value: object) -> tuple[str, ...]:
    if not isinstance(value, (list, tuple)):
        return ()
    return tuple(str(item) for item in value)


def _coerce_letter_pairs(value: object) -> tuple[tuple[str, str], ...]:
    if not isinstance(value, (list, tuple)):
        return ()
    pairs: list[tuple[str, str]] = []
    for item in value:
        if isinstance(item, (list, tuple)) and len(item) == 2:
            pairs.append((str(item[0]), str(item[1])))
    return tuple(pairs)


def _coerce_float(value: object, default: float) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return default
    return float(value)


def _coerce_int(value: object, default: int) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        return default
    return value


def _coerce_optional_int(value: object) -> int | None:
    if isinstance(value, bool) or not isinstance(value, int):
        return None
    return value


def session_choice_overrides_from_resolved(
    resolved: ResolvedInputConfigurationV1,
) -> dict[str, object]:
    """Rebuild conservative workspace draft choices from one frozen run.

    Existing session choices take precedence because they retain richer fields
    such as exclusion expiry dates and selector display labels. This fallback
    prevents blocked-attempt recovery from silently dropping run-only choices
    when the original configuration session is unavailable.
    """
    region_decisions: dict[str, dict[str, object]] = {}
    selectors: dict[str, list[dict[str, object]]] = {}
    sheet_renames: dict[str, dict[str, str]] = {}
    for member in resolved.members:
        member_selectors: list[dict[str, object]] = []
        member_renames: dict[str, str] = {}
        for sheet in member.sheets:
            current_sheet = sheet.current_sheet_name
            baseline_sheet = sheet.baseline_sheet_name
            if current_sheet and baseline_sheet and current_sheet != baseline_sheet:
                member_renames[current_sheet] = baseline_sheet
            for region in sheet.regions:
                current_range = region.current_outer_range or region.current_data_range
                if current_range is None:
                    continue
                identity_columns = tuple(
                    column.current_letter
                    for column in region.columns
                    if column.alignment_role == "identity" and column.current_letter
                )
                ordinal_columns = tuple(
                    column.current_letter
                    for column in region.columns
                    if column.alignment_role == "ordinal" and column.current_letter
                )
                ignore_columns = tuple(
                    column.current_letter
                    for column in region.columns
                    if column.comparison_policy == "ignore" and column.current_letter
                )
                expected_refresh_columns = tuple(
                    column.current_letter
                    for column in region.columns
                    if column.comparison_policy == "expected_refresh"
                    and column.current_letter
                )
                baseline_letters = tuple(
                    (column.current_letter, column.baseline_letter)
                    for column in region.columns
                    if column.current_letter
                    and column.baseline_letter
                    and column.current_letter != column.baseline_letter
                )
                region_decisions[region.region_id] = {
                    "region_id": region.region_id,
                    "mode": region.mode,
                    "header_intent": region.header_intent,
                    "first_data_row": region.current_first_data_row,
                    "preamble_rows": region.current_preamble_rows,
                    "footer_rows": region.current_footer_rows,
                    "baseline_range": region.baseline_outer_range or "",
                    "identity_columns": list(identity_columns),
                    "ordinal_columns": list(ordinal_columns),
                    "ignore_columns": list(ignore_columns),
                    "expected_refresh_columns": list(expected_refresh_columns),
                    "column_baseline_letters": [list(pair) for pair in baseline_letters],
                    "trim_identity_whitespace": any(
                        column.trim_outer_whitespace for column in region.columns
                    ),
                    "blank_key_policy": region.blank_key_policy,
                    "duplicate_key_policy": region.duplicate_key_policy,
                    "exclusion_reason": region.degraded_reason,
                    "exclusion_expires_on": "",
                    "ignore_columns_reason": "",
                    "ignore_columns_expires_on": "",
                    "confirmed": region.coverage
                    in {
                        "confirmed",
                        "positional",
                        "excluded",
                        "degraded_acknowledged",
                    },
                }
            for selector in sheet.selectors:
                if current_sheet is None or selector.current_cell is None:
                    continue
                member_selectors.append(
                    {
                        "sheet_name": current_sheet,
                        "selector_id": selector.selector_id,
                        "label": selector.selector_id,
                        "cell": selector.current_cell,
                        "baseline_cell": selector.baseline_cell or "",
                    }
                )
        if member_selectors:
            selectors[member.member_id] = member_selectors
        if member_renames:
            sheet_renames[member.member_id] = member_renames
    return {
        "region_decisions": region_decisions,
        "selectors": selectors,
        "sheet_renames": sheet_renames,
        "warnings_acknowledged": list(resolved.warnings_acknowledged),
    }


def _profile_hash_or_none(profiles_dir: Path, name: str) -> str | None:
    """The saved profile file's content hash, or `None` when no file exists
    under that name yet (Step 11's "optimistic conflict protection"
    criterion: this is the value captured at the moment a profile name was
    selected in this session, compared against the file's hash again right
    before a save actually overwrites it).
    """
    path = profile_path(profiles_dir, name)
    if not path.exists():
        return None
    return sha256_file(path)


def _role_display_name(path_str: str) -> str:
    try:
        return Path(path_str).name
    except (TypeError, ValueError):
        return path_str


def _role_label(role: str) -> str:
    prefix, _, member_id = role.partition(":")
    labels = {
        _EXCEL_BASELINE_PREFIX: "Baseline workbook",
        _EXCEL_CURRENT_PREFIX: "Current workbook",
        _PPT_BASELINE_PREFIX: "Baseline presentation",
        _PPT_CURRENT_PREFIX: "Current presentation",
    }
    label = labels.get(prefix, prefix.replace("_", " ").title())
    return f"{label} · {member_id}" if member_id else label


def _role_sort_key(role: str) -> tuple[int, str, int]:
    prefix, _, member_id = role.partition(":")
    artifact_order = 0 if "excel" in prefix else 1
    side_order = 0 if prefix.startswith("baseline") else 1
    return artifact_order, member_id or "primary", side_order


def _serialize_region(region) -> dict[str, object]:
    return {
        "region_id": region.region_id,
        "mode": region.mode,
        "header_intent": region.header_intent,
        "first_data_row": region.first_data_row,
        "preamble_rows": region.preamble_rows,
        "footer_rows": region.footer_rows,
        "baseline_range": region.baseline_range,
        "identity_columns": list(region.identity_columns),
        "ordinal_columns": list(region.ordinal_columns),
        "ignore_columns": list(region.ignore_columns),
        "expected_refresh_columns": list(region.expected_refresh_columns),
        "column_baseline_letters": [list(pair) for pair in region.column_baseline_letters],
        "trim_identity_whitespace": region.trim_identity_whitespace,
        "blank_key_policy": region.blank_key_policy,
        "duplicate_key_policy": region.duplicate_key_policy,
        "exclusion_reason": region.exclusion_reason,
        "exclusion_expires_on": region.exclusion_expires_on,
        "ignore_columns_reason": region.ignore_columns_reason,
        "ignore_columns_expires_on": region.ignore_columns_expires_on,
        "confirmed": region.confirmed,
    }


def _apply_saved_region_choices(
    member: MemberReview, saved: dict[str, dict[str, object]]
) -> MemberReview:
    """Overlay previously saved decisions (draft recovery) onto a freshly
    scanned member review, matched by ``region_id``. A region the fresh
    scan no longer has is silently dropped -- the scan is authoritative for
    structure, the session store only for choices.
    """
    from dataclasses import replace as _replace

    new_sheets = []
    for sheet in member.current_sheets:
        new_regions = []
        for region in sheet.regions:
            saved_region = saved.get(region.region_id)
            if saved_region is None:
                new_regions.append(region)
                continue
            new_regions.append(
                _replace(
                    region,
                    mode=saved_region.get("mode", region.mode),
                    header_intent=saved_region.get("header_intent", region.header_intent),
                    first_data_row=(
                        _coerce_optional_int(saved_region["first_data_row"])
                        if "first_data_row" in saved_region
                        else region.first_data_row
                    ),
                    preamble_rows=_coerce_int(
                        saved_region.get("preamble_rows"), region.preamble_rows
                    ),
                    footer_rows=_coerce_int(saved_region.get("footer_rows"), region.footer_rows),
                    baseline_range=str(
                        saved_region.get("baseline_range", region.baseline_range) or ""
                    ),
                    identity_columns=_coerce_str_tuple(saved_region.get("identity_columns", ())),
                    ordinal_columns=_coerce_str_tuple(saved_region.get("ordinal_columns", ())),
                    ignore_columns=_coerce_str_tuple(saved_region.get("ignore_columns", ())),
                    expected_refresh_columns=_coerce_str_tuple(
                        saved_region.get("expected_refresh_columns", ())
                    ),
                    column_baseline_letters=_coerce_letter_pairs(
                        saved_region.get("column_baseline_letters", ())
                    ),
                    trim_identity_whitespace=bool(
                        saved_region.get(
                            "trim_identity_whitespace", region.trim_identity_whitespace
                        )
                    ),
                    blank_key_policy=saved_region.get("blank_key_policy", region.blank_key_policy),
                    duplicate_key_policy=saved_region.get(
                        "duplicate_key_policy", region.duplicate_key_policy
                    ),
                    exclusion_reason=saved_region.get("exclusion_reason", ""),
                    exclusion_expires_on=saved_region.get("exclusion_expires_on", ""),
                    ignore_columns_reason=saved_region.get("ignore_columns_reason", ""),
                    ignore_columns_expires_on=saved_region.get("ignore_columns_expires_on", ""),
                    confirmed=bool(saved_region.get("confirmed", False)),
                )
            )
        new_sheets.append(_replace(sheet, regions=tuple(new_regions)))
    return _replace(member, current_sheets=tuple(new_sheets))


def _apply_pending_row_matching(
    member: MemberReview, pending: list[dict[str, object]]
) -> MemberReview:
    """Overlay row-matching proposals from a blocked attempt for review."""
    from dataclasses import replace as _replace

    new_sheets = []
    for sheet in member.current_sheets:
        new_regions = []
        for region in sheet.regions:
            proposal = next(
                (
                    item
                    for item in pending
                    if item.get("member_id", "primary") == member.member_id
                    and item.get("sheet") == sheet.sheet_name
                    and (
                        item.get("current_range") == region.current_range
                        or item.get("anchor_cell") == region.anchor_cell
                    )
                ),
                None,
            )
            if proposal is None:
                new_regions.append(region)
                continue
            header_row = _coerce_optional_int(proposal.get("header_row"))
            new_regions.append(
                _replace(
                    region,
                    mode="keyed",
                    header_intent=(
                        "first_data_row"
                        if header_row is not None
                        else region.header_intent
                    ),
                    first_data_row=(
                        header_row + 1 if header_row is not None else region.first_data_row
                    ),
                    identity_columns=_coerce_str_tuple(
                        proposal.get("identity_columns", ())
                    ),
                    ordinal_columns=_coerce_str_tuple(
                        proposal.get("ordinal_columns", ())
                    ),
                    duplicate_key_policy=str(
                        proposal.get("duplicate_policy", "skip")
                    ),
                    ranked_candidate_pending=False,
                    confirmed=True,
                )
            )
        new_sheets.append(_replace(sheet, regions=tuple(new_regions)))
    return _replace(member, current_sheets=tuple(new_sheets))


def _serialize_selectors(member: MemberReview) -> list[dict[str, object]]:
    entries: list[dict[str, object]] = []
    for sheet in member.current_sheets:
        for selector in sheet.selectors:
            entries.append(
                {
                    "sheet_name": sheet.sheet_name,
                    "selector_id": selector.selector_id,
                    "label": selector.label,
                    "cell": selector.cell,
                    "baseline_cell": selector.baseline_cell,
                }
            )
    return entries


def _apply_saved_selectors(member: MemberReview, saved: list[dict[str, object]]) -> MemberReview:
    """Restore analyst-declared selector prerequisites (draft recovery),
    matched by sheet name -- the scan is authoritative for which sheets
    exist, the session store only for which selectors an analyst added.
    """
    from dataclasses import replace as _replace

    by_sheet: dict[str, list[SelectorDecision]] = {}
    for entry in saved:
        sheet_name = str(entry.get("sheet_name", ""))
        if not sheet_name:
            continue
        by_sheet.setdefault(sheet_name, []).append(
            SelectorDecision(
                selector_id=str(entry.get("selector_id", "")),
                label=str(entry.get("label", "")),
                cell=str(entry.get("cell", "")),
                baseline_cell=str(entry.get("baseline_cell", "")),
            )
        )
    if not by_sheet:
        return member
    new_sheets = []
    for sheet in member.current_sheets:
        restored = by_sheet.get(sheet.sheet_name)
        if not restored:
            new_sheets.append(sheet)
            continue
        new_sheets.append(_replace(sheet, selectors=tuple(restored)))
    return _replace(member, current_sheets=tuple(new_sheets))


def _render_preview_grid(
    outcome: Any,
    *,
    on_cell_click,
    effective_data_range: str | None = None,
) -> None:
    """Render a bounded, read-only preview window as a plain HTML table.
    Every cell is clickable (Step 7's "click controls" criterion) so an
    armed anchor-pick action can resolve a cell without drag selection.
    """
    with ui.element("div").classes("previewscroll"), ui.element("table").classes(
        "previewgrid"
    ).mark("active-preview-grid"):
        with ui.element("thead"), ui.element("tr"):
            ui.element("th")
            for col_offset in range(len(outcome.rows[0]) if outcome.rows else 0):
                col_number = outcome.resolved_min_col + col_offset
                with ui.element("th"):
                    ui.label(get_column_letter(col_number))
        data_bounds = (
            parse_a1_range(effective_data_range)
            if effective_data_range is not None
            else None
        )
        with ui.element("tbody"):
            for row_offset, row_values in enumerate(outcome.rows):
                row_number = outcome.resolved_min_row + row_offset
                formula_row = (
                    outcome.formula_cells[row_offset]
                    if row_offset < len(outcome.formula_cells)
                    else []
                )
                preview_row = ui.element("tr")
                if data_bounds is not None:
                    data_min_row, _min_col, data_max_row, _max_col = data_bounds
                    preview_row.classes(
                        "previewrow-data"
                        if data_min_row <= row_number <= data_max_row
                        else "previewrow-outside"
                    )
                with preview_row:
                    with ui.element("th"):
                        ui.label(str(row_number))
                    for col_offset, text in enumerate(row_values):
                        col_number = outcome.resolved_min_col + col_offset
                        is_formula = (
                            bool(formula_row[col_offset])
                            if col_offset < len(formula_row)
                            else False
                        )
                        cell = ui.element("td").classes("previewcell")
                        if is_formula:
                            cell.classes("previewcell-formula")
                        with cell:
                            formula_text = outcome.formula_text.get(
                                f"{row_number},{col_number}"
                            )
                            ui.label(formula_text or text or "\u00a0")
                        cell.on(
                            "click",
                            lambda _event, r=row_number, c=col_number: on_cell_click(r, c),
                        )


def render_config_workspace(
    work_dir: Path,
    profiles_dir: Path,
    session_key: str,
    *,
    queue_manager: RunQueueManager,
    credential_vault: CredentialVault | None = None,
    network_mode: str = "local",
    on_settings=None,
    on_shutdown=None,
) -> None:
    """Render the full-page mode-aware configuration workspace for one
    already-created config session. Registered by `qc_tool.ui.app` at
    ``/configure``.
    """
    session_store = _config_session_store(work_dir)
    record = session_store.get(session_key)
    if record is None or "files" not in record.choices:
        with page_frame(
            "run", network_mode=network_mode, on_settings=on_settings, on_shutdown=on_shutdown
        ):
            ui.label("This configuration session was not found").classes("runhead")
            ui.label(
                "It may have expired, or the link is stale. Return to Compare and "
                "select your files again."
            ).classes("note")
            ui.link("Back to Compare", "/").classes("guide-jump")
        return

    choices = record.choices
    session_revision: dict[str, int] = {"value": record.revision}
    raw_profile_snapshot = choices.get("profile_snapshot")

    def _load_workspace_profile(name: str) -> DeliverableProfile:
        if isinstance(raw_profile_snapshot, dict):
            frozen = DeliverableProfile.model_validate(raw_profile_snapshot)
            if frozen.name == name:
                return frozen
        return load_profile_by_name(profiles_dir, name)

    raw_files: dict[str, str] = _coerce_str_dict(choices.get("files", {}))
    file_hashes: dict[str, str] = _coerce_str_dict(choices.get("file_hashes", {}))
    credentials = (
        credential_vault.claim_snapshot(
            record.session_id, record.input_generation, file_hashes
        )
        if credential_vault is not None
        else {}
    )
    try:
        mode = QCRunMode(str(choices.get("mode", "cycle_comparison")))
    except ValueError:
        mode = QCRunMode.CYCLE_COMPARISON

    invalid_roles: list[str] = []
    files: dict[str, Path] = {}
    for role, path_str in raw_files.items():
        path = Path(path_str)
        expected_hash = file_hashes.get(role)
        if not path.exists() or (expected_hash and sha256_file(path) != expected_hash):
            invalid_roles.append(role)
            continue
        files[role] = path

    with page_frame(
        "run", network_mode=network_mode, on_settings=on_settings, on_shutdown=on_shutdown
    ):
        ui.button(
            "Back to files",
            icon="arrow_back",
            on_click=lambda: ui.navigate.to(
                f"/?config_session={record.session_id}"
            ),
        ).classes("ghostbtn").props("flat no-caps")
        ui.label(f"Configure & run — {_MODE_LABELS[mode]}").classes("pagetitle")
        ui.label(
            "Review the structure of your files, resolve how rows and columns "
            "match, and confirm before QC runs. This scan is bounded and never "
            "becomes evidence -- the actual run reloads your files fresh."
        ).classes("lede")

        def _persist_input_state(
            *,
            next_mode: QCRunMode = mode,
            next_files: dict[str, str] | None = None,
            next_hashes: dict[str, str] | None = None,
            choice_updates: dict[str, object] | None = None,
        ) -> None:
            updated_files = dict(raw_files if next_files is None else next_files)
            updated_hashes = dict(file_hashes if next_hashes is None else next_hashes)
            try:
                saved = session_store.save_choices(
                    record.session_id,
                    profile_name=str(choices.get("profile_name", "default")),
                    choices={
                        **choices,
                        "mode": next_mode.value,
                        "files": updated_files,
                        "file_hashes": updated_hashes,
                        **(choice_updates or {}),
                    },
                    expected_revision=session_revision["value"],
                )
            except ConfigSessionConflictError:
                ui.notify(
                    "This configuration changed in another tab. Reload before "
                    "changing files.",
                    type="warning",
                )
                return
            session_revision["value"] = saved.revision
            setup_coordinator = get_setup_coordinator(work_dir)
            if saved.input_generation != record.input_generation:
                setup_coordinator.discard(
                    record.session_id, record.input_generation
                )
            if credential_vault is not None:
                credential_vault.store_bundle(
                    saved.session_id,
                    input_generation=saved.input_generation,
                    source_hashes=updated_hashes,
                    credentials=_credentials_for_unchanged_sources(
                        credentials, file_hashes, updated_hashes
                    ),
                )
            credentials.clear()
            ui.navigate.to(f"/configure?session={record.session_id}")

        def _change_mode(event: events.ValueChangeEventArguments) -> None:
            try:
                next_mode = QCRunMode(str(event.value))
            except ValueError:
                return
            _persist_input_state(next_mode=next_mode)

        ui.toggle(
            {item.value: label for item, label in _MODE_LABELS.items()},
            value=mode.value,
            on_change=_change_mode,
        ).classes("mode-select").props("no-caps spread")

        if invalid_roles:
            ui.label(
                "Some selected files are missing or changed since intake: "
                + ", ".join(sorted(invalid_roles))
            ).classes("notecard")

        section("1 · Input roles")
        role_options = set(raw_files)
        pending_roles_raw = choices.get("pending_roles", [])
        if isinstance(pending_roles_raw, list):
            role_options.update(
                role for role in pending_roles_raw if isinstance(role, str)
            )
        if mode is QCRunMode.CYCLE_COMPARISON:
            role_options.update(
                {
                    _EXCEL_BASELINE_PREFIX,
                    _EXCEL_CURRENT_PREFIX,
                    _PPT_BASELINE_PREFIX,
                    _PPT_CURRENT_PREFIX,
                }
            )
        else:
            role_options.update({_EXCEL_CURRENT_PREFIX, _PPT_CURRENT_PREFIX})
        with ui.element("div").classes("config-files-grid w-full"):
            for role in sorted(role_options, key=_role_sort_key):
                with ui.element("div").classes("config-file-row w-full"):
                    if role in files:
                        status_chip(
                            "ok",
                            _role_label(role),
                            _role_display_name(raw_files[role]),
                        )
                    elif role in invalid_roles:
                        ui.label(f"{_role_label(role)} · missing or changed").classes(
                            "notecard"
                        )
                    else:
                        ui.label(f"{_role_label(role)} · not selected").classes(
                            "note"
                        )

                    async def _replace_file(
                        event: events.UploadEventArguments,
                        role: str = role,
                    ) -> None:
                        try:
                            filename = _safe_upload_name(event.file.name)
                        except ValueError as exc:
                            ui.notify(str(exc), type="negative")
                            return
                        suffix = Path(filename).suffix.casefold()
                        allowed = (
                            {".xlsx", ".xlsm", ".xlsb"}
                            if "_excel" in role
                            else {".pptx"}
                        )
                        if suffix not in allowed:
                            ui.notify(
                                "Choose a supported file for this role.",
                                type="warning",
                            )
                            return
                        upload_dir = private_directory(
                            work_dir
                            / "uploads"
                            / "configuration"
                            / record.session_id
                            / f"{record.input_generation + 1}-{uuid.uuid4().hex}"
                            / role.replace(":", "-")
                        )
                        temporary = upload_dir / f".{uuid.uuid4().hex}.upload"
                        target = upload_dir / filename
                        try:
                            await event.file.save(temporary)
                            private_file(temporary)
                            digest = await asyncio.to_thread(sha256_file, temporary)
                            temporary.replace(target)
                            private_file(target)
                        finally:
                            temporary.unlink(missing_ok=True)
                        updated_files = dict(raw_files)
                        updated_hashes = dict(file_hashes)
                        updated_files[role] = str(target)
                        updated_hashes[role] = digest
                        _persist_input_state(
                            next_files=updated_files,
                            next_hashes=updated_hashes,
                        )

                    ui.upload(
                        label="Replace" if role in raw_files else "Choose",
                        auto_upload=True,
                        max_file_size=512 * 1024 * 1024,
                        on_upload=_replace_file,
                    ).props(
                        'accept=".xlsx,.xlsm,.xlsb" flat'
                        if "_excel" in role
                        else 'accept=".pptx" flat'
                    ).classes("config-role-upload")

                    def _clear_file(role: str = role) -> None:
                        updated_files = dict(raw_files)
                        updated_hashes = dict(file_hashes)
                        updated_files.pop(role, None)
                        updated_hashes.pop(role, None)
                        _persist_input_state(
                            next_files=updated_files,
                            next_hashes=updated_hashes,
                        )

                    ui.button(
                        icon="delete_outline",
                        on_click=_clear_file,
                    ).props("flat round dense").tooltip(f"Clear {_role_label(role)}")

            with ui.row().classes("items-end gap-2"):
                new_member_id = ui.input("New Excel member ID").props(
                    "outlined dense"
                ).classes("w-56")

                def _add_member_roles() -> None:
                    import re

                    member_id = str(new_member_id.value or "").strip()
                    if re.fullmatch(MEMBER_ID_PATTERN, member_id) is None:
                        ui.notify(
                            "Member ID must use letters, numbers, dots, dashes, "
                            "or underscores.",
                            type="warning",
                        )
                        return
                    prefixes = (
                        (_EXCEL_BASELINE_PREFIX, _EXCEL_CURRENT_PREFIX)
                        if mode is QCRunMode.CYCLE_COMPARISON
                        else (_EXCEL_CURRENT_PREFIX,)
                    )
                    pending = {
                        role
                        for role in pending_roles_raw
                        if isinstance(role, str)
                    } if isinstance(pending_roles_raw, list) else set()
                    pending.update(f"{prefix}:{member_id}" for prefix in prefixes)
                    _persist_input_state(
                        choice_updates={"pending_roles": sorted(pending)}
                    )

                ui.button(
                    "Add workbook member",
                    icon="add",
                    on_click=_add_member_roles,
                ).classes("ghostbtn").props("flat no-caps")

        members = _excel_members({role: str(path) for role, path in files.items()})
        # Preflight/final-package have no baseline side -- self-scan the
        # current file so structural facts (sheets/regions/visibility) are
        # still reviewable without inventing a fake baseline concept.
        scan_members: dict[str, tuple[str, str]] = {}
        for member_id, (baseline_path, current_path) in members.items():
            if current_path is None:
                continue
            if mode is QCRunMode.CYCLE_COMPARISON and baseline_path is None:
                continue
            scan_members[member_id] = (baseline_path or current_path, current_path)

        coordinator = get_setup_coordinator(work_dir)
        client = ui.context.client
        subscriber_id = str(client.id)
        job = coordinator.attach(
            record.session_id, record.input_generation, subscriber_id
        )
        setup_member_inputs: dict[str, SetupMemberInput] = {}
        passwords_by_member: dict[str, dict[str, str]] = {}
        for member_id, (baseline_path, current_path) in scan_members.items():
            current_role = _excel_role(
                _EXCEL_CURRENT_PREFIX, member_id, file_hashes
            )
            baseline_role = _excel_role(
                _EXCEL_BASELINE_PREFIX, member_id, file_hashes
            )
            current_hash = file_hashes[current_role]
            baseline_hash = file_hashes.get(baseline_role, current_hash)
            setup_member_inputs[member_id] = SetupMemberInput(
                member_id=member_id,
                baseline_path=baseline_path,
                current_path=current_path,
                baseline_hash=baseline_hash,
                current_hash=current_hash,
            )
            member_passwords = {
                "baseline": credentials.get(baseline_role, ""),
                "current": credentials.get(current_role, ""),
            }
            if any(member_passwords.values()):
                passwords_by_member[member_id] = member_passwords

        def _rebuild_passwords_by_member() -> None:
            passwords_by_member.clear()
            for member_id in scan_members:
                current_role = _excel_role(
                    _EXCEL_CURRENT_PREFIX, member_id, file_hashes
                )
                baseline_role = _excel_role(
                    _EXCEL_BASELINE_PREFIX, member_id, file_hashes
                )
                member_passwords = {
                    "baseline": credentials.get(baseline_role, ""),
                    "current": credentials.get(current_role, ""),
                }
                if any(member_passwords.values()):
                    passwords_by_member[member_id] = member_passwords

        def _missing_password_roles() -> tuple[str, ...]:
            roles: list[str] = []
            snapshot = job.snapshot()
            requirements = (
                snapshot.credential_roles_required
                or snapshot.missing_credential_roles
            )
            if not requirements and snapshot.status is not SetupStatus.AWAITING_CREDENTIALS:
                return ()
            for member_id, sides in requirements.items():
                current_role = _excel_role(
                    _EXCEL_CURRENT_PREFIX, member_id, file_hashes
                )
                baseline_role = _excel_role(
                    _EXCEL_BASELINE_PREFIX, member_id, file_hashes
                )
                for side in sides:
                    role = baseline_role if side == "baseline" else current_role
                    if role not in files and current_role in files:
                        role = current_role
                    if role in files and role not in roles:
                        roles.append(role)
            if not roles:
                roles.extend(
                    sorted(role for role in files if "_excel" in role)
                )
            return tuple(roles)

        section("2 · Setup analysis")
        scan_status_box = ui.column().classes("gap-2 w-full")
        profile_box = ui.column().classes("gap-2 w-full")
        regions_box = ui.column().classes("gap-2 w-full")
        preview_box = ui.column().classes("gap-2 w-full")
        deck_box = ui.column().classes("gap-2 w-full")
        section("3 · Review & run")
        warnings_box = ui.column().classes("gap-2 w-full")
        advanced_box = ui.column().classes("gap-2 w-full")
        actions_box = ui.row().classes(
            "config-workspace-actions items-center gap-2 flex-wrap"
        )

        workspace_state: dict[str, ConfigWorkspaceState] = {
            "value": ConfigWorkspaceState(
                mode=mode,
                profile_name=str(choices.get("profile_name", "default")),
                allow_large_workbooks=bool(choices.get("allow_large_workbooks", False)),
                allow_dependency_indexing=bool(choices.get("allow_dependency_indexing", False)),
                profile_opened_hash=_profile_hash_or_none(
                    profiles_dir, str(choices.get("profile_name", "default"))
                ),
            )
        }
        selected_profile: dict[str, DeliverableProfile | None] = {"value": None}
        #: One in-memory preview panel: member/sheet/window position plus the
        #: last fetched outcome. Never persisted -- a refresh regenerates it
        #: on demand (matches the plan's own "preview content regenerates,
        #: choices restore" boundary).
        preview_state: dict[str, Any] = {
            "member_id": None,
            "sheet_name": None,
            "region_id": None,
            "min_row": 1,
            "min_col": 1,
            "reveal_formulas": False,
            "loading": False,
            "error": "",
            "outcome": None,
            "auto_request_key": None,
            "request_token": 0,
            "editor_tab": "rows",
            #: (member_id, sheet_name, region_id) currently armed to receive
            #: the next preview-grid cell click as its new anchor, or None.
            "anchor_target": None,
        }
        #: Last-computed confirmed-identity key-overlap ratio per
        #: region_id (Step 12 Fix 5) -- an ephemeral, on-demand query
        #: result, never a persisted draft choice; a region absent here was
        #: never queried (mirrors ``preview_state``'s own "regenerates on
        #: demand" convention).
        key_overlap_state: dict[str, dict[str, float]] = {"ratios": {}}
        #: Strong references to in-flight fire-and-forget background tasks
        #: (Step 12 Fix 5's key-overlap recompute) so the event loop cannot
        #: garbage-collect one mid-flight; each removes itself on completion.
        background_tasks: set[asyncio.Task[None]] = set()
        preview_task: dict[str, asyncio.Task[None] | None] = {"value": None}

        def _schedule_preview_load() -> None:
            preview_state["request_token"] += 1
            active_task = preview_task["value"]
            if active_task is not None and not active_task.done():
                return
            task = asyncio.create_task(_drain_preview_requests())
            preview_task["value"] = task
            background_tasks.add(task)

            def _preview_done(done: asyncio.Task[None]) -> None:
                background_tasks.discard(done)
                if preview_task["value"] is done:
                    preview_task["value"] = None

            task.add_done_callback(_preview_done)
        #: The analyst-typed target profile name for "Save profile"/"Save
        #: profile and run" (Step 12 fix: these two actions previously had
        #: NO way to reach a real, non-empty target name whenever the
        #: currently-selected profile was the immutable "default" -- every
        #: click silently failed with "Choose a profile name before saving".
        #: Persists across `refresh()` re-renders (a plain re-seeded
        #: `ui.input(value=...)` would otherwise discard an in-progress
        #: keystroke on any unrelated state change).
        save_as_state: dict[str, str] = {"name": ""}
        profile_editor_open: dict[str, bool] = {"value": False}
        hydrated_profile_shape: dict[str, tuple[tuple[str, tuple[str, ...]], ...]] = {
            "value": ()
        }
        credential_prompt_prepared: set[tuple[str, ...]] = set()

        def _job_result() -> SetupAnalysisResult | None:
            payloads = job.result_payloads_snapshot()
            if not payloads:
                return None
            return SetupAnalysisResult(
                members={
                    member_id: _member_profile_from_payload(payload)
                    for member_id, payload in payloads.items()
                }
            )

        def _setup_complete() -> bool:
            excel_ready = (
                not scan_members or job.snapshot().status is SetupStatus.COMPLETE
            )
            deck_ready = (
                _PPT_CURRENT_PREFIX not in files
                or workspace_state["value"].deck_review is not None
            )
            return excel_ready and deck_ready

        def _persist_choices() -> None:
            choices.pop("pending_row_matching", None)
            state = workspace_state["value"]
            region_choices: dict[str, dict[str, object]] = {}
            selector_choices: dict[str, list[dict[str, object]]] = {}
            rename_choices: dict[str, dict[str, str]] = {}
            for member in state.member_reviews:
                for sheet in member.current_sheets:
                    for region in sheet.regions:
                        region_choices[region.region_id] = _serialize_region(region)
                member_selectors = _serialize_selectors(member)
                if member_selectors:
                    selector_choices[member.member_id] = member_selectors
                if member.sheet_renames:
                    rename_choices[member.member_id] = dict(member.sheet_renames)
            deck = state.deck_review
            excluded_slides = (
                [s.slide_index for s in deck.current_slides if not s.included]
                if deck is not None
                else []
            )
            slide_renames = dict(deck.slide_renames) if deck is not None else {}
            try:
                saved = session_store.save_choices(
                    record.session_id,
                    profile_name=state.profile_name,
                    choices={
                        **choices,
                        "region_decisions": region_choices,
                        "selectors": selector_choices,
                        "sheet_renames": rename_choices,
                        "excluded_slides": excluded_slides,
                        "slide_renames": slide_renames,
                        "allow_large_workbooks": state.allow_large_workbooks,
                        "allow_dependency_indexing": state.allow_dependency_indexing,
                    },
                    expected_revision=session_revision["value"],
                )
            except ConfigSessionConflictError:
                ui.notify(
                    "This configuration changed in another tab. Reload or take "
                    "over before saving more changes.",
                    type="warning",
                )
                return
            session_revision["value"] = saved.revision

        def _seed_member_reviews() -> None:
            result = _job_result()
            if result is None:
                return
            saved_regions_raw = choices.get("region_decisions", {})
            saved_regions: dict[str, dict[str, object]] = (
                dict(saved_regions_raw) if isinstance(saved_regions_raw, dict) else {}
            )
            saved_selectors_raw = choices.get("selectors", {})
            saved_selectors: dict[str, object] = (
                dict(saved_selectors_raw) if isinstance(saved_selectors_raw, dict) else {}
            )
            saved_renames_raw = choices.get("sheet_renames", {})
            saved_renames: dict[str, object] = (
                dict(saved_renames_raw) if isinstance(saved_renames_raw, dict) else {}
            )
            pending_row_matching_raw = choices.get("pending_row_matching", [])
            pending_row_matching = (
                [
                    dict(item)
                    for item in pending_row_matching_raw
                    if isinstance(item, dict)
                ]
                if isinstance(pending_row_matching_raw, list)
                else []
            )
            reviews = []
            for member_id, member_profile in result.members.items():
                review = member_review_from_scan(member_id, member_profile)
                if saved_regions:
                    review = _apply_saved_region_choices(review, saved_regions)
                member_selectors = saved_selectors.get(member_id)
                if isinstance(member_selectors, list):
                    review = _apply_saved_selectors(review, member_selectors)
                member_renames = saved_renames.get(member_id)
                if isinstance(member_renames, dict):
                    review = dataclasses.replace(
                        review,
                        sheet_renames={str(k): str(v) for k, v in member_renames.items()},
                    )
                if pending_row_matching:
                    review = _apply_pending_row_matching(
                        review, pending_row_matching
                    )
                reviews.append(review)
            workspace_state["value"] = dataclasses.replace(
                workspace_state["value"], member_reviews=tuple(reviews)
            )
            hydrated_profile_shape["value"] = tuple(
                sorted(
                    (
                        member_id,
                        tuple(sheet.sheet_name for sheet in member.current_sheets),
                    )
                    for member_id, member in result.members.items()
                )
            )

        def _seed_deck_review() -> None:
            """Peek PowerPoint slide titles (Step 9's "cycle slide inclusion
            and explicit renamed-slide pins" / "preflight current-side
            setup" criteria) -- a cheap, direct, in-process read (unlike
            Excel's disposable-subprocess setup scan; a deck is small
            enough that no bounded worker is warranted). Runs exactly
            once: `deck_review` starts `None` and this always leaves it a
            real `DeckReview` (empty when there is no PPT file at all), so
            the `poll()` guard below never re-seeds and silently discards
            an analyst's in-progress inclusion/rename choices.
            """
            from qc_tool.io.peek import peek_slide_titles

            current_ppt = files.get(_PPT_CURRENT_PREFIX)
            if current_ppt is None:
                workspace_state["value"] = dataclasses.replace(
                    workspace_state["value"], deck_review=DeckReview()
                )
                return
            baseline_ppt = files.get(_PPT_BASELINE_PREFIX)
            baseline_titles = peek_slide_titles(baseline_ppt) if baseline_ppt is not None else []
            current_titles = peek_slide_titles(current_ppt)
            deck = deck_review_from_titles(baseline_titles, current_titles)
            saved_excluded_raw = choices.get("excluded_slides", [])
            if isinstance(saved_excluded_raw, list):
                excluded = {int(i) for i in saved_excluded_raw if str(i).lstrip("-").isdigit()}
                if excluded:
                    deck = dataclasses.replace(
                        deck,
                        current_slides=tuple(
                            dataclasses.replace(s, included=False)
                            if s.slide_index in excluded
                            else s
                            for s in deck.current_slides
                        ),
                    )
            saved_slide_renames_raw = choices.get("slide_renames", {})
            if isinstance(saved_slide_renames_raw, dict):
                deck = dataclasses.replace(
                    deck,
                    slide_renames={
                        str(k): str(v) for k, v in saved_slide_renames_raw.items()
                    },
                )
            workspace_state["value"] = dataclasses.replace(
                workspace_state["value"], deck_review=deck
            )

        def _region_by_id(region_id: str):
            for member in workspace_state["value"].member_reviews:
                for sheet in member.current_sheets:
                    for region in sheet.regions:
                        if region.region_id == region_id:
                            return member.member_id, sheet.sheet_name, region
            return None, None, None

        def _update_region(member_id: str, sheet_name: str, region_id: str, **changes) -> None:
            state = workspace_state["value"]
            member = state.member_review(member_id)
            if member is None:
                return
            updated_member = update_region_decision(member, sheet_name, region_id, **changes)
            workspace_state["value"] = state.with_member_review(updated_member)
            _persist_choices()
            refresh()

        def _apply_region_transform(
            member_id: str, sheet_name: str, region_id: str, transform
        ) -> None:
            """Backs the "click controls"/"A1 controls" criterion: a
            transform that must recompute several region fields together
            (anchor, range, available columns) instead of one field at a
            time. A `ValueError` raised by `transform` (an unparseable typed
            range) is surfaced as a notification, never a crash.
            """
            state = workspace_state["value"]
            member = state.member_review(member_id)
            if member is None:
                return
            try:
                updated_member = apply_region_transform(member, sheet_name, region_id, transform)
            except ValueError as exc:
                ui.notify(str(exc), type="warning")
                return
            workspace_state["value"] = state.with_member_review(updated_member)
            _persist_choices()
            refresh()

        def _confirm_all_regions() -> None:
            """Bulk confirmation (Step 7 criterion): mark every currently
            valid, detected region across every member as analyst-confirmed
            in one action.
            """
            state = workspace_state["value"]
            new_reviews = tuple(confirm_all_regions(member) for member in state.member_reviews)
            workspace_state["value"] = dataclasses.replace(state, member_reviews=new_reviews)
            _persist_choices()
            refresh()

        def _update_sheet_rename(
            member_id: str, current_name: str, baseline_name: str | None
        ) -> None:
            """Declare or clear a renamed-sheet pairing (Step 8's "renamed
            sheets require explicit mapping" criterion).
            """
            state = workspace_state["value"]
            member = state.member_review(member_id)
            if member is None:
                return
            updated_member = set_sheet_rename(member, current_name, baseline_name)
            workspace_state["value"] = state.with_member_review(updated_member)
            _persist_choices()
            refresh()

        def _render_sheet_rename_controls(
            member: MemberReview, added: tuple[str, ...], removed: tuple[str, ...]
        ) -> None:
            """One dropdown per added (current-only) sheet, offering every
            still-unclaimed removed (baseline-only) sheet as a rename
            target -- promotes an added+removed pair into one logical sheet
            instead of two separate acknowledgements.
            """
            claimed = set(member.sheet_renames.values())
            for current_name in added:
                declared = member.sheet_renames.get(current_name)
                options = {"": "New sheet (no rename)"}
                for baseline_name in removed:
                    if baseline_name == declared or baseline_name not in claimed:
                        options[baseline_name] = f"Same as {baseline_name!r}"

                def _on_rename_change(
                    event: events.ValueChangeEventArguments,
                    member_id=member.member_id,
                    current_name=current_name,
                ) -> None:
                    value = str(event.value or "")
                    _update_sheet_rename(member_id, current_name, value or None)

                with ui.row().classes("items-center gap-2"):
                    ui.label(f"{current_name!r} (new)").classes("note")
                    ui.select(
                        options, value=declared or "", on_change=_on_rename_change
                    ).props("outlined dense").classes("w-56")
            still_removed = [name for name in removed if name not in claimed]
            for baseline_name in still_removed:
                ui.label(f"{baseline_name!r} removed from the current file").classes("note")

        def _set_slide_included(slide_index: int, included: bool) -> None:
            state = workspace_state["value"]
            deck = state.deck_review
            if deck is None:
                return
            workspace_state["value"] = dataclasses.replace(
                state, deck_review=set_slide_included(deck, slide_index, included)
            )
            _persist_choices()
            refresh()

        def _update_slide_rename(current_title: str, baseline_title: str | None) -> None:
            """Declare or clear a renamed-slide pin (Step 9's "explicit
            renamed-slide pins" criterion) -- mirrors `_update_sheet_rename`.
            """
            state = workspace_state["value"]
            deck = state.deck_review
            if deck is None:
                return
            workspace_state["value"] = dataclasses.replace(
                state, deck_review=set_slide_rename(deck, current_title, baseline_title)
            )
            _persist_choices()
            refresh()

        def _render_slide_rename_controls(
            deck: DeckReview, added: tuple[str, ...], removed: tuple[str, ...]
        ) -> None:
            """Mirrors `_render_sheet_rename_controls` exactly, for slide
            titles instead of sheet names.
            """
            claimed = set(deck.slide_renames.values())
            for current_title in added:
                declared = deck.slide_renames.get(current_title)
                options = {"": "New slide (no rename)"}
                for baseline_title in removed:
                    if baseline_title == declared or baseline_title not in claimed:
                        options[baseline_title] = f"Same as {baseline_title!r}"

                def _on_rename_change(
                    event: events.ValueChangeEventArguments, current_title=current_title
                ) -> None:
                    value = str(event.value or "")
                    _update_slide_rename(current_title, value or None)

                with ui.row().classes("items-center gap-2"):
                    ui.label(f"{current_title!r} (new)").classes("note")
                    ui.select(
                        options, value=declared or "", on_change=_on_rename_change
                    ).props("outlined dense").classes("w-56")
            still_removed = [title for title in removed if title not in claimed]
            for baseline_title in still_removed:
                ui.label(f"{baseline_title!r} removed from the current deck").classes("note")

        def _add_selector(member_id: str, sheet_name: str, label: str, cell: str) -> None:
            state = workspace_state["value"]
            member = state.member_review(member_id)
            if member is None:
                return
            try:
                updated_member = add_selector(member, sheet_name, label=label, cell=cell)
            except ValueError as exc:
                ui.notify(str(exc), type="warning")
                return
            workspace_state["value"] = state.with_member_review(updated_member)
            _persist_choices()
            refresh()

        def _remove_selector(member_id: str, sheet_name: str, selector_id: str) -> None:
            state = workspace_state["value"]
            member = state.member_review(member_id)
            if member is None:
                return
            updated_member = remove_selector(member, sheet_name, selector_id)
            workspace_state["value"] = state.with_member_review(updated_member)
            _persist_choices()
            refresh()

        def _set_selector_baseline_cell(
            member_id: str, sheet_name: str, selector_id: str, baseline_cell: str
        ) -> None:
            state = workspace_state["value"]
            member = state.member_review(member_id)
            if member is None:
                return
            try:
                updated_member = set_selector_baseline_cell(
                    member, sheet_name, selector_id, baseline_cell
                )
            except ValueError as exc:
                ui.notify(str(exc), type="warning")
                return
            workspace_state["value"] = state.with_member_review(updated_member)
            _persist_choices()
            refresh()

        def _render_add_selector_control(
            member: MemberReview, *, initial_sheet: str | None = None
        ) -> None:
            """Add a scenario/parameter check to the active sheet."""
            sheet_names = [sheet.sheet_name for sheet in member.current_sheets]
            if not sheet_names:
                return
            selected_sheet = (
                initial_sheet if initial_sheet in sheet_names else sheet_names[0]
            )
            with ui.row().classes("items-center gap-2 flex-wrap"):
                ui.label("Add scenario or parameter check").classes("note")
                cell_input = ui.input("Cell (A1)").props("outlined dense").classes("w-32")
                label_input = ui.input("Label").props("outlined dense").classes("w-40")

                def _on_add(member_id=member.member_id) -> None:
                    _add_selector(
                        member_id,
                        selected_sheet,
                        str(label_input.value or ""),
                        str(cell_input.value or ""),
                    )
                    cell_input.set_value("")
                    label_input.set_value("")

                ui.button("Add check", on_click=_on_add).props(
                    "flat dense no-caps"
                )

        def _render_selector_row(member_id: str, sheet_name: str, selector) -> None:
            with ui.row().classes("items-center gap-2"):
                ui.label(f"{selector.label} ({selector.cell})").classes("note")

                def _on_baseline_cell_change(
                    event: events.ValueChangeEventArguments,
                    member_id=member_id,
                    sheet_name=sheet_name,
                    selector_id=selector.selector_id,
                ) -> None:
                    _set_selector_baseline_cell(
                        member_id, sheet_name, selector_id, str(event.value or "")
                    )

                ui.input(
                    "Cell in baseline (if moved)",
                    value=selector.baseline_cell,
                    on_change=_on_baseline_cell_change,
                ).props("outlined dense").classes("w-40")

                def _on_remove(
                    member_id=member_id, sheet_name=sheet_name, selector_id=selector.selector_id
                ) -> None:
                    _remove_selector(member_id, sheet_name, selector_id)

                ui.button(icon="close", on_click=_on_remove).props("flat dense round size=sm")

        def _toggle_warning(code: str, value: bool) -> None:
            state = workspace_state["value"]
            acknowledged = set(state.warnings_acknowledged)
            if value:
                acknowledged.add(code)
            else:
                acknowledged.discard(code)
            workspace_state["value"] = dataclasses.replace(
                state, warnings_acknowledged=frozenset(acknowledged)
            )
            refresh()

        def render_scan_status() -> None:
            def _start_job() -> None:
                nonlocal job
                _rebuild_passwords_by_member()
                job = coordinator.start(
                    record.session_id,
                    record.input_generation,
                    setup_member_inputs,
                    passwords_by_member=passwords_by_member,
                )

            def _retry_analysis() -> None:
                nonlocal job
                job = coordinator.restart(
                    record.session_id, record.input_generation
                )
                _start_job()
                refresh()

            def _cancel_analysis() -> None:
                coordinator.cancel(record.session_id, record.input_generation)
                ui.navigate.to(f"/?config_session={record.session_id}")

            def _discard_setup() -> None:
                nonlocal job
                cleared_choices = {
                    key: value
                    for key, value in choices.items()
                    if key
                    not in {
                        "region_decisions",
                        "selectors",
                        "sheet_renames",
                        "excluded_slides",
                        "slide_renames",
                    }
                }
                try:
                    saved = session_store.save_choices(
                        record.session_id,
                        profile_name=str(
                            choices.get("profile_name", "default")
                        ),
                        choices=cleared_choices,
                        expected_revision=session_revision["value"],
                    )
                except ConfigSessionConflictError:
                    ui.notify(
                        "This configuration changed in another tab. Reload "
                        "before discarding setup.",
                        type="warning",
                    )
                    return
                session_revision["value"] = saved.revision
                coordinator.discard(
                    record.session_id, record.input_generation
                )
                SetupScanStore(
                    work_dir / "setup-inspection.sqlite3"
                ).delete(record.session_id)
                credentials.clear()
                ui.navigate.to(f"/?config_session={record.session_id}")

            def _render_password_inputs(
                roles: tuple[str, ...],
                *,
                resume_setup: bool,
            ) -> None:
                for role in roles:
                    def _set_password(
                        event: events.ValueChangeEventArguments,
                        role: str = role,
                    ) -> None:
                        value = str(event.value or "")
                        if value:
                            credentials[role] = value
                        else:
                            credentials.pop(role, None)

                    ui.input(
                        f"Password for {role.replace('_', ' ')}",
                        password=True,
                        password_toggle_button=True,
                        on_change=_set_password,
                    ).props("outlined dense").classes("w-72")

                def _accept_credentials() -> None:
                    if any(not credentials.get(role) for role in roles):
                        ui.notify(
                            "Enter every required password before continuing.",
                            type="warning",
                        )
                        return
                    credential_prompt_prepared.clear()
                    if resume_setup:
                        _retry_analysis()
                    else:
                        _seed_member_reviews()
                        refresh()

                ui.button(
                    "Resume analysis" if resume_setup else "Use passwords",
                    on_click=_accept_credentials,
                ).classes("runbtn").props("no-caps")

            scan_status_box.clear()
            with scan_status_box:
                if not scan_members:
                    if workspace_state["value"].deck_review is not None:
                        status_chip(
                            "ok",
                            "PowerPoint inventory ready",
                            "No Excel setup scan is required for this run",
                        )
                    else:
                        ui.label("Reading PowerPoint inventory...").classes("note")
                    return
                snapshot = job.snapshot()
                status = snapshot.status
                if status in {SetupStatus.AWAITING_INPUTS, SetupStatus.STALE}:
                    ui.label("Setup analysis will start automatically.").classes("note")
                elif status is SetupStatus.AWAITING_CREDENTIALS:
                    missing_roles = _missing_password_roles()
                    if missing_roles not in credential_prompt_prepared:
                        for role in missing_roles:
                            credentials.pop(role, None)
                        credential_prompt_prepared.add(missing_roles)
                    ui.label(
                        "A password is required before setup analysis can continue."
                    ).classes("notecard")
                    _render_password_inputs(missing_roles, resume_setup=True)
                    ui.button(
                        "Back to files",
                        on_click=lambda: ui.navigate.to(
                            f"/?config_session={record.session_id}"
                        ),
                    ).props("flat no-caps")
                elif status in {
                    SetupStatus.WAITING_FOR_SLOT,
                    SetupStatus.INVENTORY_READY,
                    SetupStatus.SCANNING,
                    SetupStatus.PARTIAL_READY,
                    SetupStatus.CANCELLING,
                } and not snapshot.terminal:
                    with ui.row().classes("items-center gap-2"):
                        ui.spinner(size="1.2rem")
                        if status is SetupStatus.WAITING_FOR_SLOT:
                            detail = "Waiting for the QC worker slot..."
                        elif snapshot.total:
                            detail = (
                                f"Analyzing structure... {snapshot.processed:,} of "
                                f"{snapshot.total:,} sheet passes complete"
                            )
                        else:
                            detail = "Reading workbook inventory..."
                        ui.label(detail).classes("note")
                        ui.button(
                            "Cancel and go back", on_click=_cancel_analysis
                        ).props("flat no-caps dense")
                elif status is SetupStatus.COMPLETE:
                    result = _job_result()
                    total_regions = sum(
                        len(sheet.regions)
                        for member in result.members.values()
                        for sheet in member.current_sheets
                    ) if result else 0
                    status_chip(
                        "ok",
                        "Analysis complete",
                        f"{len(scan_members)} member(s), {total_regions} region(s) detected",
                    )
                    missing_roles = tuple(
                        role
                        for role in _missing_password_roles()
                        if not credentials.get(role)
                    )
                    if missing_roles:
                        ui.label(
                            "Re-enter the required password before running."
                        ).classes("notecard")
                        _render_password_inputs(
                            missing_roles, resume_setup=False
                        )
                elif status is SetupStatus.CANCELLED:
                    ui.label("Setup analysis was cancelled.").classes("notecard")
                    ui.button(
                        "Resume analysis",
                        on_click=_retry_analysis,
                    ).props("flat no-caps dense")
                    ui.button(
                        "Discard setup",
                        on_click=_discard_setup,
                    ).props("flat no-caps dense color=negative")
                    ui.button(
                        "Back to files",
                        on_click=lambda: ui.navigate.to(
                            f"/?config_session={record.session_id}"
                        ),
                    ).props("flat no-caps dense")
                else:
                    disclosures = "; ".join(
                        sorted(set(snapshot.member_disclosures.values()))
                    )
                    ui.label(f"Setup analysis failed: {disclosures}").classes("notecard")
                    ui.button(
                        "Retry analysis",
                        on_click=_retry_analysis,
                    ).props("flat no-caps dense")

        def render_profile_section() -> None:
            profile_box.clear()
            state = workspace_state["value"]
            with profile_box:
                if not state.member_reviews and scan_members:
                    ui.label(
                        "Choose a saved profile once the first sheet is ready."
                    ).classes("note")
                    return
                options = list_profiles(profiles_dir)
                if isinstance(raw_profile_snapshot, dict):
                    frozen_name = str(raw_profile_snapshot.get("name", ""))
                    if frozen_name and frozen_name not in options:
                        options.append(frozen_name)
                current_name = state.profile_name if state.profile_name in options else "default"

                def _on_profile_change(event: events.ValueChangeEventArguments) -> None:
                    new_name = str(event.value)
                    workspace_state["value"] = dataclasses.replace(
                        workspace_state["value"],
                        profile_name=new_name,
                        profile_opened_hash=_profile_hash_or_none(profiles_dir, new_name),
                    )
                    _persist_choices()
                    refresh()

                with ui.row().classes("items-end gap-2 flex-wrap"):
                    ui.select(
                        options,
                        value=current_name,
                        label="Deliverable profile",
                        on_change=_on_profile_change,
                    ).classes("w-64").props("outlined dense")

                    def _open_profile_policy_editor() -> None:
                        if profile_editor_open["value"]:
                            return
                        profile_editor_open["value"] = True
                        with (
                            ui.dialog().props("persistent") as profile_dialog,
                            ui.card().classes(
                                "w-[64rem] max-w-[96vw] max-h-[92vh] overflow-auto"
                            ) as profile_card,
                        ):
                            ui.label("Edit profile policy").classes("runhead")
                            with ui.row().classes("items-end gap-2 w-full no-wrap"):
                                new_profile_name = ui.input(
                                    "New profile name"
                                ).props("outlined dense").classes("flex-1")
                                ui.button(
                                    "Create",
                                    on_click=lambda: _create_profile(),
                                ).props("flat no-caps")
                            editing = ui.select(
                                list_profiles(profiles_dir),
                                value=current_name,
                                label="Profile",
                            ).props("outlined dense").classes("w-full")
                            editor_controller: ProfileEditorController | None = None

                            def _create_profile() -> None:
                                name = str(new_profile_name.value or "").strip()
                                if not name or name == "default":
                                    ui.notify(
                                        "Choose a non-empty, non-default name",
                                        type="warning",
                                    )
                                    return
                                try:
                                    path = profile_path(profiles_dir, name)
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
                                editing.options = list_profiles(profiles_dir)
                                editing.update()
                                if editor_controller is not None:
                                    editor_controller.request_load(name)
                                ui.notify(f"Profile {name!r} created")

                            def _profile_saved(_old_name: str, new_name: str) -> None:
                                editing.options = list_profiles(profiles_dir)
                                editing.value = new_name
                                editing.update()

                            def _apply_profile_to_setup() -> None:
                                nonlocal raw_profile_snapshot
                                profile_name = str(editing.value or "")
                                if not profile_name:
                                    ui.notify("Choose a saved profile first.", type="warning")
                                    return
                                raw_profile_snapshot = None
                                choices.pop("profile_snapshot", None)
                                workspace_state["value"] = dataclasses.replace(
                                    workspace_state["value"],
                                    profile_name=profile_name,
                                    profile_opened_hash=_profile_hash_or_none(
                                        profiles_dir, profile_name
                                    ),
                                )
                                _persist_choices()
                                profile_dialog.close()
                                refresh()

                            def _close_profile_editor() -> None:
                                profile_editor_open["value"] = False
                                profile_dialog.close()

                            editor_controller = open_profile_editor(
                                profile_card,
                                profiles_dir,
                                editing,
                                on_close=_close_profile_editor,
                                selected_files=lambda: dict(files),
                                selected_passwords=lambda: dict(credentials),
                                on_saved=_profile_saved,
                            )
                            ui.button(
                                "Apply saved profile to this setup",
                                icon="check",
                                on_click=_apply_profile_to_setup,
                            ).classes("runbtn").props("no-caps")
                            def _dispose_profile_editor() -> None:
                                profile_editor_open["value"] = False
                                editor_controller.dispose()

                            profile_dialog.on("hide", _dispose_profile_editor)
                        profile_dialog.open()

                    ui.button(
                        "Edit profile policy",
                        icon="tune",
                        on_click=_open_profile_policy_editor,
                    ).props("flat no-caps dense")
                try:
                    profile = _load_workspace_profile(current_name)
                except Exception:
                    profile = None
                selected_profile["value"] = profile
                result = _job_result()
                if profile is None or result is None:
                    return
                diff = diff_profile_against_scan(profile, result, mode=mode)
                if is_clean_profile_diff(diff) or not diff:
                    ui.label(
                        "This profile matches the current scan -- nothing new to "
                        "confirm."
                    ).classes("note")
                else:
                    with ui.expansion(
                        f"Configuration diff ({len(diff)} item(s))", value=False
                    ).classes("w-full"):
                        for entry in diff:
                            ui.label(f"[{entry.kind}] {entry.description}").classes("note")

        def render_regions_section() -> None:
            regions_box.clear()
            state = workspace_state["value"]
            with regions_box:
                if not state.member_reviews:
                    return
                member_ids = [member.member_id for member in state.member_reviews]
                if preview_state["member_id"] not in member_ids:
                    preview_state["member_id"] = member_ids[0]
                    preview_state["sheet_name"] = None
                    preview_state["region_id"] = None
                    preview_state["outcome"] = None
                active_member_id = str(preview_state["member_id"])
                active_member = next(
                    member
                    for member in state.member_reviews
                    if member.member_id == active_member_id
                )
                any_region = any(
                    sheet.regions
                    for member in state.member_reviews
                    for sheet in member.current_sheets
                )
                total_sheets = sum(
                    len(member.current_sheets) for member in state.member_reviews
                )
                total_regions = sum(
                    len(sheet.regions)
                    for member in state.member_reviews
                    for sheet in member.current_sheets
                )
                attention_targets = [
                    (member.member_id, sheet.sheet_name, region)
                    for member in state.member_reviews
                    for sheet in member.current_sheets
                    for region in sheet.regions
                    if (
                        region.ranked_candidate_pending
                        and region.mode == "automatic"
                    )
                    or (region.mode != "automatic" and not region.is_valid)
                ]

                def _activate_attention_target() -> None:
                    if not attention_targets:
                        return
                    active_key = (
                        preview_state["member_id"],
                        preview_state["sheet_name"],
                        preview_state["region_id"],
                    )
                    keys = [
                        (member_id, name, item.region_id)
                        for member_id, name, item in attention_targets
                    ]
                    next_index = (
                        (keys.index(active_key) + 1) % len(keys)
                        if active_key in keys
                        else 0
                    )
                    member_id, name, item = attention_targets[next_index]
                    preview_state["member_id"] = member_id
                    preview_state["sheet_name"] = name
                    preview_state["region_id"] = item.region_id
                    preview_state["anchor_target"] = None
                    preview_state["outcome"] = None
                    try:
                        row, col = parse_a1_cell(item.anchor_cell)
                    except ValueError:
                        row, col = 1, 1
                    preview_state["min_row"] = row
                    preview_state["min_col"] = col
                    render_regions_section()
                    render_preview_section()

                with ui.row().classes(
                    "config-workspace-nav items-end gap-2 flex-wrap w-full"
                ):
                    if len(member_ids) > 1:

                        def _on_active_member_change(
                            event: events.ValueChangeEventArguments,
                        ) -> None:
                            preview_state["member_id"] = str(event.value)
                            preview_state["sheet_name"] = None
                            preview_state["region_id"] = None
                            preview_state["outcome"] = None
                            render_regions_section()
                            render_preview_section()

                        ui.select(
                            member_ids,
                            value=active_member_id,
                            label="Workbook member",
                            on_change=_on_active_member_change,
                        ).props("outlined dense").classes("w-48").mark(
                            "active-member-select"
                        )

                    sheet_names = [
                        sheet.sheet_name for sheet in active_member.current_sheets
                    ]
                    if not sheet_names:
                        ui.label("No current-side sheets are available.").classes(
                            "note"
                        )
                        return
                    if preview_state["sheet_name"] not in sheet_names:
                        preview_state["sheet_name"] = sheet_names[0]
                        preview_state["region_id"] = None
                        preview_state["outcome"] = None
                    active_sheet_name = str(preview_state["sheet_name"])

                    def _on_active_sheet_change(
                        event: events.ValueChangeEventArguments,
                    ) -> None:
                        preview_state["sheet_name"] = str(event.value)
                        preview_state["region_id"] = None
                        preview_state["outcome"] = None
                        preview_state["min_row"] = 1
                        preview_state["min_col"] = 1
                        render_regions_section()
                        render_preview_section()

                    ui.select(
                        sheet_names,
                        value=active_sheet_name,
                        label="Active sheet",
                        on_change=_on_active_sheet_change,
                    ).props("outlined dense options-dense").classes(
                        "config-sheet-select"
                    ).mark("active-sheet-select")
                    ui.label(
                        f"{total_sheets} sheet(s) · {total_regions} region(s)"
                    ).classes("note")
                    if attention_targets:
                        ui.button(
                            f"Next needs attention ({len(attention_targets)})",
                            icon="error_outline",
                            on_click=_activate_attention_target,
                        ).props("flat no-caps dense color=warning").mark(
                            "next-region-attention"
                        )
                    if any_region:
                        ui.button(
                            "Confirm all detected regions",
                            on_click=lambda: _confirm_all_regions(),
                        ).props("flat no-caps dense")

                pairs, added, removed = sheet_pairing(active_member)
                if added or removed:
                    note = []
                    if pairs:
                        note.append(f"{len(pairs)} sheet(s) matched by name")
                    ui.label(
                        f"[{active_member.member_id}] " + "; ".join(note)
                    ).classes("note")
                    _render_sheet_rename_controls(active_member, added, removed)

                active_sheet = next(
                    sheet
                    for sheet in active_member.current_sheets
                    if sheet.sheet_name == active_sheet_name
                )
                visibility = (
                    "very hidden"
                    if active_sheet.very_hidden
                    else "hidden"
                    if active_sheet.hidden
                    else "visible"
                )
                with ui.element("div").classes("config-active-sheet w-full"):
                    ui.label(
                        f"{active_sheet.sheet_name} ({visibility})"
                    ).classes("runhead")
                    if active_sheet.failure_detail:
                        ui.label(active_sheet.failure_detail).classes("notecard")
                    regions = list(active_sheet.regions)
                    if regions:
                        region_ids = [region.region_id for region in regions]
                        if preview_state["region_id"] not in region_ids:
                            preview_state["region_id"] = region_ids[0]
                            preview_state["outcome"] = None
                            try:
                                row, col = parse_a1_cell(regions[0].anchor_cell)
                            except ValueError:
                                row, col = 1, 1
                            preview_state["min_row"] = row
                            preview_state["min_col"] = col
                        active_region_id = str(preview_state["region_id"])
                        active_index = region_ids.index(active_region_id)
                        region_options = {
                            region.region_id: (
                                f"Region {index + 1} of {len(regions)} · "
                                f"{region.current_range}"
                            )
                            for index, region in enumerate(regions)
                        }

                        def _activate_region(region_id: str) -> None:
                            selected = next(
                                region
                                for region in regions
                                if region.region_id == region_id
                            )
                            preview_state["region_id"] = region_id
                            preview_state["anchor_target"] = None
                            preview_state["outcome"] = None
                            try:
                                row, col = parse_a1_cell(selected.anchor_cell)
                            except ValueError:
                                row, col = 1, 1
                            preview_state["min_row"] = row
                            preview_state["min_col"] = col
                            render_regions_section()
                            render_preview_section()

                        with ui.row().classes(
                            "items-center gap-1 flex-wrap w-full"
                        ):
                            ui.button(
                                icon="chevron_left",
                                on_click=lambda: _activate_region(
                                    region_ids[(active_index - 1) % len(region_ids)]
                                ),
                            ).props("flat round dense").tooltip("Previous region")
                            ui.select(
                                region_options,
                                value=active_region_id,
                                label="Active region",
                                on_change=lambda event: _activate_region(
                                    str(event.value)
                                ),
                            ).props("outlined dense options-dense").classes(
                                "config-region-select"
                            ).mark("active-region-select")
                            ui.button(
                                icon="chevron_right",
                                on_click=lambda: _activate_region(
                                    region_ids[(active_index + 1) % len(region_ids)]
                                ),
                            ).props("flat round dense").tooltip("Next region")
                        _render_region_row(
                            active_member.member_id,
                            active_sheet.sheet_name,
                            regions[active_index],
                        )
                    else:
                        preview_state["region_id"] = None
                        ui.label(
                            "No table regions were detected on this sheet."
                        ).classes("note")
                preview_request_key = (
                    active_member.member_id,
                    active_sheet.sheet_name,
                    preview_state["region_id"],
                )
                if (
                    _setup_complete()
                    and preview_state["auto_request_key"] != preview_request_key
                ):
                    preview_state["auto_request_key"] = preview_request_key
                    _schedule_preview_load()

        def render_deck_section() -> None:
            """PowerPoint slide review (Step 9): cycle slide inclusion plus
            explicit renamed-slide pins for CYCLE_COMPARISON; a plain
            current-side inventory for preflight/final-package (no
            baseline deck, so no pairing/inclusion concept applies).
            """
            deck_box.clear()
            state = workspace_state["value"]
            deck = state.deck_review
            with deck_box:
                if deck is None or not deck.current_slides:
                    return
                ui.label("PowerPoint slides").classes("dk")
                if mode is QCRunMode.CYCLE_COMPARISON and deck.baseline_slides:
                    pairs, added, removed = effective_slide_pairing(deck)
                    if pairs:
                        ui.label(f"{len(pairs)} slide(s) matched by title").classes("note")
                    if added or removed:
                        _render_slide_rename_controls(deck, added, removed)
                for slide in deck.current_slides:
                    with ui.row().classes("items-center gap-2"):
                        if mode is QCRunMode.CYCLE_COMPARISON:
                            ui.checkbox(
                                value=slide.included,
                                on_change=(
                                    lambda event, idx=slide.slide_index: _set_slide_included(
                                        idx, bool(event.value)
                                    )
                                ),
                            ).mark(f"slide-include-{slide.slide_index}")
                        ui.label(f"{slide.slide_index}. {slide.title}").classes("note")



        def _baseline_region_for(member_id: str, sheet_name: str, region_id: str):
            member = workspace_state["value"].member_review(member_id)
            if member is None:
                return None
            current_sheet = next(
                (sheet for sheet in member.current_sheets if sheet.sheet_name == sheet_name),
                None,
            )
            if current_sheet is None:
                return None
            region_index = next(
                (
                    index
                    for index, item in enumerate(current_sheet.regions)
                    if item.region_id == region_id
                ),
                None,
            )
            if region_index is None:
                return None
            pairs, _added, _removed = effective_sheet_pairing(member)
            baseline_name = next(
                (base for base, current in pairs if current == sheet_name), None
            )
            baseline_sheet = next(
                (
                    sheet
                    for sheet in member.baseline_sheets
                    if sheet.sheet_name == baseline_name
                ),
                None,
            )
            if baseline_sheet is None or region_index >= len(baseline_sheet.regions):
                return None
            return baseline_sheet.regions[region_index]

        def _render_region_row(member_id: str, sheet_name: str, region) -> None:
            """Render one task-first region editor with progressive disclosure."""

            def _status() -> tuple[str, str, str]:
                if region.mode == "excluded":
                    return "limited", "Removed", "Excluded from this run"
                if region.ranked_candidate_pending and region.mode == "automatic":
                    return "attention", "Needs attention", "Choose how rows line up"
                if not region.is_valid:
                    return "attention", "Needs details", "Complete the selected option"
                if region.mode == "automatic":
                    return "neutral", "Automatic", "Using detected structure"
                return "ok", "Configured", "Ready for this run"

            def _set_tab(event: events.ValueChangeEventArguments) -> None:
                preview_state["editor_tab"] = str(event.value)

            def _toggle_exclusion() -> None:
                if region.mode == "excluded":
                    _update_region(
                        member_id,
                        sheet_name,
                        region.region_id,
                        mode="automatic",
                        confirmed=False,
                    )
                else:
                    _update_region(
                        member_id,
                        sheet_name,
                        region.region_id,
                        mode="excluded",
                        confirmed=True,
                    )

            with ui.element("div").classes("config-region-editor w-full").mark(
                "active-region-editor"
            ):
                with ui.row().classes(
                    "config-region-heading items-center justify-between gap-2 w-full"
                ):
                    tone, title, detail = _status()
                    status_chip(tone, title, detail)
                    ui.button(
                        (
                            "Restore to this run"
                            if region.mode == "excluded"
                            else "Remove from this run"
                        ),
                        icon=("undo" if region.mode == "excluded" else "remove_circle_outline"),
                        on_click=_toggle_exclusion,
                    ).props(
                        "flat dense no-caps "
                        + ("color=primary" if region.mode == "excluded" else "color=negative")
                    ).mark("toggle-region-exclusion")

                with ui.tabs().classes("config-region-tabs").props(
                    "dense no-caps align=left"
                ) as region_tabs:
                    ui.tab("rows", label="Rows")
                    ui.tab("bounds", label="Data bounds")
                    ui.tab("scenarios", label="Scenario checks")
                    ui.tab("advanced", label="Advanced")
                region_tabs.value = preview_state["editor_tab"]
                region_tabs.on_value_change(_set_tab)

                with ui.tab_panels(
                    region_tabs,
                    value=preview_state["editor_tab"],
                    animated=False,
                    keep_alive=True,
                ).classes("config-region-panels w-full"):
                    with ui.tab_panel("rows"):
                        if region.mode == "excluded":
                            ui.label(
                                "This table is outside this run. Record why and when "
                                "the exclusion should be reviewed."
                            ).classes("note")

                            reason_input = ui.input(
                                "Reason", value=region.exclusion_reason
                            ).props("outlined dense").classes("w-full max-w-xl")

                            def _commit_reason(_event=None) -> None:
                                _update_region(
                                    member_id,
                                    sheet_name,
                                    region.region_id,
                                    exclusion_reason=str(reason_input.value or ""),
                                )

                            reason_input.on("blur", _commit_reason)
                            reason_input.on("keydown.enter", _commit_reason)
                            ui.input(
                                "Review on",
                                value=region.exclusion_expires_on,
                                on_change=lambda event: _update_region(
                                    member_id,
                                    sheet_name,
                                    region.region_id,
                                    exclusion_expires_on=str(event.value or ""),
                                    confirmed=True,
                                ),
                            ).props("outlined dense type=date").classes("w-44")
                        else:
                            ui.label("Rows line up by").classes("dk")

                            def _on_mode_change(
                                event: events.ValueChangeEventArguments,
                            ) -> None:
                                next_mode = str(event.value)
                                _update_region(
                                    member_id,
                                    sheet_name,
                                    region.region_id,
                                    mode=next_mode,
                                    confirmed=next_mode != "automatic",
                                )

                            ui.toggle(
                                {
                                    "automatic": "Automatic",
                                    "keyed": "Matching key columns",
                                    "positional": "Row position",
                                },
                                value=region.mode,
                                on_change=_on_mode_change,
                            ).props("no-caps").mark("region-mode-toggle")
                            if region.ranked_candidate_pending and region.mode == "automatic":
                                ui.label(
                                    "This table appears sorted or ranked. Choose matching "
                                    "key columns or row position before running."
                                ).classes("notecard")
                            if region.mode == "keyed":

                                def _on_identity_change(
                                    event: events.ValueChangeEventArguments,
                                ) -> None:
                                    columns = tuple(event.value or ())
                                    _apply_region_transform(
                                        member_id,
                                        sheet_name,
                                        region.region_id,
                                        lambda item: dataclasses.replace(
                                            set_identity_columns(item, columns),
                                            confirmed=True,
                                        ),
                                    )
                                    overlap_task = asyncio.create_task(
                                        _recompute_key_overlap(
                                            member_id, sheet_name, region.region_id
                                        )
                                    )
                                    background_tasks.add(overlap_task)
                                    overlap_task.add_done_callback(
                                        background_tasks.discard
                                    )

                                ui.select(
                                    list(region.available_columns),
                                    value=list(region.identity_columns),
                                    multiple=True,
                                    label="Key columns",
                                    on_change=_on_identity_change,
                                ).props("outlined dense use-chips").classes(
                                    "w-full max-w-xl"
                                ).mark("identity-columns")
                        if not region.is_valid:
                            ui.label("More detail is required before running.").classes(
                                "notecard"
                            )

                    with ui.tab_panel("bounds"):
                        current_effective = effective_region_data_range(region)
                        baseline_region = _baseline_region_for(
                            member_id, sheet_name, region.region_id
                        )
                        baseline_outer = region.baseline_range or (
                            baseline_region.current_range
                            if baseline_region is not None
                            else region.current_range
                        )
                        baseline_effective = effective_region_data_range(
                            region, outer_range=baseline_outer
                        )
                        ui.label(
                            "Effective current data range: "
                            + (current_effective or f"automatic within {region.current_range}")
                        ).classes("note")
                        if mode is QCRunMode.CYCLE_COMPARISON:
                            ui.label(
                                "Effective baseline data range: "
                                + (baseline_effective or f"automatic within {baseline_outer}")
                            ).classes("note")

                        range_min_row, _min_col, _max_row, _max_col = parse_a1_range(
                            region.current_range
                        )
                        data_start_value = (
                            str(region.first_data_row)
                            if region.header_intent == "first_data_row"
                            and region.first_data_row is not None
                            else str(range_min_row)
                            if region.header_intent == "no_header"
                            else ""
                        )
                        data_start_input = ui.input(
                            "Data starts on row",
                            value=data_start_value,
                            placeholder="Automatic",
                        ).props("outlined dense type=number").classes("w-44")

                        def _commit_data_start(_event=None) -> None:
                            raw = str(data_start_input.value or "").strip()
                            try:
                                if raw and not raw.isdigit():
                                    raise ValueError("enter a whole row number or leave blank")
                                next_region = set_region_data_start(
                                    region, int(raw) if raw else None
                                )
                            except ValueError as exc:
                                data_start_input.error = str(exc)
                                return
                            data_start_input.error = None
                            _apply_region_transform(
                                member_id,
                                sheet_name,
                                region.region_id,
                                lambda _item: next_region,
                            )

                        data_start_input.on("blur", _commit_data_start)
                        data_start_input.on("keydown.enter", _commit_data_start)

                        footer_input = ui.input(
                            "Footer rows",
                            value=str(region.footer_rows),
                        ).props("outlined dense type=number min=0").classes("w-36")

                        def _commit_footer(_event=None) -> None:
                            raw = str(footer_input.value or "0").strip()
                            try:
                                if not raw.isdigit():
                                    raise ValueError("enter a non-negative whole number")
                                next_region = set_region_footer_rows(region, int(raw))
                            except ValueError as exc:
                                footer_input.error = str(exc)
                                return
                            footer_input.error = None
                            _apply_region_transform(
                                member_id,
                                sheet_name,
                                region.region_id,
                                lambda _item: next_region,
                            )

                        footer_input.on("blur", _commit_footer)
                        footer_input.on("keydown.enter", _commit_footer)

                    with ui.tab_panel("scenarios"):
                        member = workspace_state["value"].member_review(member_id)
                        if member is not None:
                            _render_add_selector_control(
                                member, initial_sheet=sheet_name
                            )
                            active_sheet = next(
                                (
                                    item
                                    for item in member.current_sheets
                                    if item.sheet_name == sheet_name
                                ),
                                None,
                            )
                            if active_sheet is None or not active_sheet.selectors:
                                ui.label(
                                    "No scenario or parameter checks on this sheet."
                                ).classes("note")
                            elif active_sheet is not None:
                                for selector in active_sheet.selectors:
                                    _render_selector_row(
                                        member_id, sheet_name, selector
                                    )

                    with ui.tab_panel("advanced"):
                        with ui.row().classes("items-center gap-2 flex-wrap w-full"):
                            ui.label(f"Anchor {region.anchor_cell}").classes("dk")

                            def _on_pick_anchor() -> None:
                                preview_state["member_id"] = member_id
                                preview_state["sheet_name"] = sheet_name
                                preview_state["anchor_target"] = (
                                    member_id,
                                    sheet_name,
                                    region.region_id,
                                )
                                render_preview_section()

                            is_armed = preview_state["anchor_target"] == (
                                member_id,
                                sheet_name,
                                region.region_id,
                            )
                            ui.button(
                                (
                                    "Click preview to set anchor"
                                    if is_armed
                                    else "Pick anchor from preview"
                                ),
                                on_click=_on_pick_anchor,
                            ).props(
                                "flat dense no-caps"
                                + (" color=primary" if is_armed else "")
                            )

                        current_range_input = ui.input(
                            "Current outer range (A1)", value=region.current_range
                        ).props("outlined dense").classes("w-56")

                        def _commit_current_range(_event=None) -> None:
                            text = str(current_range_input.value or "")
                            try:
                                apply_manual_range(region, text)
                            except ValueError as exc:
                                current_range_input.error = str(exc)
                                return
                            current_range_input.error = None
                            _apply_region_transform(
                                member_id,
                                sheet_name,
                                region.region_id,
                                lambda item: apply_manual_range(item, text),
                            )

                        current_range_input.on("blur", _commit_current_range)
                        current_range_input.on("keydown.enter", _commit_current_range)

                        if mode is QCRunMode.CYCLE_COMPARISON:
                            baseline_range_input = ui.input(
                                "Baseline outer range (optional)",
                                value=region.baseline_range,
                            ).props("outlined dense").classes("w-56")

                            def _commit_baseline_range(_event=None) -> None:
                                text = str(baseline_range_input.value or "").strip()
                                try:
                                    normalized = (
                                        format_a1_range(*parse_a1_range(text))
                                        if text
                                        else ""
                                    )
                                except ValueError as exc:
                                    baseline_range_input.error = str(exc)
                                    return
                                baseline_range_input.error = None
                                _update_region(
                                    member_id,
                                    sheet_name,
                                    region.region_id,
                                    baseline_range=normalized,
                                )

                            baseline_range_input.on("blur", _commit_baseline_range)
                            baseline_range_input.on(
                                "keydown.enter", _commit_baseline_range
                            )

                        if region.mode == "keyed":
                            ui.select(
                                list(region.available_columns),
                                value=list(region.ordinal_columns),
                                multiple=True,
                                label="Rank/order columns (skip value changes only)",
                                on_change=lambda event: _apply_region_transform(
                                    member_id,
                                    sheet_name,
                                    region.region_id,
                                    lambda item: set_ordinal_columns(
                                        item, tuple(event.value or ())
                                    ),
                                ),
                            ).props("outlined dense use-chips").classes(
                                "w-full max-w-xl"
                            )
                            ui.select(
                                {
                                    "skip": "Skip duplicate key groups",
                                    "occurrence": "Match by occurrence",
                                    "position": "Match remaining rows by position",
                                },
                                value=region.duplicate_key_policy,
                                label="Duplicate key handling",
                                on_change=lambda event: _update_region(
                                    member_id,
                                    sheet_name,
                                    region.region_id,
                                    duplicate_key_policy=str(event.value),
                                ),
                            ).props("outlined dense").classes("w-full max-w-xl")
                            ui.select(
                                {
                                    "system_default": "System default",
                                    "tolerate": "Tolerate blank keys",
                                    "block": "Block on blank keys",
                                },
                                value=region.blank_key_policy,
                                label="Blank key handling",
                                on_change=lambda event: _update_region(
                                    member_id,
                                    sheet_name,
                                    region.region_id,
                                    blank_key_policy=str(event.value),
                                ),
                            ).props("outlined dense").classes("w-full max-w-xl")
                            ui.checkbox(
                                "Trim outer whitespace in key columns",
                                value=region.trim_identity_whitespace,
                                on_change=lambda event: _update_region(
                                    member_id,
                                    sheet_name,
                                    region.region_id,
                                    trim_identity_whitespace=bool(event.value),
                                ),
                            )

                        with ui.expansion(
                            "Ignored and expected-refresh columns"
                        ).classes("w-full"):
                            ui.select(
                                list(region.available_columns),
                                value=list(region.ignore_columns),
                                multiple=True,
                                label="Ignore value changes in",
                                on_change=lambda event: _apply_region_transform(
                                    member_id,
                                    sheet_name,
                                    region.region_id,
                                    lambda item: set_ignore_columns(
                                        item, tuple(event.value or ())
                                    ),
                                ),
                            ).props("outlined dense use-chips").classes(
                                "w-full max-w-xl"
                            )
                            if region.ignore_columns:
                                ui.input(
                                    "Reason",
                                    value=region.ignore_columns_reason,
                                    on_change=lambda event: _update_region(
                                        member_id,
                                        sheet_name,
                                        region.region_id,
                                        ignore_columns_reason=str(event.value or ""),
                                    ),
                                ).props("outlined dense").classes("w-full max-w-xl")
                                ui.input(
                                    "Review on",
                                    value=region.ignore_columns_expires_on,
                                    on_change=lambda event: _update_region(
                                        member_id,
                                        sheet_name,
                                        region.region_id,
                                        ignore_columns_expires_on=str(
                                            event.value or ""
                                        ),
                                    ),
                                ).props("outlined dense type=date").classes("w-44")
                            ui.select(
                                list(region.available_columns),
                                value=list(region.expected_refresh_columns),
                                multiple=True,
                                label="Expected-refresh value columns",
                                on_change=lambda event: _apply_region_transform(
                                    member_id,
                                    sheet_name,
                                    region.region_id,
                                    lambda item: set_expected_refresh_columns(
                                        item, tuple(event.value or ())
                                    ),
                                ),
                            ).props("outlined dense use-chips").classes(
                                "w-full max-w-xl"
                            )

                        with ui.expansion(
                            "Column letter differs from baseline?"
                        ).classes("w-full"):
                            lettered_columns = tuple(
                                dict.fromkeys(
                                    (
                                        *region.identity_columns,
                                        *region.ordinal_columns,
                                        *region.ignore_columns,
                                        *region.expected_refresh_columns,
                                    )
                                )
                            )
                            if not lettered_columns:
                                ui.label(
                                    "Assign a column role first."
                                ).classes("note")
                            for letter in lettered_columns:
                                current_override = next(
                                    (
                                        baseline
                                        for current, baseline in region.column_baseline_letters
                                        if current == letter
                                    ),
                                    "",
                                )

                                def _on_baseline_letter_change(
                                    event: events.ValueChangeEventArguments,
                                    letter=letter,
                                    is_identity=letter in region.identity_columns,
                                ) -> None:
                                    _apply_region_transform(
                                        member_id,
                                        sheet_name,
                                        region.region_id,
                                        lambda item: set_column_baseline_letter(
                                            item, letter, str(event.value or "")
                                        ),
                                    )
                                    if is_identity:
                                        overlap_task = asyncio.create_task(
                                            _recompute_key_overlap(
                                                member_id,
                                                sheet_name,
                                                region.region_id,
                                            )
                                        )
                                        background_tasks.add(overlap_task)
                                        overlap_task.add_done_callback(
                                            background_tasks.discard
                                        )

                                ui.input(
                                    f"{letter} in baseline",
                                    value=current_override,
                                    on_change=_on_baseline_letter_change,
                                ).props("outlined dense").classes("w-32")

        def render_preview_section() -> None:
            """Current-first preview with an explicit baseline toggle (Step
            7 criterion): loads only the side/window the analyst asked for,
            never both sides, never the whole sheet.
            """
            preview_box.clear()
            state = workspace_state["value"]
            with preview_box:
                result = _job_result()
                if result is None or not any(
                    member.current_sheets for member in result.members.values()
                ):
                    ui.label(
                        "Preview becomes available when the first sheet is ready."
                    ).classes("note")
                    return
                member_ids = sorted(result.members)
                if not member_ids:
                    return
                if preview_state["member_id"] not in member_ids:
                    preview_state["member_id"] = member_ids[0]
                active_member_id = preview_state["member_id"]
                member_profile = result.members[active_member_id]
                sheet_names = [sheet.sheet_name for sheet in member_profile.current_sheets]
                if not sheet_names:
                    ui.label("No sheets available to preview.").classes("note")
                    return
                if preview_state["sheet_name"] not in sheet_names:
                    preview_state["sheet_name"] = sheet_names[0]
                    preview_state["min_row"], preview_state["min_col"] = 1, 1
                active_sheet_name = preview_state["sheet_name"]
                with ui.element("div").classes(
                    "config-preview-panel w-full"
                ).mark("active-preview-panel"):
                    ui.label(f"Preview · {active_sheet_name}").classes("runhead")
                    ui.label(
                        "Use this current/baseline window while configuring the "
                        "active region above."
                    ).classes("note")
                    with ui.row().classes(
                        "config-preview-toolbar items-center gap-2 flex-wrap w-full"
                    ):
                        if mode is QCRunMode.CYCLE_COMPARISON:

                            def _on_side_change(
                                event: events.ValueChangeEventArguments,
                            ) -> None:
                                workspace_state["value"] = dataclasses.replace(
                                    workspace_state["value"],
                                    preview_side=str(event.value),
                                )
                                preview_state["outcome"] = None
                                render_preview_section()
                                _schedule_preview_load()

                            ui.toggle(
                                {"current": "Current", "baseline": "Baseline"},
                                value=state.preview_side,
                                on_change=_on_side_change,
                            ).props("dense")

                        def _on_jump_change(
                            event: events.ValueChangeEventArguments,
                        ) -> None:
                            try:
                                row, col = parse_a1_cell(str(event.value))
                            except ValueError as exc:
                                ui.notify(str(exc), type="warning")
                                return
                            preview_state["min_row"] = row
                            preview_state["min_col"] = col
                            preview_state["outcome"] = None
                            render_preview_section()
                            _schedule_preview_load()

                        ui.input(
                            "Jump to cell",
                            value=(
                                f"{get_column_letter(preview_state['min_col'])}"
                                f"{preview_state['min_row']}"
                            ),
                            on_change=_on_jump_change,
                        ).props("outlined dense").classes("w-32")

                        def _page(delta_rows: int) -> None:
                            preview_state["min_row"] = max(
                                1, preview_state["min_row"] + delta_rows
                            )
                            preview_state["outcome"] = None
                            render_preview_section()
                            _schedule_preview_load()

                        ui.button(
                            icon="chevron_left",
                            on_click=lambda: _page(-_PREVIEW_WINDOW_ROWS),
                        ).props("flat round dense").tooltip("Previous rows")
                        ui.button(
                            icon="chevron_right",
                            on_click=lambda: _page(_PREVIEW_WINDOW_ROWS),
                        ).props("flat round dense").tooltip("Next rows")

                        def _on_reveal_change(
                            event: events.ValueChangeEventArguments,
                        ) -> None:
                            preview_state["reveal_formulas"] = bool(event.value)
                            preview_state["outcome"] = None
                            render_preview_section()
                            _schedule_preview_load()

                        active_side = (
                            state.preview_side
                            if mode is QCRunMode.CYCLE_COMPARISON
                            else "current"
                        )
                        baseline_path, current_path = members.get(
                            active_member_id, (None, None)
                        )
                        active_source_path = (
                            current_path if active_side == "current" else baseline_path
                        )
                        xlsb_preview = bool(
                            active_source_path
                            and Path(active_source_path).suffix.lower() == ".xlsb"
                        )
                        if xlsb_preview:
                            preview_state["reveal_formulas"] = False
                        formula_reveal = ui.checkbox(
                            "Reveal formula text",
                            value=preview_state["reveal_formulas"],
                            on_change=_on_reveal_change,
                        )
                        if xlsb_preview:
                            formula_reveal.disable()
                            formula_reveal.tooltip(
                                "XLSB setup previews show formula presence only; "
                                "formula text is unavailable without Office."
                            )
                        ui.button(
                            "Reload preview",
                            icon="refresh",
                            on_click=_schedule_preview_load,
                        ).props("no-caps flat dense")

                    if xlsb_preview:
                        ui.label(
                            "XLSB preview highlights formula cells, but formula text "
                            "is unavailable during Office-free setup."
                        ).classes("note")

                    if preview_state["anchor_target"] is not None:
                        _, target_sheet, target_region_id = preview_state[
                            "anchor_target"
                        ]
                        ui.label(
                            f"Click a cell below to set the anchor for "
                            f"{target_region_id!r} on {target_sheet!r}."
                        ).classes("notecard")

                    if preview_state["loading"]:
                        ui.spinner(size="1.2rem")
                        return
                    if preview_state["error"]:
                        ui.label(preview_state["error"]).classes("notecard")
                        return
                    outcome = preview_state["outcome"]
                    if outcome is None:
                        ui.label("Loading the active preview...").classes("note")
                        return
                    if outcome.missing_credential:
                        ui.label(
                            "A password is required to preview this source."
                        ).classes("notecard")
                        return
                    if outcome.disclosure:
                        ui.label(outcome.disclosure).classes("notecard")
                        return

                    def _on_cell_click(row: int, col: int) -> None:
                        target = preview_state["anchor_target"]
                        if target is None:
                            return
                        target_member_id, target_sheet, target_region_id = target
                        preview_state["anchor_target"] = None
                        _apply_region_transform(
                            target_member_id,
                            target_sheet,
                            target_region_id,
                            lambda region: apply_anchor_click(region, row, col),
                        )

                    active_region = None
                    member_review = workspace_state["value"].member_review(
                        str(active_member_id)
                    )
                    if member_review is not None:
                        active_sheet = next(
                            (
                                sheet
                                for sheet in member_review.current_sheets
                                if sheet.sheet_name == active_sheet_name
                            ),
                            None,
                        )
                        if active_sheet is not None:
                            active_region = next(
                                (
                                    item
                                    for item in active_sheet.regions
                                    if item.region_id == preview_state["region_id"]
                                ),
                                None,
                            )
                    preview_data_range = None
                    if active_region is not None:
                        if active_side == "current":
                            preview_data_range = effective_region_data_range(
                                active_region
                            )
                        else:
                            baseline_region = _baseline_region_for(
                                str(active_member_id),
                                str(active_sheet_name),
                                active_region.region_id,
                            )
                            baseline_outer = active_region.baseline_range or (
                                baseline_region.current_range
                                if baseline_region is not None
                                else active_region.current_range
                            )
                            preview_data_range = effective_region_data_range(
                                active_region, outer_range=baseline_outer
                            )
                    if preview_data_range is not None:
                        ui.label(
                            f"Effective data band: {preview_data_range}"
                        ).classes("note")
                    _render_preview_grid(
                        outcome,
                        on_cell_click=_on_cell_click,
                        effective_data_range=preview_data_range,
                    )

        async def _recompute_key_overlap(
            member_id: str, sheet_name: str, region_id: str
        ) -> None:
            """Bounded on-demand query (Step 12 Fix 5): re-open both
            sources and measure the analyst's CONFIRMED identity columns'
            overlap ratio for one region. Silent on any failure (missing
            credential, timed-out query, moved source) -- this is a
            best-effort diagnostic, never a run requirement; a failed or
            never-attempted query simply means no low-overlap warning is
            shown, not a false "clean" claim.
            """
            state = workspace_state["value"]
            member = state.member_review(member_id)
            if member is None:
                return
            sheet = next(
                (s for s in member.current_sheets if s.sheet_name == sheet_name), None
            )
            if sheet is None:
                return
            region = next((r for r in sheet.regions if r.region_id == region_id), None)
            if region is None or region.mode != "keyed" or not region.identity_columns:
                key_overlap_state["ratios"].pop(region_id, None)
                render_warnings_section()
                return
            pairs, _added, _removed = effective_sheet_pairing(member)
            baseline_by_current = {curr: base for base, curr in pairs}
            baseline_sheet_name = baseline_by_current.get(sheet_name)
            if baseline_sheet_name is None:
                return
            index = next(
                (i for i, r in enumerate(sheet.regions) if r.region_id == region_id), None
            )
            baseline_sheet = next(
                (s for s in member.baseline_sheets if s.sheet_name == baseline_sheet_name),
                None,
            )
            baseline_region = None
            if (
                baseline_sheet is not None
                and index is not None
                and index < len(baseline_sheet.regions)
            ):
                baseline_region = baseline_sheet.regions[index]
            bounds = region_key_overlap_query_bounds(region, baseline_region)
            if bounds is None:
                return
            base_first_row, base_last_row, curr_first_row, curr_last_row = bounds
            baseline_path, current_path = members.get(member_id, (None, None))
            if not baseline_path or not current_path:
                return
            baseline_role = _excel_role(
                _EXCEL_BASELINE_PREFIX, member_id, file_hashes
            )
            current_role = _excel_role(_EXCEL_CURRENT_PREFIX, member_id, file_hashes)
            baseline_letters = tuple(
                region.baseline_letter_for(letter) for letter in region.identity_columns
            )
            request = KeyOverlapRequest(
                sidecar_path=str(work_dir / "setup-inspection.sqlite3"),
                session_key=record.session_id,
                input_generation=record.input_generation,
                member_id=member_id,
                baseline_hash=file_hashes[baseline_role],
                current_hash=file_hashes[current_role],
                baseline_sheet=baseline_sheet_name,
                current_sheet=sheet_name,
                baseline_first_row=base_first_row,
                baseline_last_row=base_last_row,
                current_first_row=curr_first_row,
                current_last_row=curr_last_row,
                identity_columns=region.identity_columns,
                baseline_identity_columns=baseline_letters,
                trim_identity_whitespace=region.trim_identity_whitespace,
            )
            outcome = None
            for attempt in range(5):
                try:
                    outcome = await asyncio.to_thread(
                        run_exclusive,
                        work_dir,
                        f"key-overlap:{member_id}:{region_id}",
                        lambda: run_key_overlap_worker(request),
                    )
                    break
                except QueueBusyError:
                    if attempt == 4:
                        return
                    await asyncio.sleep(0.2)
            if outcome is None:
                return
            if outcome.ok and outcome.ratio is not None:
                key_overlap_state["ratios"][region_id] = outcome.ratio
            else:
                key_overlap_state["ratios"].pop(region_id, None)
            # Full refresh (not just the warnings section): a newly-
            # discovered low-overlap warning is now `severity="block"`, so
            # the run-action buttons' enabled state (computed inside
            # `refresh()`'s own actions_box rebuild) must be recomputed
            # too, not just the warning text.
            refresh()

        async def _load_preview(request_token: int) -> None:
            active_member_id = preview_state["member_id"]
            active_sheet_name = preview_state["sheet_name"]
            if active_member_id is None or active_sheet_name is None:
                return
            baseline_path, current_path = members.get(active_member_id, (None, None))
            active_side = (
                workspace_state["value"].preview_side
                if mode is QCRunMode.CYCLE_COMPARISON
                else "current"
            )
            source_path_str = current_path if active_side == "current" else baseline_path
            if source_path_str is None:
                if request_token != preview_state["request_token"]:
                    return
                preview_state["error"] = "No file is available for that side."
                render_preview_section()
                return
            min_row = int(preview_state["min_row"])
            min_col = int(preview_state["min_col"])
            reveal_formulas = bool(preview_state["reveal_formulas"])
            if Path(source_path_str).suffix.lower() == ".xlsb":
                reveal_formulas = False
            preview_state["loading"] = True
            preview_state["error"] = ""
            render_preview_section()
            try:
                source_role = _excel_role(
                    _EXCEL_CURRENT_PREFIX
                    if active_side == "current"
                    else _EXCEL_BASELINE_PREFIX,
                    active_member_id,
                    file_hashes,
                )
                request = PreviewWindowRequest(
                    side=active_side,
                    sidecar_path=str(work_dir / "setup-inspection.sqlite3"),
                    session_key=record.session_id,
                    input_generation=record.input_generation,
                    member_id=active_member_id,
                    source_hash=file_hashes[source_role],
                    sheet=active_sheet_name,
                    min_row=min_row,
                    min_col=min_col,
                    max_row=min_row + _PREVIEW_WINDOW_ROWS - 1,
                    max_col=min_col + _PREVIEW_WINDOW_COLS - 1,
                    reveal_formulas=reveal_formulas,
                )
                try:
                    outcome = await asyncio.to_thread(
                        run_exclusive,
                        work_dir,
                        f"setup-preview:{active_member_id}",
                        lambda: run_preview_window_worker(request),
                    )
                except QueueBusyError as exc:
                    if request_token == preview_state["request_token"]:
                        preview_state["error"] = str(exc)
                    return
                if request_token == preview_state["request_token"]:
                    preview_state["outcome"] = outcome
            finally:
                if request_token == preview_state["request_token"]:
                    preview_state["loading"] = False
                    render_preview_section()

        async def _drain_preview_requests() -> None:
            """Finish one in-flight window, then fetch only the newest context."""
            while True:
                request_token = int(preview_state["request_token"])
                await _load_preview(request_token)
                if request_token == preview_state["request_token"]:
                    return

        def render_warnings_section() -> None:
            warnings_box.clear()
            state = workspace_state["value"]
            with warnings_box:
                if not _setup_complete():
                    return
                warnings = current_warnings()
                if not warnings:
                    return
                ui.label("Warnings").classes("runhead")
                for warning in warnings:
                    with ui.row().classes("items-center gap-2"):
                        ui.checkbox(
                            warning.message,
                            value=warning.code in state.warnings_acknowledged,
                            on_change=lambda event, code=warning.code: _toggle_warning(
                                code, bool(event.value)
                            ),
                        )

        def current_warnings():
            state = workspace_state["value"]
            result = _job_result()
            warnings = (
                compute_warnings(result, state=state)
                + compute_sheet_pairing_warnings(state)
                if result is not None
                else ()
            )
            warnings += compute_key_overlap_warnings(state, key_overlap_state["ratios"])
            deck = state.deck_review
            warnings += compute_slide_pairing_warnings(deck)
            profile = selected_profile["value"]
            if profile is not None:
                warnings += compute_required_slide_warnings(
                    deck, tuple(profile.ppt.required_slides)
                )
                if mode is QCRunMode.FINAL_PACKAGE:
                    anchor_titles = tuple(
                        dict.fromkeys(m.slide for m in profile.crosscheck.mappings)
                    )
                    warnings += resolve_slide_anchors(deck, anchor_titles)
            return warnings

        def current_readiness():
            readiness = compute_workspace_readiness(
                workspace_state["value"],
                current_warnings(),
                setup_complete=_setup_complete(),
                expected_excel_member_ids=frozenset(scan_members),
                ppt_required=_PPT_CURRENT_PREFIX in files,
            )
            role_blockers = _workspace_role_blockers(mode, files)
            credential_blockers = (
                ("Re-enter required passwords before running",)
                if any(
                    not credentials.get(role)
                    for role in _missing_password_roles()
                )
                else ()
            )
            if not role_blockers and not credential_blockers:
                return readiness
            return dataclasses.replace(
                readiness,
                ready=False,
                blockers=(
                    *role_blockers,
                    *credential_blockers,
                    *readiness.blockers,
                ),
            )

        async def do_run_once() -> None:
            await _finalize(save=False, run=True, as_new_name=None)

        async def do_save_profile() -> None:
            await _finalize(save=True, run=False, as_new_name=save_as_state["name"].strip() or None)

        async def do_save_profile_and_run() -> None:
            await _finalize(
                save=True, run=True, as_new_name=save_as_state["name"].strip() or None
            )

        async def do_export_configuration() -> None:
            profile = selected_profile["value"] or _load_workspace_profile(
                workspace_state["value"].profile_name
            )
            resolved = build_resolved_configuration(
                workspace_state["value"],
                profile=profile,
                profile_sha256=profile_sha256(profile),
                warnings_acknowledged=tuple(sorted(workspace_state["value"].warnings_acknowledged)),
                file_hashes=file_hashes,
                key_overlap_ratios=key_overlap_state["ratios"],
            )
            ui.download(
                configuration_export_bytes(profile, resolved),
                filename="qc-configuration.json",
                media_type="application/json",
            )

        async def _finalize(*, save: bool, run: bool, as_new_name: str | None) -> None:
            state = workspace_state["value"]
            readiness = current_readiness()
            if not readiness.ready:
                ui.notify("; ".join(readiness.blockers), type="warning")
                return
            try:
                profile = _load_workspace_profile(state.profile_name)
            except Exception as exc:
                ui.notify(f"Profile failed to load: {exc}", type="negative")
                return
            if save:
                contract = build_input_contract(state)
                target_name = as_new_name or (
                    state.profile_name if state.profile_name != "default" else ""
                )
                if not target_name:
                    ui.notify(
                        "Choose a profile name before saving (the built-in "
                        "default profile cannot be overwritten).",
                        type="warning",
                    )
                    return
                if target_name == state.profile_name:
                    # Updating the same profile this session opened --
                    # Step 11's "optimistic conflict protection" criterion:
                    # refuse rather than silently clobber a concurrent edit
                    # from another tab/session (mirrors the standalone
                    # profile editor's own `save_draft` staleness check).
                    current_hash = _profile_hash_or_none(profiles_dir, target_name)
                    if current_hash != state.profile_opened_hash:
                        ui.notify(
                            f"Profile {target_name!r} changed elsewhere "
                            "while this workspace was open -- reopen it to "
                            "see the latest version before saving.",
                            type="negative",
                        )
                        return
                elif profile_path(profiles_dir, target_name).exists():
                    ui.notify(
                        f"Profile {target_name!r} already exists -- choose "
                        "a different name.",
                        type="negative",
                    )
                    return
                saved_profile = profile.model_copy(
                    update={"name": target_name, "input_contract": contract}
                )
                try:
                    save_profile(saved_profile, profile_path(profiles_dir, target_name))
                except Exception as exc:
                    ui.notify(f"Profile could not be saved: {exc}", type="negative")
                    return
                profile = saved_profile
                workspace_state["value"] = dataclasses.replace(
                    state,
                    profile_name=target_name,
                    profile_opened_hash=_profile_hash_or_none(profiles_dir, target_name),
                )
                ui.notify(f"Profile {target_name!r} saved")
            if not run:
                return

            async def _submit_run() -> None:
                resolved = build_resolved_configuration(
                    workspace_state["value"],
                    profile=profile,
                    profile_sha256=profile_sha256(profile),
                    warnings_acknowledged=tuple(
                        sorted(workspace_state["value"].warnings_acknowledged)
                    ),
                    file_hashes=file_hashes,
                    key_overlap_ratios=key_overlap_state["ratios"],
                )
                try:
                    manifest = PackageManifest.from_role_files(
                        {role: str(path) for role, path in files.items()}
                    )
                except ValueError as exc:
                    ui.notify(str(exc), type="negative")
                    return
                from qc_tool.ui.app import SessionState, build_run_request

                live_state = workspace_state["value"]
                deck = live_state.deck_review
                available_slides = (
                    [(s.slide_index, s.title) for s in deck.current_slides]
                    if deck is not None
                    else []
                )
                selected_slides = (
                    {s.slide_index for s in deck.current_slides if s.included}
                    if deck is not None
                    else set()
                )
                run_state = SessionState(
                    files=files,
                    file_hashes=file_hashes,
                    profile_name=profile.name,
                    mode=mode,
                    output_mode=_output_mode_from_choices(choices),
                    allow_large_workbooks=live_state.allow_large_workbooks,
                    allow_dependency_indexing=live_state.allow_dependency_indexing,
                    acceptance_absolute=_coerce_float(
                        choices.get("acceptance_absolute"), 0.0
                    ),
                    acceptance_percent=_coerce_float(
                        choices.get("acceptance_percent"), 0.0
                    ),
                    rerun_of=_coerce_optional_int(choices.get("rerun_of")),
                    available_slides=available_slides,
                    selected_slides=selected_slides,
                )
                request = build_run_request(
                    work_dir=work_dir,
                    state=run_state,
                    files=files,
                    profile=profile,
                    manifest=manifest,
                    resolved_input_configuration=resolved.model_dump(mode="json"),
                    resolved_input_digest=resolved.canonical_sha256(),
                )
                try:
                    queue_manager.submit(request, dict(credentials))
                except QueueBusyError as exc:
                    ui.notify(str(exc), type="warning")
                    return
                ui.notify("Run started")
                ui.navigate.to("/")

            # Volume-projection courtesy warning (mirrors the main page's
            # own pre-wizard check, ported here so it is not silently lost
            # now that every fresh submission enters this workspace first).
            # Scoped to the same single-workbook cycle case the original
            # check covered; a package/preflight/final-package submission
            # skips straight to _submit_run().
            acknowledged_large_volume = (
                "large_comparison_confirmed"
                in workspace_state["value"].warnings_acknowledged
            )
            if (
                mode is QCRunMode.CYCLE_COMPARISON
                and "baseline_excel" in files
                and "current_excel" in files
                and not acknowledged_large_volume
            ):
                projection = await asyncio.to_thread(
                    project_cycle_volume,
                    files["baseline_excel"],
                    files["current_excel"],
                )
                if (
                    projection is not None
                    and projection.projected_max_findings > REPORT_DEFER_FINDINGS
                ):

                    async def _run_anyway() -> None:
                        workspace_state["value"] = dataclasses.replace(
                            workspace_state["value"],
                            warnings_acknowledged=frozenset(
                                workspace_state["value"].warnings_acknowledged
                            )
                            | {"large_comparison_confirmed"},
                        )
                        volume_dialog.close()
                        await _submit_run()

                    with (
                        ui.dialog().props("persistent") as volume_dialog,
                        ui.card().classes("w-[36rem] max-w-full"),
                    ):
                        ui.label(
                            "This looks like a very large comparison"
                        ).classes("runhead")
                        ui.label(
                            f"Up to ~{projection.projected_max_findings:,} "
                            f"findings across {len(projection.changed_sheets)} "
                            "changed sheet"
                            f"{'s' if len(projection.changed_sheets) != 1 else ''}. "
                            "Narrow the region/sheet choices above to reduce "
                            "this, or continue -- a full run remains the "
                            "sign-off artifact."
                        ).classes("notecard")
                        with ui.row().classes("items-center gap-2"):
                            ui.button(
                                "Run anyway", on_click=_run_anyway
                            ).classes("runbtn").props("no-caps")
                            ui.button(
                                "Cancel", on_click=volume_dialog.close
                            ).props("flat no-caps")
                    volume_dialog.open()
                    return
            await _submit_run()

        def _set_allow_large_workbooks(value: bool) -> None:
            workspace_state["value"] = dataclasses.replace(
                workspace_state["value"], allow_large_workbooks=value
            )
            _persist_choices()

        def _set_allow_dependency_indexing(value: bool) -> None:
            workspace_state["value"] = dataclasses.replace(
                workspace_state["value"], allow_dependency_indexing=value
            )
            _persist_choices()

        def render_advanced_section() -> None:
            """Compact tolerance/materiality/waiver + availability summary
            (Step 9) plus run-only workload/dependency safety overrides.
            The full profile editor remains the place to actually CHANGE
            tolerance/waiver/availability policy -- this is read-only.
            """
            advanced_box.clear()
            state = workspace_state["value"]
            with advanced_box:
                if _setup_complete():
                    profile = selected_profile["value"]
                    if profile is not None:
                        with ui.expansion(
                            "Advanced controls & policy summary", value=False
                        ).classes("w-full"):
                            ui.label(
                                f"Tolerance: ±{profile.tolerance.absolute:g} "
                                f"absolute, ±{profile.tolerance.relative:g} "
                                "relative"
                            ).classes("note")
                            ui.label(
                                f"{len(profile.waivers)} waiver(s), "
                                f"{len(profile.materiality_severity)} materiality "
                                "severity override(s)"
                            ).classes("note")
                            ui.label(
                                f"{len(profile.ppt.availability_rules)} PowerPoint "
                                "availability rule(s)"
                            ).classes("note")
                            ui.label(
                                "Use Edit profile policy above to change these settings "
                                "without leaving this setup."
                            ).classes("note")
                with ui.row().classes("items-center gap-2"):
                    ui.checkbox(
                        "Allow large workbooks (override the workload safety gate)",
                        value=state.allow_large_workbooks,
                        on_change=lambda event: _set_allow_large_workbooks(bool(event.value)),
                    )
                with ui.row().classes("items-center gap-2"):
                    ui.checkbox(
                        "Force full dependency indexing (override the size gate)",
                        value=state.allow_dependency_indexing,
                        on_change=lambda event: _set_allow_dependency_indexing(bool(event.value)),
                    )

        def refresh() -> None:
            render_scan_status()
            render_profile_section()
            render_regions_section()
            render_deck_section()
            render_preview_section()
            render_warnings_section()
            render_advanced_section()
            actions_box.clear()
            with actions_box:
                state = workspace_state["value"]
                readiness = current_readiness()
                enabled = readiness.ready

                def _on_save_as_change(event: events.ValueChangeEventArguments) -> None:
                    save_as_state["name"] = str(event.value or "")

                if _setup_complete():
                    ui.input(
                        "Save as profile name",
                        value=save_as_state["name"]
                        or (state.profile_name if state.profile_name != "default" else ""),
                        on_change=_on_save_as_change,
                    ).props("outlined dense").classes("w-64")
                ui.button("Run once", on_click=do_run_once).classes("runbtn").props(
                    "no-caps"
                ).set_enabled(enabled)
                ui.button("Save profile", on_click=do_save_profile).classes("ghostbtn").props(
                    "no-caps flat"
                ).set_enabled(enabled)
                ui.button(
                    "Save profile and run", on_click=do_save_profile_and_run
                ).classes("ghostbtn").props("no-caps flat").set_enabled(enabled)
                ui.button(
                    "Update profile and run", on_click=do_save_profile_and_run
                ).classes("ghostbtn").props("no-caps flat").set_enabled(
                    enabled and workspace_state["value"].profile_name != "default"
                )
                ui.button(
                    "Export configuration", on_click=do_export_configuration
                ).classes("ghostbtn").props("no-caps flat").set_enabled(
                    enabled
                )
                if readiness.blockers:
                    ui.label("; ".join(readiness.blockers)).classes("notecard")

        status = job.snapshot().status
        if scan_members and (
            status
            in {
                SetupStatus.AWAITING_INPUTS,
                SetupStatus.STALE,
                SetupStatus.COMPLETE,
            }
            or (
                status is SetupStatus.AWAITING_CREDENTIALS
                and bool(credentials)
            )
        ):
            job = coordinator.start(
                record.session_id,
                record.input_generation,
                setup_member_inputs,
                passwords_by_member=passwords_by_member,
            )
        if _job_result() is not None and not workspace_state["value"].member_reviews:
            _seed_member_reviews()
        if not scan_members and workspace_state["value"].deck_review is None:
            _seed_deck_review()

        last_poll_lifecycle: dict[
            str, tuple[SetupStatus, str, int, int, bool] | None
        ] = {"value": None}

        def poll() -> None:
            coordinator.evict_terminal(max_age_seconds=60.0)
            seeded = False
            snapshot = job.snapshot()
            lifecycle = (
                snapshot.status,
                snapshot.phase,
                snapshot.processed,
                snapshot.total,
                snapshot.terminal,
            )
            lifecycle_changed = lifecycle != last_poll_lifecycle["value"]
            previous_status = (
                last_poll_lifecycle["value"][0]
                if last_poll_lifecycle["value"] is not None
                else None
            )
            last_poll_lifecycle["value"] = lifecycle
            result = _job_result()
            result_shape = (
                tuple(
                    sorted(
                        (
                            member_id,
                            tuple(
                                sheet.sheet_name
                                for sheet in member.current_sheets
                            ),
                        )
                        for member_id, member in result.members.items()
                    )
                )
                if result is not None
                else ()
            )
            if result is not None and result_shape != hydrated_profile_shape["value"]:
                _seed_member_reviews()
                seeded = True
            if (
                (not scan_members or snapshot.status is SetupStatus.COMPLETE)
                and workspace_state["value"].deck_review is None
            ):
                _seed_deck_review()
                seeded = True
            terminal_transition = (
                previous_status is not snapshot.status
                and snapshot.status
                in {
                    SetupStatus.COMPLETE,
                    SetupStatus.CANCELLED,
                    SetupStatus.FAILED,
                }
            )
            if seeded or terminal_transition:
                refresh()
            elif lifecycle_changed:
                render_scan_status()

        refresh()
        initial_snapshot = job.snapshot()
        last_poll_lifecycle["value"] = (
            initial_snapshot.status,
            initial_snapshot.phase,
            initial_snapshot.processed,
            initial_snapshot.total,
            initial_snapshot.terminal,
        )
        poll_timer = ui.timer(0.5, poll)
        cleaned_up = False

        def _cleanup_workspace() -> None:
            nonlocal cleaned_up
            if cleaned_up:
                return
            cleaned_up = True
            poll_timer.cancel()
            for task in tuple(background_tasks):
                task.cancel()
            background_tasks.clear()
            credentials.clear()
            coordinator.detach(
                record.session_id, record.input_generation, subscriber_id
            )

        client.on_disconnect(_cleanup_workspace)
        client.on_delete(_cleanup_workspace)


def _output_mode_from_choices(choices: dict[str, object]):
    from qc_tool.coverage import FindingOutputMode

    try:
        return FindingOutputMode(str(choices.get("output_mode", "decision")))
    except ValueError:
        return FindingOutputMode.DECISION
