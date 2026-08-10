"""Command-line interface: serve (default), run, sanitize, lint.

`cadence-diff` with no subcommand serves the web UI. The legacy `qc-tool`
entry point remains supported.
`qc-tool run` executes headless QC with CI-style exit codes, `sanitize`
supports local numeric scrambling or strict verified redaction,
`fingerprint` emits structural-only JSON, and `lint` validates a profile.

Exit codes: 0 ok · 1 usage/runtime error · 2 findings at/above the
--fail-on threshold (run) or lint errors (lint).
"""

import argparse
import getpass
import logging
import os
import re
import sys
from pathlib import Path

from platformdirs import user_data_dir

from qc_tool import __version__
from qc_tool.package import MEMBER_ID_PATTERN, PackageManifest
from qc_tool.security import private_directory

_SUBCOMMANDS = {
    "serve",
    "launch",
    "shortcut",
    "run",
    "sanitize",
    "sanitize-package",
    "verify-sanitized",
    "fingerprint",
    "verify-attestation",
    "network",
    "lint",
}
_MODE_ALIASES = {
    "cycle": "cycle_comparison",
    "preflight": "current_file_preflight",
    "package": "final_package",
}
_ROLES = ("baseline_excel", "current_excel", "baseline_ppt", "current_ppt")


def _program_name() -> str:
    executable = Path(sys.argv[0]).name.casefold()
    return "qc-tool" if executable.startswith("qc-tool") else "cadence-diff"


def default_data_dir() -> Path:
    """Per-user writable location for uploads, profiles, history, reports."""
    return Path(user_data_dir("qc-tool", appauthor=False))


# --- serve -------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=_program_name(),
        description=(
            "Local, read-only QC for cadence Excel/PowerPoint deliverables: "
            "current-file preflight, baseline/current cycle comparison, and "
            "final-package reconciliation. Serves a web UI on 127.0.0.1. "
            "Subcommands: run (headless QC), sanitize (shareable scrambled "
            "or verified-redacted copies), fingerprint (structural-only JSON), "
            "lint (profile validation)."
        ),
    )
    parser.add_argument(
        "--port", type=int, default=8080, help="localhost port (default: 8080)"
    )
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=None,
        help=(
            "directory for uploads, profiles, run history, and reports "
            f"(default: {default_data_dir()})"
        ),
    )
    parser.add_argument(
        "--version", action="version", version=f"%(prog)s {__version__}"
    )
    parser.add_argument(
        "--network",
        choices=["local", "lan"],
        default=None,
        help="process exposure override; LAN requires --expose-for",
    )
    parser.add_argument(
        "--expose-for",
        type=int,
        default=None,
        metavar="MINUTES",
        help="temporary LAN duration (1-1440 minutes)",
    )
    parser.add_argument(
        "--desktop-focus",
        action="store_true",
        help=(
            "opt in to jumping from a finding to the same location in an "
            "already-open Windows Excel or PowerPoint document (Windows and "
            "loopback only; off by default)"
        ),
    )
    parser.add_argument("--no-browser", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--launcher-child", action="store_true", help=argparse.SUPPRESS)
    return parser


def _cmd_serve(args: list[str]) -> int:
    ns = build_parser().parse_args(args)
    data_dir = ns.data_dir or default_data_dir()
    data_dir.mkdir(parents=True, exist_ok=True)
    if ns.launcher_child:
        from qc_tool.launcher import configure_launcher_logging

        configure_launcher_logging(data_dir)
    from qc_tool.server_config import (
        load_server_config,
        local_config,
        save_server_config,
        temporary_lan_config,
    )

    stored_config = load_server_config(data_dir)
    if ns.network == "lan":
        if ns.expose_for is None:
            raise ValueError("--network lan requires --expose-for MINUTES")
        config = temporary_lan_config(
            ns.expose_for,
            desktop_focus=stored_config.desktop_focus,
        )
        save_server_config(data_dir, config)
    elif ns.network == "local":
        if ns.expose_for is not None:
            raise ValueError("--expose-for only applies to --network lan")
        config = local_config(desktop_focus=stored_config.desktop_focus)
        save_server_config(data_dir, config)
    else:
        if ns.expose_for is not None:
            raise ValueError("--expose-for requires --network lan")
        config = stored_config
    try:
        from qc_tool.ui.app import run_app  # deferred: keep --help/--version instant

        run_app(
            data_dir,
            port=ns.port,
            host=config.host,
            network_mode=config.network,
            expires_at=config.expires_at,
            desktop_focus=ns.desktop_focus or config.desktop_focus,
            show=not ns.no_browser,
        )
    except Exception as exc:
        if ns.launcher_child:
            logging.getLogger(__name__).error(
                "server-start-failed %s",
                type(exc).__name__,
            )
        raise
    return 0


def _launch_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=f"{_program_name()} launch",
        description=(
            "Open an authenticated running local QC Tool instance or start this "
            "Python environment quietly and wait for it to become ready."
        ),
    )
    parser.add_argument("--port", type=int, default=8080)
    parser.add_argument("--data-dir", type=Path, default=None)
    return parser


def _cmd_launch(args: list[str]) -> int:
    ns = _launch_parser().parse_args(args)
    data_dir = ns.data_dir or default_data_dir()
    from qc_tool.launcher import (
        LaunchOutcome,
        configure_launcher_logging,
        launch_local_app,
    )

    configure_launcher_logging(data_dir)
    result = launch_local_app(data_dir, port=ns.port)
    print(f"launch: {result.outcome.value} {result.url}")
    if result.outcome in {
        LaunchOutcome.STARTED,
        LaunchOutcome.REUSED,
        LaunchOutcome.IN_PROGRESS,
    }:
        return 0
    logging.getLogger(__name__).error("launch-%s", result.outcome.value)
    if result.detail:
        print(f"error: {result.detail}", file=sys.stderr)
    return 1


def _shortcut_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=f"{_program_name()} shortcut",
        description="Install, inspect, or remove the per-user Windows desktop shortcut.",
    )
    parser.add_argument("action", choices=["install", "status", "remove"])
    parser.add_argument("--port", type=int, default=8080)
    parser.add_argument("--data-dir", type=Path, default=None)
    return parser


def _cmd_shortcut(args: list[str]) -> int:
    ns = _shortcut_parser().parse_args(args)
    data_dir = ns.data_dir or default_data_dir()
    from qc_tool.shortcut import (
        ShortcutState,
        install_shortcut,
        remove_shortcut,
        shortcut_status,
    )

    operations = {
        "install": install_shortcut,
        "status": shortcut_status,
        "remove": remove_shortcut,
    }
    result = operations[ns.action](data_dir, port=ns.port)
    print(f"shortcut: {result.state.value}")
    if result.detail:
        print(result.detail)
    if ns.action == "remove":
        return 0 if result.state is ShortcutState.MISSING else 1
    return 0 if result.state is ShortcutState.INSTALLED else 1


# --- run ---------------------------------------------------------------------


def _run_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=f"{_program_name()} run",
        description=(
            "Headless QC run. Mode is inferred from the supplied files "
            "(any baseline -> cycle; current Excel+PPT -> package; a single "
            "current file -> preflight) unless --mode is given."
        ),
    )
    for role in _ROLES:
        parser.add_argument(
            f"--{role.replace('_', '-')}", type=Path, default=None, metavar="FILE"
        )
    parser.add_argument(
        "--baseline-workbook",
        action="append",
        default=[],
        metavar="MEMBER=FILE",
        help="add a baseline Excel package member (repeatable; max eight)",
    )
    parser.add_argument(
        "--current-workbook",
        action="append",
        default=[],
        metavar="MEMBER=FILE",
        help="add a current Excel package member (repeatable; max eight)",
    )
    parser.add_argument(
        "--mode",
        choices=[*_MODE_ALIASES, *_MODE_ALIASES.values()],
        default=None,
        help="cycle | preflight | package (default: inferred)",
    )
    parser.add_argument(
        "--profile",
        default=None,
        metavar="NAME_OR_PATH",
        help="profile name (resolved in <data-dir>/profiles) or a YAML path",
    )
    parser.add_argument("--data-dir", type=Path, default=None)
    parser.add_argument(
        "--json", type=Path, default=None, help="write machine-readable findings JSON"
    )
    parser.add_argument(
        "--json-context",
        action="store_true",
        help=(
            "include raw cell-neighborhood excerpts and mapping candidate values in JSON; "
            "private/sensitive, disabled by default"
        ),
    )
    parser.add_argument(
        "--password",
        action="append",
        default=[],
        metavar="ROLE_OR_FILENAME=PW",
        help=(
            "INSECURE: inline password, visible in process lists and shell history; "
            "prefer --password-env, --password-file, or --password-prompt"
        ),
    )
    parser.add_argument(
        "--password-env",
        action="append",
        default=[],
        metavar="ROLE_OR_FILENAME=ENV_VAR",
        help="read a password from an environment variable (repeatable)",
    )
    parser.add_argument(
        "--password-file",
        action="append",
        default=[],
        metavar="ROLE_OR_FILENAME=PATH",
        help="read a password from a private text file (repeatable)",
    )
    parser.add_argument(
        "--password-prompt",
        action="append",
        default=[],
        metavar="ROLE_OR_FILENAME",
        help="prompt securely for a password (repeatable)",
    )
    parser.add_argument(
        "--fail-on",
        choices=["critical", "warning", "never"],
        default="critical",
        help="exit 2 when findings at/above this severity exist (default: critical)",
    )
    parser.add_argument("--top", type=int, default=10, help="findings to print")
    parser.add_argument(
        "--individual-findings",
        action="store_true",
        help="print atomic findings instead of grouped review items",
    )
    parser.add_argument(
        "--json-review-summary",
        action="store_true",
        help=(
            "add a versioned semantic review summary to the JSON export; "
            "atomic findings remain unchanged"
        ),
    )
    parser.add_argument(
        "--allow-large-workbooks",
        action="store_true",
        help=(
            "override physical-size and formula-link workload refusals for all "
            "workbooks in this run; requires sufficient local memory and "
            "degrades workload coverage"
        ),
    )
    parser.add_argument(
        "--accept-absolute",
        type=float,
        default=0.0,
        metavar="VALUE",
        help=(
            "analyst acceptance threshold: numeric differences within this "
            "absolute value report as within-tolerance Info in cycle comparisons "
            "(default 0 = off)"
        ),
    )
    parser.add_argument(
        "--accept-percent",
        type=float,
        default=0.0,
        metavar="PCT",
        help=(
            "analyst acceptance threshold as a percentage, e.g. 0.1 for 0.1%%; "
            "differences within either bound report as within-tolerance Info "
            "in cycle comparisons (default 0 = off)"
        ),
    )
    parser.add_argument(
        "--sheets",
        default="",
        metavar="NAMES",
        help=(
            "comma-separated Excel sheet names to compare; other sheets "
            "produce no findings (files still load fully; disclosed)"
        ),
    )
    parser.add_argument(
        "--member-sheets",
        action="append",
        default=[],
        metavar="MEMBER=NAMES",
        help=(
            "comma-separated sheet scope for one current workbook member "
            "(repeatable)"
        ),
    )
    parser.add_argument(
        "--slides",
        default="",
        metavar="INDEXES",
        help=(
            "comma-separated 1-based PPT slide numbers or ranges (e.g. 1,3-5) "
            "to compare; deck-level findings always report"
        ),
    )
    parser.add_argument(
        "--progress",
        action="store_true",
        help="write phase progress to stderr",
    )
    parser.add_argument(
        "--attestation",
        type=Path,
        default=None,
        help="write an HMAC-signed QC evidence bundle",
    )
    parser.add_argument(
        "--attestation-key-file",
        type=Path,
        default=None,
        help="private HMAC key (default: <data-dir>/attestation.key)",
    )
    return parser


def _infer_mode(files: dict[str, Path]) -> str:
    if any(
        key.startswith("baseline_excel") or key == "baseline_ppt"
        for key in files
    ):
        return "cycle_comparison"
    if any(key.startswith("current_excel") for key in files) and (
        "current_ppt" in files
    ):
        return "final_package"
    return "current_file_preflight"


def _resolve_profile(spec: str | None, data_dir: Path):
    from qc_tool.config.profile import (
        default_profile,
        load_profile,
        load_profile_by_name,
    )

    if spec is None:
        return default_profile()
    path = Path(spec)
    if path.suffix in {".yaml", ".yml"} or path.exists():
        return load_profile(path)
    return load_profile_by_name(data_dir / "profiles", spec)


def _parse_passwords(pairs: list[str], files: dict[str, Path]) -> dict[str, str]:
    passwords: dict[str, str] = {}
    for pair in pairs:
        key, sep, password = pair.partition("=")
        if not sep:
            raise ValueError(f"--password expects ROLE_OR_FILENAME=PW, got {pair!r}")
        for role in _password_roles(key, files):
            passwords[role] = password
    return passwords


def _password_roles(key: str, files: dict[str, Path]) -> list[str]:
    if key in files:
        return [key]
    matches = [role for role, path in files.items() if path.name == key]
    if not matches:
        raise ValueError(f"password key {key!r} matches no supplied file or role")
    if len(matches) > 1:
        raise ValueError(
            f"password filename {key!r} is ambiguous; use an exact role key"
        )
    return matches


def _parse_password_sources(ns: argparse.Namespace, files: dict[str, Path]) -> dict[str, str]:
    passwords = _parse_passwords(ns.password, files)
    for pair in ns.password_env:
        key, sep, variable = pair.partition("=")
        if not sep:
            raise ValueError(f"--password-env expects ROLE_OR_FILENAME=ENV_VAR, got {pair!r}")
        if variable not in os.environ:
            raise ValueError(f"environment variable {variable!r} is not set")
        for role in _password_roles(key, files):
            passwords[role] = os.environ[variable]
    for pair in ns.password_file:
        key, sep, raw_path = pair.partition("=")
        if not sep:
            raise ValueError(f"--password-file expects ROLE_OR_FILENAME=PATH, got {pair!r}")
        path = Path(raw_path)
        password = path.read_text(encoding="utf-8").rstrip("\r\n")
        if not password:
            raise ValueError(f"password file {path} is empty")
        if os.name == "posix" and path.stat().st_mode & 0o077:
            raise ValueError(f"password file {path} must not be group/world accessible")
        for role in _password_roles(key, files):
            passwords[role] = password
    for key in ns.password_prompt:
        password = getpass.getpass(f"Password for {key}: ")
        if not password:
            raise ValueError(f"password for {key!r} is empty")
        for role in _password_roles(key, files):
            passwords[role] = password
    return passwords


def _parse_sheet_list(raw: str) -> list[str] | None:
    names = [part.strip() for part in raw.split(",") if part.strip()]
    return names or None


def _parse_slide_list(raw: str) -> list[int] | None:
    """Parse '1,3-5' style 1-based slide selections."""
    indexes: set[int] = set()
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        start, sep, end = part.partition("-")
        try:
            if sep:
                first, last = int(start), int(end)
            else:
                first = last = int(part)
        except ValueError as exc:
            raise ValueError(f"invalid slide selection {part!r}") from exc
        if first < 1 or last < first:
            raise ValueError(f"invalid slide selection {part!r}")
        indexes.update(range(first, last + 1))
    return sorted(indexes) or None


def _cmd_run(args: list[str]) -> int:
    ns = _run_parser().parse_args(args)
    if ns.accept_absolute < 0 or ns.accept_percent < 0:
        raise ValueError("acceptance thresholds cannot be negative")
    files = {
        role: getattr(ns, role)
        for role in _ROLES
        if getattr(ns, role) is not None
    }
    def _add_member_workbooks(pairs: list[str], prefix: str) -> None:
        for pair in pairs:
            key, sep, raw_path = pair.partition("=")
            if not sep:
                raise ValueError(f"expected MEMBER=PATH, got {pair!r}")
            member = key.strip()
            if member == "primary":
                raise ValueError(
                    "member id 'primary' uses --baseline-excel/--current-excel"
                )
            if re.fullmatch(MEMBER_ID_PATTERN, member) is None:
                raise ValueError(f"invalid member id {member!r}")
            path = Path(raw_path)
            if not path.exists():
                raise ValueError(f"member workbook {path} not found")
            role = f"{prefix}:{member}"
            if role in files:
                raise ValueError(f"duplicate member role {role}")
            files[role] = path

    _add_member_workbooks(ns.baseline_workbook, "baseline_excel")
    _add_member_workbooks(ns.current_workbook, "current_excel")

    # Parse member-qualified sheet selections
    compare_member_sheets: dict[str, tuple[str, ...]] = {}
    for pair in ns.member_sheets:
        key, sep, raw_names = pair.partition("=")
        if not sep:
            raise ValueError(f"expected MEMBER=NAMES, got {pair!r}")
        member = key.strip()
        if re.fullmatch(MEMBER_ID_PATTERN, member) is None:
            raise ValueError(f"invalid member id {member!r}")
        if member in compare_member_sheets:
            raise ValueError(f"duplicate member sheets for {member}")
        names = [n.strip() for n in raw_names.split(",") if n.strip()]
        if not names:
            raise ValueError(f"no sheet names supplied for {member}")
        compare_member_sheets[member] = tuple(names)
    if ns.sheets and compare_member_sheets:
        raise ValueError("--sheets cannot be combined with --member-sheets")
    if not files:
        raise ValueError(
            f"{_program_name()} run: supply at least one file (see --help)"
        )
    for role, path in files.items():
        if not path.exists():
            raise ValueError(f"{role}: {path} not found")

    from qc_tool.coverage import QCRunMode, capability_limited
    from qc_tool.findings import Severity
    from qc_tool.progress import ProgressEvent
    from qc_tool.report.json_report import write_json_report
    from qc_tool.review_stream import (
        counts_from_summaries,
        summarize_pattern_groups,
        summarize_review_groups,
    )
    from qc_tool.run_service import perform_run

    data_dir = ns.data_dir or default_data_dir()
    private_directory(data_dir)
    mode_value = _MODE_ALIASES.get(ns.mode, ns.mode) if ns.mode else _infer_mode(files)
    profile = _resolve_profile(ns.profile, data_dir)
    passwords = _parse_password_sources(ns, files)

    on_progress = None
    if ns.progress:

        def print_progress(event: ProgressEvent) -> None:
            label = event.phase.value.replace("_", " ")
            counts = (
                f" ({event.processed}/{event.total})" if event.total else ""
            )
            detail = f" - {event.detail}" if event.detail else ""
            print(f"progress: {label}{counts}{detail}", file=sys.stderr)

        on_progress = print_progress

    artifacts = perform_run(
        data_dir,
        files,
        passwords,
        profile,
        mode=QCRunMode(mode_value),
        allow_large_workbooks=ns.allow_large_workbooks,
        acceptance_absolute=ns.accept_absolute,
        acceptance_relative=ns.accept_percent / 100.0,
        compare_sheets=_parse_sheet_list(ns.sheets),
        compare_slides=_parse_slide_list(ns.slides),
        package_manifest=PackageManifest.from_role_files(files),
        compare_member_sheets=dict(compare_member_sheets),
        on_progress=on_progress,
        write_reports=True,  # CLI is batch: the report files ARE the output
    )
    result = artifacts.result

    print(f"mode: {result.mode.value}   profile: {result.profile_name}")
    print("files:", "  ".join(f"{r}={n}" for r, n in result.files.items()))
    if capability_limited(result.coverage):
        print(
            "status: capability-limited - one or more required checks were "
            "unavailable, so a low finding count is not a clean result"
        )
    # Streaming summaries: identical counts to the list builders without
    # holding every finding of a monster run in memory.
    pattern_summaries = summarize_pattern_groups(result.findings)
    pattern_counts = counts_from_summaries(pattern_summaries)
    grouped_counts = counts_from_summaries(
        summarize_review_groups(result.findings)
    )
    print(
        "pattern review items:",
        "  ".join(
            f"{severity.value}={count}"
            for severity, count in pattern_counts.review_items.items()
        ),
    )
    print(
        "spatial review items:",
        "  ".join(
            f"{severity.value}={count}"
            for severity, count in grouped_counts.review_items.items()
        ),
    )
    print(
        "atomic findings:",
        "  ".join(
            f"{severity.value}={count}"
            for severity, count in pattern_counts.atomic_findings.items()
        ),
    )
    if result.coverage:
        states: dict[str, int] = {}
        for item in result.coverage:
            states[item.state.value] = states.get(item.state.value, 0) + 1
        print("coverage:", "  ".join(f"{k}={v}" for k, v in sorted(states.items())))
    if result.mapping_coverage is not None:
        mc = result.mapping_coverage
        print(
            f"mapping: eligible={mc.eligible} mapped={mc.mapped} "
            f"verified={mc.verified} mismatched={mc.mismatched} "
            f"unmapped={mc.unmapped}"
        )
    for disclosure in result.disclosures:
        print(f"note: {disclosure}")
    if ns.individual_findings:
        ranked = [
            finding
            for finding in result.findings
            if finding.severity in (Severity.CRITICAL, Severity.WARNING)
        ]
        lines = [
            (
                f"  {finding.finding_id} "
                f"{finding.severity.value if finding.severity else '?':8s} "
                f"{finding.finding_class.value:24s} "
                f"{finding.sheet or finding.slide or ''}!"
                f"{finding.location or finding.element or ''}  {finding.message}"
            )
            for finding in ranked[: ns.top]
        ]
    else:
        ranked = [
            summary
            for summary in pattern_summaries
            if summary.severity in (Severity.CRITICAL, Severity.WARNING)
        ]
        lines = [
            (
                f"  {summary.group_id} {summary.severity.value:8s} "
                f"{summary.finding_class.value:24s} "
                f"{summary.sheet or summary.slide or ''}!{summary.bounding_range}  "
                f"{summary.member_count:,} affected finding"
                f"{'s' if summary.member_count != 1 else ''}"
            )
            for summary in ranked[: ns.top]
        ]
    for line in lines:
        print(line)
    if len(ranked) > ns.top:
        print(f"  ... {len(ranked) - ns.top} more (see reports)")
    print("reports:", "  ".join(str(p) for p in artifacts.report_paths.values()))
    if ns.json is not None:
        write_json_report(
            result,
            ns.json,
            include_context=ns.json_context,
            include_review_summary=ns.json_review_summary,
        )
        print("json:", ns.json)
    print(f"history: run #{artifacts.run_id} in {data_dir}")
    if ns.attestation is not None:
        from qc_tool.attestation import (
            create_attestation,
            load_attestation_key,
            load_or_create_attestation_key,
        )

        if ns.attestation_key_file is not None:
            key = load_attestation_key(ns.attestation_key_file)
        else:
            key_path, key = load_or_create_attestation_key(data_dir)
            print(f"attestation key: {key_path}")
        create_attestation(
            ns.attestation,
            result=result,
            profile=profile,
            input_files=files,
            report_paths=artifacts.report_paths,
            key=key,
        )
        print(f"attestation: {ns.attestation}")

    counts = result.counts
    failing = counts[Severity.CRITICAL]
    if ns.fail_on == "warning":
        failing += counts[Severity.WARNING]
    if ns.fail_on != "never" and failing > 0:
        return 2
    return 0


# --- sanitize -------------------------------------------------------------------


def _sanitize_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=f"{_program_name()} sanitize",
        description=(
            "Write a structure-preserving copy. Default mode is a local numeric "
            "scrambler and is NOT safe to share. --redact-text enables strict "
            "identifier/content removal and fail-closed privacy verification."
        ),
    )
    parser.add_argument("source", type=Path)
    parser.add_argument("-o", "--output", type=Path, default=None)
    parser.add_argument("--seed", type=int, default=0)
    parser.add_argument(
        "--redact-text",
        action="store_true",
        help="also replace text labels (period labels are always kept)",
    )
    return parser


def _cmd_sanitize(args: list[str]) -> int:
    ns = _sanitize_parser().parse_args(args)
    from qc_tool.sanitize import SanitizeError, sanitize_file

    try:
        output, stats = sanitize_file(
            ns.source, ns.output, seed=ns.seed, redact_text=ns.redact_text
        )
    except SanitizeError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    print(f"sanitized -> {output}")
    print(
        f"numbers scrambled: {stats.numbers_scrambled}   "
        f"texts redacted: {stats.texts_redacted}   "
        f"charts rebuilt: {stats.charts_rebuilt}"
        + (f"   charts skipped: {stats.charts_skipped}" if stats.charts_skipped else "")
    )
    for note in stats.notes:
        print(f"note: {note}")
    if not ns.redact_text:
        print(
            "warning: numeric scramble only — NOT privacy verified or safe to share; "
            "use --redact-text for strict verified redaction",
            file=sys.stderr,
        )
    return 0


def _sanitize_package_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=f"{_program_name()} sanitize-package",
        description=(
            "Strictly sanitize current Excel member(s)+PowerPoint together, "
            "re-project confirmed mappings, verify privacy, and write a "
            "redaction manifest."
        ),
    )
    parser.add_argument(
        "--excel",
        type=Path,
        default=None,
        help="primary current workbook",
    )
    parser.add_argument(
        "--workbook",
        action="append",
        default=[],
        metavar="MEMBER=FILE",
        help="add a current workbook member (repeatable; max eight)",
    )
    parser.add_argument("--ppt", type=Path, required=True)
    parser.add_argument("--profile", required=True, metavar="NAME_OR_PATH")
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--data-dir", type=Path, default=None)
    parser.add_argument("--seed", type=int, default=0)
    parser.add_argument(
        "--forbid",
        action="append",
        default=[],
        metavar="TOKEN",
        help="case-insensitive token that must not remain (repeatable)",
    )
    return parser


def _cmd_sanitize_package(args: list[str]) -> int:
    ns = _sanitize_package_parser().parse_args(args)
    from qc_tool.package_sanitize import sanitize_package_files
    from qc_tool.sanitize import SanitizeError

    data_dir = ns.data_dir or default_data_dir()
    profile = _resolve_profile(ns.profile, data_dir)
    files: dict[str, Path] = {"current_ppt": ns.ppt}
    if ns.excel is not None:
        files["current_excel"] = ns.excel
    for pair in ns.workbook:
        raw_member, separator, raw_path = pair.partition("=")
        if not separator:
            raise ValueError(f"expected MEMBER=PATH, got {pair!r}")
        member_id = raw_member.strip()
        if member_id == "primary":
            raise ValueError("member id 'primary' uses --excel")
        if re.fullmatch(MEMBER_ID_PATTERN, member_id) is None:
            raise ValueError(f"invalid member id {member_id!r}")
        role_key = f"current_excel:{member_id}"
        if role_key in files:
            raise ValueError(f"duplicate member role {role_key}")
        files[role_key] = Path(raw_path)
    if not any(role.startswith("current_excel") for role in files):
        raise ValueError("sanitize-package needs --excel and/or --workbook")
    for role, path in files.items():
        if not path.is_file():
            raise ValueError(f"{role}: {path} not found")
    package_manifest = PackageManifest.from_role_files(files)
    try:
        manifest = sanitize_package_files(
            files,
            package_manifest,
            ns.output_dir,
            profile,
            seed=ns.seed,
            forbidden_tokens=ns.forbid,
        )
    except SanitizeError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    print(
        f"package privacy: {'SAFE' if manifest.privacy_safe else 'UNSAFE'}  "
        f"mappings verified={manifest.mappings_verified} "
        f"unverifiable={manifest.mappings_unverifiable}"
    )
    print(f"outputs: {ns.output_dir}")
    print(f"manifest: {ns.output_dir / 'redaction-manifest.json'}")
    return 0 if manifest.privacy_safe and not manifest.mappings_unverifiable else 2


def _verify_sanitized_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=f"{_program_name()} verify-sanitized",
        description="Fail-closed privacy verification for a strict sanitizer output.",
    )
    parser.add_argument("source", type=Path)
    parser.add_argument(
        "--forbid",
        action="append",
        default=[],
        metavar="TOKEN",
        help="case-insensitive token that must not remain (repeatable)",
    )
    parser.add_argument("--json", type=Path, default=None, help="write verification JSON")
    return parser


def _cmd_verify_sanitized(args: list[str]) -> int:
    ns = _verify_sanitized_parser().parse_args(args)
    from qc_tool.privacy import verify_sanitized
    from qc_tool.security import private_directory, private_file

    report = verify_sanitized(ns.source, forbidden_tokens=ns.forbid)
    for issue in report.issues:
        print(f"unsafe  {issue.code:24s} {issue.location}: {issue.message}")
    print(
        f"privacy: {'SAFE' if report.safe else 'UNSAFE'}  "
        f"checks={len(report.checks)} issues={len(report.issues)} sha256={report.sha256}"
    )
    if ns.json is not None:

        private_directory(ns.json.parent)
        ns.json.write_text(report.model_dump_json(indent=2), encoding="utf-8")
        private_file(ns.json)
    return 0 if report.safe else 2


def _fingerprint_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=f"{_program_name()} fingerprint",
        description=(
            "Write structural-only JSON: dimensions, types, period label patterns, "
            "regions, hashed formula patterns, table/chart shapes — no source "
            "values, visible text, paths, formulas, images, or identifiers."
        ),
    )
    parser.add_argument("source", type=Path, nargs="?")
    for role in _ROLES:
        parser.add_argument(
            f"--{role.replace('_', '-')}",
            type=Path,
            default=None,
            metavar="FILE",
        )
    parser.add_argument(
        "--baseline-workbook",
        action="append",
        default=[],
        metavar="MEMBER=FILE",
    )
    parser.add_argument(
        "--current-workbook",
        action="append",
        default=[],
        metavar="MEMBER=FILE",
    )
    parser.add_argument("-o", "--output", type=Path, required=True)
    parser.add_argument("--password-env", default=None, metavar="ENV_VAR")
    parser.add_argument("--password-file", type=Path, default=None)
    parser.add_argument("--password-prompt", action="store_true")
    parser.add_argument(
        "--package-password-env",
        action="append",
        default=[],
        metavar="ROLE_OR_FILENAME=ENV_VAR",
    )
    parser.add_argument(
        "--package-password-file",
        action="append",
        default=[],
        metavar="ROLE_OR_FILENAME=PATH",
    )
    parser.add_argument(
        "--package-password-prompt",
        action="append",
        default=[],
        metavar="ROLE_OR_FILENAME",
    )
    return parser


def _cmd_fingerprint(args: list[str]) -> int:
    ns = _fingerprint_parser().parse_args(args)
    from qc_tool.fingerprint import write_fingerprint, write_package_fingerprint

    package_files = {
        role: getattr(ns, role)
        for role in _ROLES
        if getattr(ns, role) is not None
    }
    for pairs, prefix in (
        (ns.baseline_workbook, "baseline_excel"),
        (ns.current_workbook, "current_excel"),
    ):
        for pair in pairs:
            raw_member, separator, raw_path = pair.partition("=")
            if not separator:
                raise ValueError(f"expected MEMBER=PATH, got {pair!r}")
            member_id = raw_member.strip()
            if member_id == "primary":
                raise ValueError(
                    f"member id 'primary' uses --{prefix.replace('_', '-')}"
                )
            if re.fullmatch(MEMBER_ID_PATTERN, member_id) is None:
                raise ValueError(f"invalid member id {member_id!r}")
            role_key = f"{prefix}:{member_id}"
            if role_key in package_files:
                raise ValueError(f"duplicate member role {role_key}")
            package_files[role_key] = Path(raw_path)
    if ns.source is not None and package_files:
        raise ValueError("a single source cannot be combined with package members")
    if ns.source is None and not package_files:
        raise ValueError("fingerprint needs a source or package member arguments")
    for role, path in package_files.items():
        if not path.is_file():
            raise ValueError(f"{role}: {path} not found")

    password = None
    sources = sum(
        value is not None and value is not False
        for value in (ns.password_env, ns.password_file, ns.password_prompt)
    )
    if sources > 1:
        raise ValueError("choose only one fingerprint password source")
    if package_files and sources:
        raise ValueError(
            "single-source password flags cannot be combined with package members"
        )
    if ns.password_env:
        if ns.password_env not in os.environ:
            raise ValueError(f"environment variable {ns.password_env!r} is not set")
        password = os.environ[ns.password_env]
    elif ns.password_file:
        if os.name == "posix" and ns.password_file.stat().st_mode & 0o077:
            raise ValueError(
                f"password file {ns.password_file} must not be group/world accessible"
            )
        password = ns.password_file.read_text(encoding="utf-8").rstrip("\r\n")
    elif ns.password_prompt:
        source_label = ns.source.name if ns.source is not None else "source"
        password = getpass.getpass(f"Password for {source_label}: ")
    if package_files:
        password_ns = argparse.Namespace(
            password=[],
            password_env=ns.package_password_env,
            password_file=ns.package_password_file,
            password_prompt=ns.package_password_prompt,
        )
        payload = write_package_fingerprint(
            package_files,
            PackageManifest.from_role_files(package_files),
            ns.output,
            passwords=_parse_password_sources(password_ns, package_files),
        )
    else:
        if (
            ns.package_password_env
            or ns.package_password_file
            or ns.package_password_prompt
        ):
            raise ValueError("package password flags need package members")
        if ns.source is None:  # guarded above; narrows Path for Pyright
            raise ValueError("fingerprint source is missing")
        payload = write_fingerprint(ns.source, ns.output, password=password)
    print(f"fingerprint: {payload['fingerprint_id']} -> {ns.output}")
    return 0


def _verify_attestation_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=f"{_program_name()} verify-attestation",
        description="Verify an attestation HMAC and every bundled member hash.",
    )
    parser.add_argument("source", type=Path)
    parser.add_argument("--data-dir", type=Path, default=None)
    parser.add_argument("--key-file", type=Path, default=None)
    return parser


def _cmd_verify_attestation(args: list[str]) -> int:
    ns = _verify_attestation_parser().parse_args(args)
    from qc_tool.attestation import (
        load_attestation_key,
        load_or_create_attestation_key,
        verify_attestation,
    )

    if ns.key_file is not None:
        key = load_attestation_key(ns.key_file)
    else:
        _, key = load_or_create_attestation_key(ns.data_dir or default_data_dir())
    result = verify_attestation(ns.source, key=key)
    for issue in result.issues:
        print(f"invalid {issue.code:20s} {issue.message}")
    print(
        f"attestation: {'VALID' if result.valid else 'INVALID'}  "
        f"key_id={result.key_id} issues={len(result.issues)}"
    )
    return 0 if result.valid else 2


def _network_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=f"{_program_name()} network",
        description=(
            "Persist fail-safe server exposure: local loopback or temporary "
            "unauthenticated LAN/WAN-capable binding."
        ),
    )
    parser.add_argument("action", choices=["status", "local", "lan"])
    parser.add_argument("--minutes", type=int, default=None)
    parser.add_argument("--data-dir", type=Path, default=None)
    return parser


def _cmd_network(args: list[str]) -> int:
    ns = _network_parser().parse_args(args)
    from qc_tool.server_config import (
        load_server_config,
        local_config,
        save_server_config,
        temporary_lan_config,
    )

    data_dir = ns.data_dir or default_data_dir()
    stored_config = load_server_config(data_dir)
    if ns.action == "local":
        if ns.minutes is not None:
            raise ValueError("--minutes only applies to network lan")
        config = local_config(desktop_focus=stored_config.desktop_focus)
        save_server_config(data_dir, config)
    elif ns.action == "lan":
        if ns.minutes is None:
            raise ValueError("network lan requires --minutes (1-1440)")
        config = temporary_lan_config(
            ns.minutes,
            desktop_focus=stored_config.desktop_focus,
        )
        save_server_config(data_dir, config)
    else:
        if ns.minutes is not None:
            raise ValueError("--minutes does not apply to network status")
        config = stored_config
    print(f"network: {config.network.value}  host={config.host}")
    if config.expires_at is not None:
        print(f"expires: {config.expires_at.isoformat(timespec='seconds')}")
    if config.network.value == "lan":
        print("warning: no built-in authentication or TLS; router forwarding is separate")
    return 0


# --- lint --------------------------------------------------------------------------


def _lint_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=f"{_program_name()} lint",
        description="Validate a deliverable profile, optionally against real files.",
    )
    parser.add_argument("profile", type=Path, help="profile YAML path")
    parser.add_argument("--against-excel", type=Path, default=None, metavar="XLSX")
    parser.add_argument("--against-ppt", type=Path, default=None, metavar="PPTX")
    return parser


def _cmd_lint(args: list[str]) -> int:
    ns = _lint_parser().parse_args(args)
    from qc_tool.config.lint import lint_profile
    from qc_tool.config.profile import load_profile

    try:
        profile = load_profile(ns.profile)
    except Exception as exc:
        print(f"error: profile failed to load: {exc}", file=sys.stderr)
        return 2
    workbook = deck = None
    if ns.against_excel is not None:
        from qc_tool.io.loader import load_workbook_snapshot

        workbook = load_workbook_snapshot(ns.against_excel)
    if ns.against_ppt is not None:
        from qc_tool.ppt.extract import load_deck_snapshot

        deck = load_deck_snapshot(ns.against_ppt)

    issues = lint_profile(profile, workbook=workbook, deck=deck)
    for issue in issues:
        print(f"{issue.level:7s} {issue.where}: {issue.message}")
    errors = sum(1 for issue in issues if issue.level == "error")
    warnings = len(issues) - errors
    print(f"lint: {errors} error(s), {warnings} warning(s)")
    return 2 if errors else 0


# --- dispatch -----------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:]) if argv is None else list(argv)
    if args and args[0] in _SUBCOMMANDS:
        command, rest = args[0], args[1:]
    else:
        command, rest = "serve", args  # backward compatible bare invocation
    handlers = {
        "serve": _cmd_serve,
        "launch": _cmd_launch,
        "shortcut": _cmd_shortcut,
        "run": _cmd_run,
        "sanitize": _cmd_sanitize,
        "sanitize-package": _cmd_sanitize_package,
        "verify-sanitized": _cmd_verify_sanitized,
        "fingerprint": _cmd_fingerprint,
        "verify-attestation": _cmd_verify_attestation,
        "network": _cmd_network,
        "lint": _cmd_lint,
    }
    try:
        return handlers[command](rest)
    except (OSError, ValueError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
