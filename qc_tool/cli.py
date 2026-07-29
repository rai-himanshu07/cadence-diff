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
import os
import sys
from pathlib import Path

from platformdirs import user_data_dir

from qc_tool import __version__
from qc_tool.security import private_directory

_SUBCOMMANDS = {
    "serve",
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
    return parser


def _cmd_serve(args: list[str]) -> int:
    ns = build_parser().parse_args(args)
    data_dir = ns.data_dir or default_data_dir()
    data_dir.mkdir(parents=True, exist_ok=True)
    from qc_tool.server_config import (
        load_server_config,
        local_config,
        save_server_config,
        temporary_lan_config,
    )

    if ns.network == "lan":
        if ns.expose_for is None:
            raise ValueError("--network lan requires --expose-for MINUTES")
        config = temporary_lan_config(ns.expose_for)
        save_server_config(data_dir, config)
    elif ns.network == "local":
        if ns.expose_for is not None:
            raise ValueError("--expose-for only applies to --network lan")
        config = local_config()
        save_server_config(data_dir, config)
    else:
        if ns.expose_for is not None:
            raise ValueError("--expose-for requires --network lan")
        config = load_server_config(data_dir)
    from qc_tool.ui.app import run_app  # deferred: keep --help/--version instant

    run_app(
        data_dir,
        port=ns.port,
        host=config.host,
        network_mode=config.network,
        expires_at=config.expires_at,
    )
    return 0


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
    if "baseline_excel" in files or "baseline_ppt" in files:
        return "cycle_comparison"
    if "current_excel" in files and "current_ppt" in files:
        return "final_package"
    return "current_file_preflight"


def _resolve_profile(spec: str | None, data_dir: Path):
    from qc_tool.config.profile import default_profile, load_profile
    from qc_tool.ui.app import load_profile_by_name

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
    if key in _ROLES:
        if key not in files:
            raise ValueError(f"password role {key!r} has no supplied file")
        return [key]
    matches = [role for role, path in files.items() if path.name == key]
    if not matches:
        raise ValueError(f"password key {key!r} matches no supplied file or role")
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


def _cmd_run(args: list[str]) -> int:
    ns = _run_parser().parse_args(args)
    files = {
        role: getattr(ns, role)
        for role in _ROLES
        if getattr(ns, role) is not None
    }
    if not files:
        raise ValueError(
            f"{_program_name()} run: supply at least one file (see --help)"
        )
    for role, path in files.items():
        if not path.exists():
            raise ValueError(f"{role}: {path} not found")

    from qc_tool.coverage import QCRunMode
    from qc_tool.findings import Severity
    from qc_tool.report.json_report import write_json_report
    from qc_tool.ui.app import perform_run

    data_dir = ns.data_dir or default_data_dir()
    private_directory(data_dir)
    mode_value = _MODE_ALIASES.get(ns.mode, ns.mode) if ns.mode else _infer_mode(files)
    profile = _resolve_profile(ns.profile, data_dir)
    passwords = _parse_password_sources(ns, files)

    artifacts = perform_run(
        data_dir, files, passwords, profile, mode=QCRunMode(mode_value)
    )
    result = artifacts.result

    print(f"mode: {result.mode.value}   profile: {result.profile_name}")
    print("files:", "  ".join(f"{r}={n}" for r, n in result.files.items()))
    print(
        "counts:",
        "  ".join(f"{sev.value}={count}" for sev, count in result.counts.items()),
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
    ranked = [
        f
        for f in result.findings
        if f.severity in (Severity.CRITICAL, Severity.WARNING)
    ]
    for finding in ranked[: ns.top]:
        severity = finding.severity.value if finding.severity else "?"
        where = finding.sheet or finding.slide or ""
        print(
            f"  {finding.finding_id} {severity:8s} {finding.finding_class.value:24s} "
            f"{where}!{finding.location or finding.element or ''}  {finding.message}"
        )
    if len(ranked) > ns.top:
        print(f"  ... {len(ranked) - ns.top} more (see reports)")
    print("reports:", "  ".join(str(p) for p in artifacts.report_paths.values()))
    if ns.json is not None:
        write_json_report(result, ns.json, include_context=ns.json_context)
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
            "Strictly sanitize current Excel+PowerPoint together, re-project "
            "confirmed mappings, verify privacy, and write a redaction manifest."
        ),
    )
    parser.add_argument("--excel", type=Path, required=True)
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
    from qc_tool.package_sanitize import sanitize_package
    from qc_tool.sanitize import SanitizeError

    data_dir = ns.data_dir or default_data_dir()
    profile = _resolve_profile(ns.profile, data_dir)
    try:
        manifest = sanitize_package(
            ns.excel,
            ns.ppt,
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
            "Write structural-only JSON: dimensions, types, period grammars, "
            "regions, hashed formula patterns, table/chart shapes — no source "
            "values, visible text, paths, formulas, images, or identifiers."
        ),
    )
    parser.add_argument("source", type=Path)
    parser.add_argument("-o", "--output", type=Path, required=True)
    parser.add_argument("--password-env", default=None, metavar="ENV_VAR")
    parser.add_argument("--password-file", type=Path, default=None)
    parser.add_argument("--password-prompt", action="store_true")
    return parser


def _cmd_fingerprint(args: list[str]) -> int:
    ns = _fingerprint_parser().parse_args(args)
    from qc_tool.fingerprint import write_fingerprint

    password = None
    sources = sum(
        value is not None and value is not False
        for value in (ns.password_env, ns.password_file, ns.password_prompt)
    )
    if sources > 1:
        raise ValueError("choose only one fingerprint password source")
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
        password = getpass.getpass(f"Password for {ns.source.name}: ")
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
    if ns.action == "local":
        if ns.minutes is not None:
            raise ValueError("--minutes only applies to network lan")
        config = local_config()
        save_server_config(data_dir, config)
    elif ns.action == "lan":
        if ns.minutes is None:
            raise ValueError("network lan requires --minutes (1-1440)")
        config = temporary_lan_config(ns.minutes)
        save_server_config(data_dir, config)
    else:
        if ns.minutes is not None:
            raise ValueError("--minutes does not apply to network status")
        config = load_server_config(data_dir)
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
