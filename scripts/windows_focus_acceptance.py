"""Aggregate-only Windows acceptance probe for click-to-focus (plan Step 9).

Drives the production focus path end to end - discover, resolve, bind,
revalidate, dispatch, post-check - with no UI and no new code path. It never
opens, saves, recalculates, closes, or quits an analyst document; the only
Office effect is the selection change the feature exists to make.

Output is fixed codes, counts, booleans, and timings. Paths, sheet names,
addresses, hashes, window titles, and user SIDs never reach the result file.

Run inside the guest, one scenario at a time:

    cd /d C:\\QC-Pilot\\QC_Tool
    copy /Y Z:\\QC_Tool\\scripts\\windows_focus_acceptance.py scripts\\
    conda run --no-capture-output -n cadence-diff-dev python -m ^
      scripts.windows_focus_acceptance --scenario excel-current ^
      --source "C:\\pilot\\current.xlsx" --sheet "Summary" --address "B5" ^
      --expect focused --output "Z:\\QC_Tool\\windows-return\\focus-excel-current.json"
"""

from __future__ import annotations

import argparse
import asyncio
import contextlib
import hashlib
import json
import os
import secrets
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from qc_tool.focus.binding import (
    BindingRegistry,
    BindingRequest,
    BindOutcome,
    resolve_binding,
    revalidate_binding,
)
from qc_tool.focus.discovery import (
    FocusApplication,
    discover_open_documents,
)
from qc_tool.focus.model import FocusRole
from qc_tool.focus.navigator import FocusNavigator
from qc_tool.focus.package_risk import scan_focus_package
from qc_tool.focus.protocol import SCHEMA_VERSION, FocusAction

_SCHEMA_VERSION = 1

SCENARIOS = (
    "excel-current",
    "excel-baseline",
    "ppt-current",
    "ppt-baseline",
    "duplicate-hash",
    "managed-copy",
    "dirty-matching",
    "dirty-changed",
    "autosave-on",
    "two-excel-processes",
    "multi-window",
    "protected-view",
    "active-content",
    "helper-timeout",
    "foreground-denial",
    "latency",
)

_ROLES = {
    "current_excel": FocusRole.CURRENT_EXCEL,
    "baseline_excel": FocusRole.BASELINE_EXCEL,
    "current_ppt": FocusRole.CURRENT_PPT,
    "baseline_ppt": FocusRole.BASELINE_PPT,
}
_APPLICATIONS = {
    FocusRole.CURRENT_EXCEL: FocusApplication.EXCEL,
    FocusRole.BASELINE_EXCEL: FocusApplication.EXCEL,
    FocusRole.CURRENT_PPT: FocusApplication.POWERPOINT,
    FocusRole.BASELINE_PPT: FocusApplication.POWERPOINT,
}

_VERDICT_PASS = "expected_outcome"
_VERDICT_FAIL = "unexpected_outcome"
_VERDICT_MUTATED = "source_or_process_mutated"
_FATAL_VERDICTS = frozenset({_VERDICT_FAIL, _VERDICT_MUTATED})


@dataclass(slots=True)
class Observation:
    """Everything the probe is allowed to report about one scenario."""

    scenario: str
    role: str
    expected: str
    bind_outcome: str = ""
    focus_outcome: str = ""
    focus_stage: str = ""
    verdict: str = ""
    documents_discovered: int = 0
    analyst_candidates: int = 0
    addin_documents: int = 0
    hidden_instance_documents: int = 0
    windowless_documents: int = 0
    duplicate_path_analyst_documents: int = 0
    processes_before: int = 0
    processes_after: int = 0
    discovery_reasons: list[str] = field(default_factory=list)
    package_risks: list[str] = field(default_factory=list)
    package_error: str = ""
    source_unchanged: bool = True
    dirty_at_bind: bool = False
    autosave_at_bind: str = "unknown"
    unsaved_changes_reported: bool = False
    dispatched: bool = False
    bind_seconds: float = 0.0
    focus_seconds: float = 0.0
    focus_seconds_samples: list[float] = field(default_factory=list)
    navigator_disabled: bool = False

    def payload(self) -> dict[str, object]:
        data = {
            key: getattr(self, key)
            for key in self.__slots__  # type: ignore[attr-defined]
        }
        return dict(sorted(data.items()))


def _sha256(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1 << 20), b""):
                digest.update(chunk)
        return digest.hexdigest()
    except PermissionError:
        if sys.platform != "win32":
            raise
        return _sha256_windows_shared(path)


def _sha256_windows_shared(path: Path) -> str:
    import win32con  # pyright: ignore[reportMissingModuleSource]
    import win32file  # pyright: ignore[reportMissingModuleSource]

    handle = win32file.CreateFile(
        str(path),
        win32con.GENERIC_READ,
        win32con.FILE_SHARE_READ
        | win32con.FILE_SHARE_WRITE
        | win32con.FILE_SHARE_DELETE,
        None,
        win32con.OPEN_EXISTING,
        win32con.FILE_ATTRIBUTE_NORMAL,
        None,
    )
    digest = hashlib.sha256()
    read_file: Any = win32file.ReadFile
    try:
        while True:
            _status, chunk = read_file(handle, 1 << 20)
            if not chunk:
                break
            digest.update(chunk)
    finally:
        handle.Close()
    return digest.hexdigest()


def _set_verdict(observation: Observation, actual: str) -> None:
    if not observation.source_unchanged or (
        observation.processes_after != observation.processes_before
    ):
        observation.verdict = _VERDICT_MUTATED
    elif actual == observation.expected:
        observation.verdict = _VERDICT_PASS
    else:
        observation.verdict = _VERDICT_FAIL


def _postcheck(
    observation: Observation,
    application: FocusApplication,
    source: Path,
    before_hash: str,
    actual: str,
) -> None:
    after = discover_open_documents(application)
    observation.processes_after = len({d.process_id for d in after.documents})
    observation.source_unchanged = _sha256(source) == before_hash
    _set_verdict(observation, actual)


def _counts(documents) -> dict[str, int]:
    seen: dict[str, int] = {}
    duplicates = 0
    for document in documents:
        if not document.analyst_candidate:
            continue
        key = document.full_name.casefold()
        seen[key] = seen.get(key, 0) + 1
    for count in seen.values():
        if count > 1:
            duplicates += count
    return {
        "documents_discovered": len(documents),
        "analyst_candidates": sum(1 for d in documents if d.analyst_candidate),
        "addin_documents": sum(1 for d in documents if d.is_addin),
        "hidden_instance_documents": sum(1 for d in documents if d.hidden_instance),
        "windowless_documents": sum(1 for d in documents if d.window_count == 0),
        "duplicate_path_analyst_documents": duplicates,
    }


async def _run(args: argparse.Namespace) -> Observation:
    role = _ROLES[args.role]
    application = _APPLICATIONS[role]
    source = Path(args.source).resolve()
    observation = Observation(
        scenario=args.scenario, role=role.value, expected=args.expect
    )

    scan = scan_focus_package(source)
    observation.package_risks = [risk.value for risk in scan.risks]
    observation.package_error = scan.error.value if scan.error else ""

    before_hash = _sha256(source)
    expected_hash = args.sha256 or before_hash

    started = time.monotonic()
    discovery = discover_open_documents(application)
    observation.bind_seconds = round(time.monotonic() - started, 3)
    observation.discovery_reasons = [r.value for r in discovery.reasons]
    for key, value in _counts(discovery.documents).items():
        setattr(observation, key, value)
    observation.processes_before = len({d.process_id for d in discovery.documents})

    managed_root = Path(args.managed_root).resolve() if args.managed_root else (
        Path(os.getcwd()) / "___no_managed_root___"
    )
    request = BindingRequest(
        run_id=1,
        role=role,
        expected_sha256=expected_hash,
        managed_root=managed_root,
        managed_path=Path(args.managed_path).resolve() if args.managed_path else None,
    )
    outcome, document, identity = resolve_binding(request, discovery)
    observation.bind_outcome = outcome.value
    if document is not None:
        observation.dirty_at_bind = document.saved is False
        observation.autosave_at_bind = (
            "unknown" if document.autosave is None else str(document.autosave).lower()
        )

    if outcome is not BindOutcome.MATCHED or document is None or identity is None:
        _postcheck(
            observation,
            application,
            source,
            before_hash,
            outcome.value,
        )
        return observation

    registry = BindingRegistry()
    binding = registry.bind("probe-client", request, document, identity)
    observation.unsaved_changes_reported = binding.unsaved_changes

    navigator = FocusNavigator(timeout=args.timeout)
    samples: list[float] = []
    reply = None
    for _attempt in range(max(1, args.repeats)):
        fresh = discover_open_documents(application)
        revalidated = revalidate_binding(binding, fresh, registry)
        if not revalidated.bound:
            observation.focus_outcome = revalidated.outcome.value
            break
        salt = secrets.token_bytes(16)
        digest = registry.path_digest("probe-client", 1, role, salt)
        reply = await navigator.submit(
            {
                "schema_version": SCHEMA_VERSION,
                "action": FocusAction.FOCUS.value,
                "application": application.value,
                "expected_sha256": binding.expected_sha256,
                "path_salt": salt.hex(),
                "expected_path_digest": digest,
                "process_id": binding.process_id,
                "process_created": binding.process_created,
                "windows_session_id": binding.windows_session_id,
                "window_handle": binding.window_handle,
                "file_id": list(binding.file_id),
                "sheet": args.sheet,
                "address": args.address,
                "slide_index": args.slide,
                "shape_id": args.shape_id,
            }
        )
        samples.append(round(reply.duration_seconds, 3))
        observation.focus_outcome = reply.code
        observation.focus_stage = reply.stage.value if reply.stage else ""
        observation.dispatched = reply.dispatched
        observation.unsaved_changes_reported |= reply.unsaved_changes

    observation.focus_seconds_samples = samples
    observation.focus_seconds = max(samples) if samples else 0.0
    observation.navigator_disabled = navigator.disabled

    _postcheck(
        observation,
        application,
        source,
        before_hash,
        observation.focus_outcome,
    )
    return observation


def _write(path: Path, payload: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(payload, indent=2, sort_keys=True), encoding="utf-8")
    os.replace(temporary, path)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="windows_focus_acceptance")
    parser.add_argument("--scenario", required=True, choices=SCENARIOS)
    parser.add_argument("--role", required=True, choices=sorted(_ROLES))
    parser.add_argument("--source", required=True, help="file staged open in Office")
    parser.add_argument("--sha256", default=None, help="override the expected run hash")
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--address", default=None)
    parser.add_argument("--slide", type=int, default=None)
    parser.add_argument("--shape-id", type=int, default=None)
    parser.add_argument("--managed-root", default=None)
    parser.add_argument("--managed-path", default=None)
    parser.add_argument("--expect", required=True, help="expected fixed outcome code")
    parser.add_argument("--timeout", type=float, default=25.0)
    parser.add_argument(
        "--repeats", type=int, default=1, help="dispatch N times to measure latency"
    )
    parser.add_argument("--output", type=Path, required=True)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if sys.platform != "win32":
        _write(
            args.output,
            {
                "schema_version": _SCHEMA_VERSION,
                "scenario": args.scenario,
                "verdict": "unsupported_platform",
            },
        )
        return 2
    owns_apartment = False
    with contextlib.suppress(Exception):
        import pythoncom  # pyright: ignore[reportMissingModuleSource]

        pythoncom.CoInitialize()
        owns_apartment = True
    try:
        observation = asyncio.run(_run(args))
    finally:
        if owns_apartment:
            with contextlib.suppress(Exception):
                import pythoncom  # pyright: ignore[reportMissingModuleSource]

                pythoncom.CoUninitialize()
    payload = {"schema_version": _SCHEMA_VERSION, **observation.payload()}
    _write(args.output, payload)
    print(json.dumps(payload, indent=2, sort_keys=True))
    return 1 if observation.verdict in _FATAL_VERDICTS else 0


if __name__ == "__main__":
    raise SystemExit(main())
