"""Audit the installed Python environment with distinct operational outcomes."""

from __future__ import annotations

import json
import subprocess
import sys
from collections.abc import Sequence

AUDIT_COMMAND = (
    sys.executable,
    "-m",
    "pip_audit",
    ".",
    "--format=json",
    "--progress-spinner=off",
    "--desc=off",
)


def interpret_audit(returncode: int, stdout: str, stderr: str) -> int:
    """Return 0 clean, 1 tool/feed failure, or 2 known vulnerabilities."""
    try:
        payload = json.loads(stdout)
    except json.JSONDecodeError:
        detail = stderr.strip().splitlines()[0] if stderr.strip() else "no JSON output"
        print(f"Dependency audit unavailable: {detail}", file=sys.stderr)
        return 1
    dependencies = payload.get("dependencies") if isinstance(payload, dict) else None
    if not isinstance(dependencies, list):
        print("Dependency audit unavailable: malformed JSON output", file=sys.stderr)
        return 1
    vulnerable: set[tuple[str, str, str]] = set()
    for dependency in dependencies:
        if not isinstance(dependency, dict):
            continue
        name = str(dependency.get("name", "unknown"))
        version = str(dependency.get("version", "unknown"))
        findings = dependency.get("vulns", [])
        if not isinstance(findings, list):
            continue
        for finding in findings:
            if isinstance(finding, dict):
                vulnerable.add((name, version, str(finding.get("id", "unknown"))))
    if vulnerable:
        print("Known dependency vulnerabilities detected:")
        for name, version, identifier in sorted(vulnerable):
            print(f"- {name} {version}: {identifier}")
        return 2
    if returncode != 0:
        detail = stderr.strip().splitlines()[0] if stderr.strip() else f"exit {returncode}"
        print(f"Dependency audit unavailable: {detail}", file=sys.stderr)
        return 1
    print(f"Dependency audit passed: {len(dependencies)} resolved packages checked")
    return 0


def main(command: Sequence[str] = AUDIT_COMMAND) -> int:
    result = subprocess.run(
        list(command),
        check=False,
        capture_output=True,
        text=True,
    )
    return interpret_audit(result.returncode, result.stdout, result.stderr)


if __name__ == "__main__":
    raise SystemExit(main())
