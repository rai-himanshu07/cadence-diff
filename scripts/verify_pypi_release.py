"""Verify exact release filenames and SHA-256 digests visible on PyPI."""

from __future__ import annotations

import argparse
import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

_DIGEST_RE = re.compile(r"^[0-9a-f]{64}$")
_FILENAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._+-]{0,199}$")


def load_expected(path: Path) -> dict[str, str]:
    expected: dict[str, str] = {}
    for raw in path.read_text(encoding="utf-8").splitlines():
        if not raw.strip():
            continue
        digest, separator, filename = raw.partition("  ")
        if (
            not separator
            or not _DIGEST_RE.fullmatch(digest)
            or not _FILENAME_RE.fullmatch(filename)
            or filename in expected
        ):
            raise ValueError("checksum manifest has an invalid row")
        expected[filename] = digest
    if not expected:
        raise ValueError("checksum manifest is empty")
    return expected


def verify_payload(payload: object, expected: dict[str, str]) -> tuple[bool, str]:
    if not isinstance(payload, dict) or not isinstance(payload.get("urls"), list):
        return False, "PyPI response has an invalid shape"
    observed: dict[str, str] = {}
    for item in payload["urls"]:
        if not isinstance(item, dict):
            return False, "PyPI response has an invalid file row"
        filename = item.get("filename")
        digests = item.get("digests")
        digest = digests.get("sha256") if isinstance(digests, dict) else None
        if isinstance(filename, str) and isinstance(digest, str):
            observed[filename] = digest
    missing = sorted(set(expected) - set(observed))
    unexpected = sorted(set(observed) - set(expected))
    mismatched = sorted(
        filename
        for filename in set(expected) & set(observed)
        if expected[filename] != observed[filename]
    )
    if missing:
        return False, f"missing expected files: {', '.join(missing)}"
    if unexpected:
        return False, f"unexpected release files: {', '.join(unexpected)}"
    if mismatched:
        return False, f"SHA-256 mismatch: {', '.join(mismatched)}"
    return True, "exact release files and hashes are visible"


def fetch_payload(project: str, version: str) -> object:
    project_part = urllib.parse.quote(project, safe="")
    version_part = urllib.parse.quote(version, safe="")
    request = urllib.request.Request(
        f"https://pypi.org/pypi/{project_part}/{version_part}/json",
        headers={"Accept": "application/json", "User-Agent": "cadence-diff-release"},
    )
    with urllib.request.urlopen(request, timeout=15) as response:
        return json.load(response)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("project")
    parser.add_argument("version")
    parser.add_argument("checksum_manifest", type=Path)
    parser.add_argument("--attempts", type=int, default=12)
    parser.add_argument("--delay-seconds", type=float, default=10.0)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    if args.attempts < 1 or args.delay_seconds < 0:
        print("invalid retry policy")
        return 2
    try:
        expected = load_expected(args.checksum_manifest)
    except (OSError, ValueError) as exc:
        print(f"verification input invalid ({type(exc).__name__})")
        return 2
    detail = "release not visible"
    for attempt in range(args.attempts):
        try:
            payload = fetch_payload(args.project, args.version)
            passed, detail = verify_payload(payload, expected)
        except (OSError, ValueError, urllib.error.URLError, json.JSONDecodeError) as exc:
            passed = False
            detail = f"PyPI query failed ({type(exc).__name__})"
        if passed:
            print(f"PyPI verification passed: {detail}")
            return 0
        if attempt + 1 < args.attempts:
            time.sleep(args.delay_seconds)
    print(f"PyPI verification failed: {detail}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
