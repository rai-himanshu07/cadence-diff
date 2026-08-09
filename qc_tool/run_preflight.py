"""Shared byte-identity safeguards applied before any QC engine work."""

from __future__ import annotations

import hashlib
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from qc_tool.coverage import QCRunMode
from qc_tool.progress import CancellationToken, check_cancelled


@dataclass(frozen=True, slots=True)
class InputByteConflict:
    kind: Literal["same_side_duplicate", "identical_cycle_pair"]
    first_role: str
    second_role: str
    artifact: str
    side: str | None = None
    member_id: str = "primary"

    def message(self, label: Callable[[str], str] = str) -> str:
        if self.kind == "same_side_duplicate":
            return (
                f"Duplicate bytes for {self.artifact} on the {self.side} side: "
                f"{label(self.first_role)} and {label(self.second_role)}"
            )
        if self.artifact == "excel":
            return (
                "Baseline and current Excel for member "
                f"{self.member_id!r} are byte-identical; a comparison would "
                "prove nothing"
            )
        return (
            "Baseline and current PowerPoint are byte-identical; "
            "a comparison would prove nothing"
        )


class DuplicateInputBytesError(ValueError):
    """Supplied roles cannot prove a meaningful independent comparison."""


class InputBytesChangedError(ValueError):
    """A source changed after intake and no longer matches the accepted bytes."""


def _role_parts(role: str) -> tuple[str, str, str]:
    prefix, separator, member_id = role.partition(":")
    side, separator2, artifact = prefix.partition("_")
    if not separator2 or side not in {"baseline", "current"}:
        raise ValueError(f"invalid run file role {role!r}")
    return side, artifact, member_id if separator else "primary"


def hash_run_files(
    files: Mapping[str, Path],
    *,
    cancellation_token: CancellationToken | None = None,
) -> dict[str, str]:
    """Hash each input once in stable role order with cancellation checks."""
    hashes: dict[str, str] = {}
    for role in sorted(files):
        check_cancelled(cancellation_token)
        digest = hashlib.sha256()
        with files[role].open("rb") as handle:
            for chunk in iter(lambda: handle.read(1 << 20), b""):
                check_cancelled(cancellation_token)
                digest.update(chunk)
        hashes[role] = digest.hexdigest()
    return hashes


def duplicate_byte_conflicts(
    mode: QCRunMode,
    files: Mapping[str, object],
    file_hashes: Mapping[str, str],
) -> tuple[InputByteConflict, ...]:
    """Return every same-side duplicate and same-member identical cycle pair."""
    conflicts: list[InputByteConflict] = []
    grouped: dict[tuple[str, str], list[str]] = {}
    for role in files:
        side, artifact, _member_id = _role_parts(role)
        grouped.setdefault((side, artifact), []).append(role)
    for (side, artifact), roles in sorted(grouped.items()):
        seen: dict[str, str] = {}
        for role in sorted(roles):
            digest = file_hashes.get(role)
            if not digest:
                continue
            previous = seen.get(digest)
            if previous is not None:
                conflicts.append(
                    InputByteConflict(
                        kind="same_side_duplicate",
                        first_role=previous,
                        second_role=role,
                        artifact=artifact,
                        side=side,
                    )
                )
            else:
                seen[digest] = role

    if QCRunMode(mode) is QCRunMode.CYCLE_COMPARISON:
        for baseline_role in sorted(files):
            side, artifact, member_id = _role_parts(baseline_role)
            if side != "baseline":
                continue
            current_role = (
                f"current_{artifact}"
                if member_id == "primary"
                else f"current_{artifact}:{member_id}"
            )
            if current_role not in files:
                continue
            baseline_hash = file_hashes.get(baseline_role)
            current_hash = file_hashes.get(current_role)
            if baseline_hash and baseline_hash == current_hash:
                conflicts.append(
                    InputByteConflict(
                        kind="identical_cycle_pair",
                        first_role=baseline_role,
                        second_role=current_role,
                        artifact=artifact,
                        member_id=member_id,
                    )
                )
    return tuple(conflicts)


def reject_duplicate_bytes(
    mode: QCRunMode,
    files: Mapping[str, object],
    file_hashes: Mapping[str, str],
) -> None:
    conflicts = duplicate_byte_conflicts(mode, files, file_hashes)
    if conflicts:
        raise DuplicateInputBytesError("; ".join(item.message() for item in conflicts))


def verify_run_file_hashes(
    files: Mapping[str, Path],
    expected: Mapping[str, str],
    *,
    cancellation_token: CancellationToken | None = None,
) -> None:
    observed = hash_run_files(files, cancellation_token=cancellation_token)
    changed = sorted(role for role in files if observed.get(role) != expected.get(role))
    if changed:
        raise InputBytesChangedError(
            "input bytes changed during QC for role(s): " + ", ".join(changed)
        )
