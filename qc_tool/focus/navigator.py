"""Parent side of the short-lived focus helper.

One process-global lock serialises every request, so a late reply from a
timed-out helper can never reorder a newer action. Timeouts terminate only the
helper process: an analyst-owned Excel or PowerPoint is never signalled.
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import subprocess
import sys
import tempfile
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

from qc_tool.focus.protocol import (
    DISPATCHED_STAGES,
    SCHEMA_VERSION,
    FocusAction,
    FocusOutcome,
    FocusStage,
    fixed_code,
)
from qc_tool.security import private_directory

logger = logging.getLogger(__name__)

#: Total helper budget for one action, including COM startup.
DEFAULT_HELPER_TIMEOUT_SECONDS = 25.0

#: Step 9 measured 20 dispatched actions: empirical p99/max = 4.141 seconds.
MEASURED_FOCUS_P99_SECONDS = 4.141
#: The approved bound is ``max(30 seconds, 2x measured p99)``.
FOCUS_COOLDOWN_SECONDS = max(30.0, 2 * MEASURED_FOCUS_P99_SECONDS)

#: One process-global lock, shared by every navigator in this server process.
_REQUEST_LOCK = asyncio.Lock()


@dataclass(frozen=True, slots=True)
class HelperRun:
    """Raw helper execution facts before they become a navigator reply."""

    timed_out: bool = False
    stage: FocusStage | None = None
    payload: dict[str, object] | None = None
    exit_code: int | None = 0
    duration_seconds: float = 0.0


@dataclass(frozen=True, slots=True)
class FocusReply:
    """Fixed-code result handed back to the UI. Never a path or a message."""

    outcome: FocusOutcome | str
    stage: FocusStage | None = None
    duration_seconds: float = 0.0
    unsaved_changes: bool = False
    discovery: dict[str, object] | None = None
    detail: dict[str, object] = field(default_factory=dict)

    @property
    def code(self) -> str:
        return fixed_code(self.outcome)

    @property
    def dispatched(self) -> bool:
        return self.stage in DISPATCHED_STAGES


HelperRunner = Callable[[dict[str, object], float], HelperRun]


def _read_stage(stage_path: Path) -> FocusStage | None:
    try:
        raw = stage_path.read_text(encoding="utf-8").strip()
    except OSError:
        return None
    try:
        return FocusStage(raw)
    except ValueError:
        return None


def _read_result(result_path: Path) -> dict[str, object] | None:
    try:
        payload = json.loads(result_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(payload, dict) or payload.get("schema_version") != SCHEMA_VERSION:
        return None
    return payload


def run_helper(request: dict[str, object], timeout: float) -> HelperRun:
    """Run one helper process and bound its total time."""
    started = time.monotonic()
    with tempfile.TemporaryDirectory(prefix="qc-tool-focus-") as temporary:
        work_dir = private_directory(Path(temporary))
        stage_path = work_dir / "stage.txt"
        result_path = work_dir / "result.json"
        payload = dict(request)
        payload["stage_path"] = str(stage_path)
        payload["result_path"] = str(result_path)
        creation_flags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        process = subprocess.Popen(
            [sys.executable, "-m", "qc_tool.focus.worker"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=os.environ.copy(),
            close_fds=True,
            creationflags=creation_flags,
        )
        timed_out = False
        try:
            process.communicate(json.dumps(payload), timeout=timeout)
        except subprocess.TimeoutExpired:
            timed_out = True
            # Kill the helper only. Analyst-owned Office is never signalled.
            process.kill()
            process.communicate()
        return HelperRun(
            timed_out=timed_out,
            stage=_read_stage(stage_path),
            payload=_read_result(result_path),
            exit_code=process.returncode,
            duration_seconds=time.monotonic() - started,
        )


class FocusNavigator:
    """Serialised, bounded dispatcher for every desktop focus action."""

    def __init__(
        self,
        *,
        runner: HelperRunner | None = None,
        timeout: float = DEFAULT_HELPER_TIMEOUT_SECONDS,
        cooldown_seconds: float = FOCUS_COOLDOWN_SECONDS,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        self._runner = runner or run_helper
        self._timeout = timeout
        self._cooldown = max(cooldown_seconds, FOCUS_COOLDOWN_SECONDS)
        self._clock = clock
        self._lock = _REQUEST_LOCK
        self._cooldown_until: float | None = None
        self._health_check_pending = False

    @property
    def disabled(self) -> bool:
        """Compatibility property; measured recovery no longer disables focus."""
        return False

    def cooling_down(self) -> bool:
        if self._cooldown_until is None:
            return False
        if self._clock() >= self._cooldown_until:
            self._cooldown_until = None
            return False
        return True

    async def submit(self, request: dict[str, object]) -> FocusReply:
        async with self._lock:
            return await asyncio.to_thread(self._submit_locked, request)

    def _submit_locked(self, request: dict[str, object]) -> FocusReply:
        if self.cooling_down():
            return FocusReply(FocusOutcome.COOLING_DOWN)
        if self._health_check_pending:
            self._health_check_pending = False
            health = self._dispatch(
                {
                    "schema_version": SCHEMA_VERSION,
                    "action": FocusAction.HEALTH_CHECK.value,
                    "application": request.get("application"),
                }
            )
            if health.outcome is not FocusOutcome.HEALTHY:
                return health
        return self._dispatch(request)

    def _dispatch(self, request: dict[str, object]) -> FocusReply:
        run = self._runner(request, self._timeout)
        if run.timed_out:
            return self._timeout_reply(run)
        if run.payload is None:
            outcome = (
                FocusOutcome.HELPER_CRASHED
                if run.exit_code not in (0, None)
                else FocusOutcome.HELPER_FAILED
            )
            logger.warning("focus-helper-%s", outcome.value)
            return FocusReply(
                outcome, stage=run.stage, duration_seconds=run.duration_seconds
            )
        return self._reply(run)

    def _timeout_reply(self, run: HelperRun) -> FocusReply:
        self._cooldown_until = self._clock() + self._cooldown
        self._health_check_pending = True
        if run.stage in DISPATCHED_STAGES:
            logger.warning("focus-timeout-after-dispatch")
            return FocusReply(
                FocusOutcome.TIMEOUT_ACTION_MAY_HAVE_COMPLETED,
                stage=run.stage,
                duration_seconds=run.duration_seconds,
            )
        logger.warning("focus-timeout-before-dispatch")
        return FocusReply(
            FocusOutcome.TIMEOUT_NO_ACTION_DISPATCHED,
            stage=run.stage,
            duration_seconds=run.duration_seconds,
        )

    def _reply(self, run: HelperRun) -> FocusReply:
        payload = run.payload or {}
        raw_outcome = payload.get("outcome")
        try:
            outcome: FocusOutcome | str = FocusOutcome(raw_outcome)
        except ValueError:
            outcome = fixed_code(raw_outcome)
        discovery = payload.get("discovery")
        return FocusReply(
            outcome=outcome,
            stage=run.stage,
            duration_seconds=run.duration_seconds,
            unsaved_changes=payload.get("unsaved_changes") is True,
            discovery=discovery if isinstance(discovery, dict) else None,
            detail={
                key: value
                for key, value in payload.items()
                if key in {"complete", "reasons"}
            },
        )
