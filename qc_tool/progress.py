"""Thread-safe progress and cooperative cancellation contracts."""

from __future__ import annotations

import threading
from collections.abc import Callable
from dataclasses import dataclass
from enum import StrEnum


class RunPhase(StrEnum):
    PREPARING = "preparing"
    LOADING_BASELINE_EXCEL = "loading_baseline_excel"
    LOADING_CURRENT_EXCEL = "loading_current_excel"
    LOADING_BASELINE_POWERPOINT = "loading_baseline_powerpoint"
    LOADING_CURRENT_POWERPOINT = "loading_current_powerpoint"
    ANALYZING_EXCEL = "analyzing_excel"
    ANALYZING_POWERPOINT = "analyzing_powerpoint"
    CROSSCHECKING = "crosschecking"
    WRITING_REPORTS = "writing_reports"
    RECORDING_HISTORY = "recording_history"
    COMPLETE = "complete"


@dataclass(frozen=True, slots=True)
class ProgressEvent:
    phase: RunPhase
    processed: int = 0
    total: int = 0
    detail: str = ""

    def __post_init__(self) -> None:
        if self.processed < 0 or self.total < 0:
            raise ValueError("progress counts cannot be negative")
        if self.total and self.processed > self.total:
            raise ValueError("processed progress cannot exceed total")


ProgressCallback = Callable[[ProgressEvent], None]


class RunCancelled(RuntimeError):
    """A run stopped at a cooperative cancellation boundary."""


class CancellationToken:
    """A cancellation flag safe to set from another thread."""

    def __init__(self) -> None:
        self._event = threading.Event()

    def cancel(self) -> None:
        self._event.set()

    @property
    def is_cancelled(self) -> bool:
        return self._event.is_set()

    def check(self) -> None:
        if self._event.is_set():
            raise RunCancelled("QC run cancelled")


def check_cancelled(token: CancellationToken | None) -> None:
    if token is not None:
        token.check()


def report_progress(
    callback: ProgressCallback | None,
    phase: RunPhase,
    *,
    processed: int = 0,
    total: int = 0,
    detail: str = "",
) -> None:
    if callback is not None:
        callback(ProgressEvent(phase, processed, total, detail))
