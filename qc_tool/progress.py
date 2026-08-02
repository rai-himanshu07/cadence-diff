"""Thread-safe progress and cooperative cancellation contracts."""

from __future__ import annotations

import datetime as dt
import importlib
import threading
import time
from collections.abc import Callable
from dataclasses import dataclass
from enum import StrEnum
from typing import Protocol


class RunPhase(StrEnum):
    PREPARING = "preparing"
    LOADING_BASELINE_EXCEL = "loading_baseline_excel"
    LOADING_CURRENT_EXCEL = "loading_current_excel"
    LOADING_BASELINE_POWERPOINT = "loading_baseline_powerpoint"
    LOADING_CURRENT_POWERPOINT = "loading_current_powerpoint"
    ANALYZING_EXCEL = "analyzing_excel"
    DIFFING_EXCEL = "diffing_excel"
    COMPARING_FORMULAS = "comparing_formulas"
    INDEXING_DEPENDENCIES = "indexing_dependencies"
    QUERYING_IMPACTS = "querying_impacts"
    BUILDING_REVIEW = "building_review"
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


class CancellationFlag(Protocol):
    """The surface shared by `threading.Event` and `multiprocessing.Event`."""

    def set(self) -> None: ...

    def is_set(self) -> bool: ...


class CancellationToken:
    """A cancellation flag safe to set from another thread or process.

    Pass a process event to cancel work running in an owned worker process.
    """

    def __init__(self, flag: CancellationFlag | None = None) -> None:
        self._event: CancellationFlag = threading.Event() if flag is None else flag

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


@dataclass(slots=True)
class PhaseRecord:
    """Timing, throughput, and memory for one run phase."""

    phase: RunPhase
    started_at: str
    finished_at: str = ""
    elapsed_seconds: float = 0.0
    processed: int = 0
    total: int = 0
    peak_rss_bytes: int = 0
    error: str = ""


class PhaseTelemetry:
    """Collect per-phase timings and peak RSS from progress events.

    Values are non-secret aggregates only; no file paths, cell values, or
    formula text ever enter a phase record.
    """

    def __init__(self) -> None:
        self.records: dict[RunPhase, PhaseRecord] = {}
        self._started: dict[RunPhase, float] = {}

    @staticmethod
    def _rss() -> int:
        """Peak RSS in bytes where the platform exposes it, else 0."""
        try:
            module = importlib.import_module("resource")
        except ImportError:  # pragma: no cover - Windows has no resource module
            return 0
        usage = module.getrusage(module.RUSAGE_SELF)
        return int(usage.ru_maxrss) * 1024

    def __call__(self, event: ProgressEvent) -> None:
        now = time.perf_counter()
        record = self.records.get(event.phase)
        if record is None:
            self._started[event.phase] = now
            record = PhaseRecord(
                phase=event.phase,
                started_at=dt.datetime.now(dt.UTC).isoformat(timespec="milliseconds"),
            )
            self.records[event.phase] = record
        record.finished_at = dt.datetime.now(dt.UTC).isoformat(timespec="milliseconds")
        record.elapsed_seconds = now - self._started[event.phase]
        record.processed = event.processed
        record.total = event.total
        record.peak_rss_bytes = max(record.peak_rss_bytes, self._rss())

    def fail(self, phase: RunPhase, error: str) -> None:
        """Record a sanitized terminal error against a phase."""
        record = self.records.get(phase)
        if record is not None:
            record.error = error

    def as_payload(self) -> list[dict[str, object]]:
        return [
            {
                "phase": record.phase.value,
                "started_at": record.started_at,
                "finished_at": record.finished_at,
                "elapsed_seconds": round(record.elapsed_seconds, 4),
                "processed": record.processed,
                "total": record.total,
                "peak_rss_bytes": record.peak_rss_bytes,
                "error": record.error,
            }
            for record in self.records.values()
        ]
