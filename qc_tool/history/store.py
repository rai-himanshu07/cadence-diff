"""Persistent QC run history (local SQLite, timezone-aware UTC).

Each record stores the run metadata, SHA256 hashes of the compared files,
severity counts, report paths, and the full findings payload so past runs
can be reopened in the UI without re-running the comparison.
"""

import datetime as dt
import hashlib
import json
import logging
import sqlite3
import zipfile
from collections.abc import Iterable
from dataclasses import dataclass, field
from pathlib import Path

from qc_tool.config.profile import (
    DeliverableProfile,
    canonical_profile_json,
    profile_sha256,
)
from qc_tool.coverage import CoverageItem, MappingCoverage, QCRunMode
from qc_tool.crosscheck.trace import MappingSuggestion
from qc_tool.engine import QCRunResult
from qc_tool.findings import Finding, Severity
from qc_tool.focus.model import (
    EMPTY_SIDECAR_JSON,
    FocusTargetSeed,
    FocusTargetSidecar,
    decode_focus_targets,
    encode_focus_targets,
)
from qc_tool.history.review_state import (
    AnnotationLineage,
    CarryForwardCandidate,
    RunFinalizedError,
    RunSignoff,
)
from qc_tool.review import build_pattern_groups, build_review_groups
from qc_tool.review import count_pattern_groups as count_pattern_review_groups
from qc_tool.review import review_counts as count_review_groups
from qc_tool.scope import ComparisonScope
from qc_tool.security import private_directory, private_file
from qc_tool.story import build_stories

logger = logging.getLogger(__name__)

_SCHEMA = """
CREATE TABLE IF NOT EXISTS runs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    started_at TEXT NOT NULL,
    profile TEXT NOT NULL,
    files TEXT NOT NULL,
    file_hashes TEXT NOT NULL,
    counts TEXT NOT NULL,
    disclosures TEXT NOT NULL,
    verified_crosschecks INTEGER NOT NULL,
    findings TEXT NOT NULL,
    report_paths TEXT NOT NULL,
    file_paths TEXT NOT NULL DEFAULT '{}',
    rerun_of INTEGER,
    mode TEXT NOT NULL DEFAULT 'cycle_comparison',
    coverage TEXT NOT NULL DEFAULT '[]',
    mapping_coverage TEXT NOT NULL DEFAULT 'null',
    mapping_suggestions TEXT NOT NULL DEFAULT '[]',
    review_counts TEXT NOT NULL DEFAULT '{}',
    pattern_review_counts TEXT NOT NULL DEFAULT '{}',
    story_counts TEXT NOT NULL DEFAULT '{}',
    comparison_scope TEXT NOT NULL DEFAULT '{}',
    focus_targets TEXT NOT NULL DEFAULT '{}'
    ,profile_snapshot TEXT NOT NULL DEFAULT 'null'
    ,profile_sha256 TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS annotations (
    run_id INTEGER NOT NULL,
    finding_id TEXT NOT NULL,
    severity TEXT,
    comment TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL,
    PRIMARY KEY (run_id, finding_id)
);
CREATE TABLE IF NOT EXISTS run_signoffs (
    run_id INTEGER PRIMARY KEY,
    finalized_at TEXT NOT NULL,
    acknowledgements TEXT NOT NULL DEFAULT '[]',
    review_state_digest TEXT NOT NULL,
    profile_sha256 TEXT NOT NULL,
    attestation_path TEXT NOT NULL,
    attestation_sha256 TEXT NOT NULL,
    report_paths TEXT NOT NULL DEFAULT '{}'
);
CREATE TABLE IF NOT EXISTS annotation_lineage (
    run_id INTEGER NOT NULL,
    finding_id TEXT NOT NULL,
    source_run_id INTEGER NOT NULL,
    source_finding_id TEXT NOT NULL,
    evidence_version INTEGER NOT NULL,
    evidence_digest TEXT NOT NULL,
    applied_at TEXT NOT NULL,
    PRIMARY KEY (run_id, finding_id)
);
CREATE TABLE IF NOT EXISTS review_sessions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id INTEGER NOT NULL,
    started_at TEXT NOT NULL,
    ended_at TEXT,
    active_seconds REAL
);
"""

#: Columns added after the first release; applied to legacy databases.
_MIGRATIONS = {
    "file_paths": "ALTER TABLE runs ADD COLUMN file_paths TEXT NOT NULL DEFAULT '{}'",
    "rerun_of": "ALTER TABLE runs ADD COLUMN rerun_of INTEGER",
    "mode": (
        "ALTER TABLE runs ADD COLUMN mode TEXT NOT NULL DEFAULT 'cycle_comparison'"
    ),
    "coverage": "ALTER TABLE runs ADD COLUMN coverage TEXT NOT NULL DEFAULT '[]'",
    "mapping_coverage": (
        "ALTER TABLE runs ADD COLUMN mapping_coverage TEXT NOT NULL DEFAULT 'null'"
    ),
    "mapping_suggestions": (
        "ALTER TABLE runs ADD COLUMN mapping_suggestions TEXT NOT NULL DEFAULT '[]'"
    ),
    "review_counts": (
        "ALTER TABLE runs ADD COLUMN review_counts TEXT NOT NULL DEFAULT '{}'"
    ),
    "pattern_review_counts": (
        "ALTER TABLE runs ADD COLUMN pattern_review_counts TEXT NOT NULL DEFAULT '{}'"
    ),
    "story_counts": (
        "ALTER TABLE runs ADD COLUMN story_counts TEXT NOT NULL DEFAULT '{}'"
    ),
    "comparison_scope": (
        "ALTER TABLE runs ADD COLUMN comparison_scope TEXT NOT NULL DEFAULT '{}'"
    ),
    "archived": "ALTER TABLE runs ADD COLUMN archived INTEGER NOT NULL DEFAULT 0",
    "focus_targets": (
        "ALTER TABLE runs ADD COLUMN focus_targets TEXT NOT NULL DEFAULT '{}'"
    ),
    "profile_snapshot": (
        "ALTER TABLE runs ADD COLUMN profile_snapshot TEXT NOT NULL DEFAULT 'null'"
    ),
    "profile_sha256": (
        "ALTER TABLE runs ADD COLUMN profile_sha256 TEXT NOT NULL DEFAULT ''"
    ),
}


@dataclass(slots=True)
class RunRecord:
    run_id: int
    started_at: dt.datetime
    profile: str
    mode: QCRunMode
    files: dict[str, str]
    file_hashes: dict[str, str]
    counts: dict[str, int]
    review_counts: dict[str, int]
    disclosures: list[str]
    verified_crosschecks: int
    report_paths: dict[str, str]
    #: Semantic pattern review counts; ``{}`` for runs recorded before Step 8.
    pattern_review_counts: dict[str, int] = field(default_factory=dict)
    #: Change-story member counts by kind; ``{}`` for earlier runs.
    story_counts: dict[str, dict[str, int]] = field(default_factory=dict)
    comparison_scope: ComparisonScope = field(default_factory=ComparisonScope)
    file_paths: dict[str, str] = field(default_factory=dict)  # role -> stored path
    rerun_of: int | None = None
    #: Retired from the default history view; the record itself is retained.
    archived: bool = False
    coverage: list[CoverageItem] = field(default_factory=list)
    mapping_coverage: MappingCoverage | None = None
    mapping_suggestions: list[MappingSuggestion] = field(default_factory=list)
    findings: list[Finding] = field(default_factory=list)
    #: Private desktop-focus locators; never exported, reported, or attested.
    focus_targets: FocusTargetSidecar = field(default_factory=FocusTargetSidecar)
    profile_snapshot: DeliverableProfile | None = None
    profile_sha256: str = ""
    signoff: RunSignoff | None = None


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _remove_managed_report(path: Path, root: Path) -> None:
    """Delete one report file, and its now-empty run directory, inside ``root``."""
    try:
        resolved = path.resolve()
        resolved.relative_to(root)
    except (OSError, ValueError):
        logger.warning("refusing to delete a report outside the managed data directory")
        return
    try:
        resolved.unlink(missing_ok=True)
        parent = resolved.parent
        if parent != root and not any(parent.iterdir()):
            parent.rmdir()
    except OSError:
        logger.warning("could not remove a stored report file")


class RunHistory:
    def __init__(self, db_path: Path) -> None:
        private_directory(db_path.parent)
        self._db_path = db_path
        with self._connect() as conn:
            conn.executescript(_SCHEMA)
            existing = {row[1] for row in conn.execute("PRAGMA table_info(runs)")}
            for column, statement in _MIGRATIONS.items():
                if column not in existing:
                    conn.execute(statement)
            private_file(db_path)

    def _connect(self) -> sqlite3.Connection:
        # The worker process, the queue manager, and the UI share this file.
        conn = sqlite3.connect(self._db_path, timeout=30.0)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.row_factory = sqlite3.Row
        return conn

    def record_run(
        self,
        result: QCRunResult,
        *,
        file_hashes: dict[str, str],
        report_paths: dict[str, str],
        file_paths: dict[str, str] | None = None,
        rerun_of: int | None = None,
        focus_targets: dict[str, tuple[FocusTargetSeed, ...]] | None = None,
        profile_snapshot: DeliverableProfile | None = None,
    ) -> int:
        started_at = dt.datetime.now(dt.UTC).isoformat(timespec="seconds")
        counts = {sev.value: count for sev, count in result.counts.items()}
        grouped = count_review_groups(build_review_groups(result.findings))
        grouped_counts = {
            severity.value: count
            for severity, count in grouped.review_items.items()
        }
        patterned = count_pattern_review_groups(build_pattern_groups(result.findings))
        pattern_counts = {
            severity.value: count
            for severity, count in patterned.review_items.items()
        }
        story_counts: dict[str, dict[str, int]] = {}
        for story in build_stories(result.findings):
            bucket = story_counts.setdefault(
                story.kind.value, {"stories": 0, "members": 0}
            )
            bucket["stories"] += 1
            bucket["members"] += story.member_count
        findings_payload = json.dumps(
            [finding.model_dump(mode="json") for finding in result.findings]
        )
        profile_payload = (
            canonical_profile_json(profile_snapshot)
            if profile_snapshot is not None
            else "null"
        )
        profile_digest = (
            profile_sha256(profile_snapshot)
            if profile_snapshot is not None
            else ""
        )
        with self._connect() as conn:
            cursor = conn.execute(
                """
                INSERT INTO runs (
                    started_at, profile, files, file_hashes, counts,
                    disclosures, verified_crosschecks, findings, report_paths,
                    file_paths, rerun_of, mode, coverage, mapping_coverage,
                    mapping_suggestions, review_counts, pattern_review_counts,
                    story_counts, comparison_scope, focus_targets,
                    profile_snapshot, profile_sha256
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    started_at,
                    result.profile_name,
                    json.dumps(result.files),
                    json.dumps(file_hashes),
                    json.dumps(counts),
                    json.dumps(result.disclosures),
                    result.verified_crosschecks,
                    findings_payload,
                    json.dumps(report_paths),
                    json.dumps(file_paths or {}),
                    rerun_of,
                    result.mode.value,
                    json.dumps([item.model_dump(mode="json") for item in result.coverage]),
                    json.dumps(
                        result.mapping_coverage.model_dump(mode="json")
                        if result.mapping_coverage is not None
                        else None
                    ),
                    json.dumps(
                        [item.model_dump(mode="json") for item in result.mapping_suggestions]
                    ),
                    json.dumps(grouped_counts),
                    json.dumps(pattern_counts),
                    json.dumps(story_counts),
                    json.dumps(result.comparison_scope.model_dump(mode="json")),
                    (
                        encode_focus_targets(focus_targets)
                        if focus_targets
                        else EMPTY_SIDECAR_JSON
                    ),
                    profile_payload,
                    profile_digest,
                ),
            )
            run_id = cursor.lastrowid
        if run_id is None:  # pragma: no cover - sqlite always returns a rowid
            raise RuntimeError("sqlite did not return a run id")
        logger.info("recorded QC run %d (%s)", run_id, result.profile_name)
        return run_id

    def _record_from_row(self, row: sqlite3.Row, *, with_findings: bool) -> RunRecord:
        findings = (
            [Finding.model_validate(item) for item in json.loads(row["findings"])]
            if with_findings
            else []
        )
        return RunRecord(
            run_id=row["id"],
            started_at=dt.datetime.fromisoformat(row["started_at"]),
            profile=row["profile"],
            mode=QCRunMode(row["mode"]),
            files=json.loads(row["files"]),
            file_hashes=json.loads(row["file_hashes"]),
            counts=json.loads(row["counts"]),
            review_counts=json.loads(row["review_counts"]),
            pattern_review_counts=json.loads(row["pattern_review_counts"]),
            story_counts=json.loads(row["story_counts"]),
            comparison_scope=ComparisonScope.model_validate(
                json.loads(row["comparison_scope"])
            ),
            disclosures=json.loads(row["disclosures"]),
            verified_crosschecks=row["verified_crosschecks"],
            report_paths=json.loads(row["report_paths"]),
            file_paths=json.loads(row["file_paths"]),
            rerun_of=row["rerun_of"],
            archived=bool(row["archived"]),
            coverage=[CoverageItem.model_validate(item) for item in json.loads(row["coverage"])],
            mapping_coverage=(
                MappingCoverage.model_validate(json.loads(row["mapping_coverage"]))
                if json.loads(row["mapping_coverage"]) is not None
                else None
            ),
            mapping_suggestions=[
                MappingSuggestion.model_validate(item)
                for item in json.loads(row["mapping_suggestions"])
            ],
            findings=findings,
            focus_targets=decode_focus_targets(row["focus_targets"]),
            profile_snapshot=(
                DeliverableProfile.model_validate(json.loads(row["profile_snapshot"]))
                if json.loads(row["profile_snapshot"]) is not None
                else None
            ),
            profile_sha256=row["profile_sha256"],
        )

    def list_runs(self, limit: int = 50, *, include_archived: bool = True) -> list[RunRecord]:
        clause = "" if include_archived else " WHERE archived = 0"
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT * FROM runs{clause} ORDER BY id DESC LIMIT ?", (limit,)
            ).fetchall()
        records = [self._record_from_row(row, with_findings=False) for row in rows]
        signoffs = self._signoffs_by_run(record.run_id for record in records)
        for record in records:
            record.signoff = signoffs.get(record.run_id)
        return records

    def set_archived(self, run_ids: Iterable[int], archived: bool) -> int:
        """Retire or restore runs without touching their recorded evidence."""
        ids = sorted({int(run_id) for run_id in run_ids})
        if not ids:
            return 0
        placeholders = ",".join("?" for _ in ids)
        with self._connect() as conn:
            cursor = conn.execute(
                f"UPDATE runs SET archived = ? WHERE id IN ({placeholders})",
                (1 if archived else 0, *ids),
            )
            return cursor.rowcount

    def delete_runs(self, run_ids: Iterable[int], *, managed_root: Path) -> int:
        """Delete runs, their annotations, and their report files.

        Report files are removed only when they resolve inside ``managed_root``,
        so a tampered or legacy path can never delete anything outside the
        managed data directory.
        """
        ids = sorted({int(run_id) for run_id in run_ids})
        if not ids:
            return 0
        root = managed_root.resolve()
        placeholders = ",".join("?" for _ in ids)
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT report_paths FROM runs WHERE id IN ({placeholders})", ids
            ).fetchall()
            signoff_rows = conn.execute(
                f"""
                SELECT report_paths, attestation_path
                FROM run_signoffs WHERE run_id IN ({placeholders})
                """,
                ids,
            ).fetchall()
            cursor = conn.execute(
                f"DELETE FROM runs WHERE id IN ({placeholders})", ids
            )
            conn.execute(
                f"DELETE FROM annotations WHERE run_id IN ({placeholders})", ids
            )
            conn.execute(
                f"DELETE FROM annotation_lineage WHERE run_id IN ({placeholders})",
                ids,
            )
            conn.execute(
                f"DELETE FROM run_signoffs WHERE run_id IN ({placeholders})", ids
            )
            conn.execute(
                f"DELETE FROM review_sessions WHERE run_id IN ({placeholders})", ids
            )
            deleted = cursor.rowcount
        for row in rows:
            for raw in json.loads(row["report_paths"]).values():
                _remove_managed_report(Path(raw), root)
        for row in signoff_rows:
            for raw in (
                *json.loads(row["report_paths"]).values(),
                row["attestation_path"],
            ):
                _remove_managed_report(Path(raw), root)
        return deleted

    def get_run(self, run_id: int) -> RunRecord:
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM runs WHERE id = ?", (run_id,)).fetchone()
        if row is None:
            raise KeyError(f"no QC run with id {run_id}")
        record = self._record_from_row(row, with_findings=True)
        record.signoff = self.get_signoff(run_id)
        annotations = self.get_annotations(run_id)
        for finding in record.findings:
            annotation = annotations.get(finding.finding_id)
            if annotation is None:
                continue
            severity, comment = annotation
            finding.analyst_comment = comment
            if severity is not None:
                finding.severity = Severity(severity)
                finding.severity_overridden = True
        return record

    def get_raw_run(self, run_id: int) -> RunRecord:
        """Rehydrate immutable engine findings without analyst annotations."""
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM runs WHERE id = ?", (run_id,)).fetchone()
        if row is None:
            raise KeyError(f"no QC run with id {run_id}")
        record = self._record_from_row(row, with_findings=True)
        record.signoff = self.get_signoff(run_id)
        return record

    def set_annotation(
        self, run_id: int, finding_id: str, *, severity: str | None, comment: str
    ) -> None:
        """Upsert an analyst annotation (severity=None keeps the engine severity)."""
        self.set_annotations_bulk(
            run_id,
            [(finding_id, severity, comment)],
        )

    def set_annotations_bulk(
        self,
        run_id: int,
        updates: list[tuple[str, str | None, str]],
    ) -> None:
        """Upsert one group decision atomically across its finding members."""
        if not updates:
            return
        self.assert_mutable(run_id)
        updated_at = dt.datetime.now(dt.UTC).isoformat(timespec="seconds")
        with self._connect() as conn:
            conn.executemany(
                """
                INSERT INTO annotations (run_id, finding_id, severity, comment, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(run_id, finding_id) DO UPDATE SET
                    severity = excluded.severity,
                    comment = excluded.comment,
                    updated_at = excluded.updated_at
                """,
                [
                    (run_id, finding_id, severity, comment, updated_at)
                    for finding_id, severity, comment in updates
                ],
            )

    def get_annotations(self, run_id: int) -> dict[str, tuple[str | None, str]]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT finding_id, severity, comment FROM annotations WHERE run_id = ?",
                (run_id,),
            ).fetchall()
        return {row["finding_id"]: (row["severity"], row["comment"]) for row in rows}

    def update_mapping_review(
        self,
        run_id: int,
        *,
        coverage: MappingCoverage,
        suggestions: list[MappingSuggestion],
        check_coverage: list[CoverageItem],
    ) -> None:
        """Persist mapping confirmations made while reviewing a stored run."""
        self.assert_mutable(run_id)
        with self._connect() as conn:
            cursor = conn.execute(
                """
                UPDATE runs
                SET mapping_coverage = ?, mapping_suggestions = ?,
                    verified_crosschecks = ?, coverage = ?
                WHERE id = ?
                """,
                (
                    json.dumps(coverage.model_dump(mode="json")),
                    json.dumps(
                        [suggestion.model_dump(mode="json") for suggestion in suggestions]
                    ),
                    coverage.verified,
                    json.dumps(
                        [item.model_dump(mode="json") for item in check_coverage]
                    ),
                    run_id,
                ),
            )
        if cursor.rowcount == 0:
            raise KeyError(f"no QC run with id {run_id}")

    def _signoffs_by_run(self, run_ids: Iterable[int]) -> dict[int, RunSignoff]:
        ids = sorted({int(run_id) for run_id in run_ids})
        if not ids:
            return {}
        placeholders = ",".join("?" for _ in ids)
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT * FROM run_signoffs WHERE run_id IN ({placeholders})", ids
            ).fetchall()
        return {row["run_id"]: self._signoff_from_row(row) for row in rows}

    @staticmethod
    def _signoff_from_row(row: sqlite3.Row) -> RunSignoff:
        return RunSignoff(
            run_id=row["run_id"],
            finalized_at=row["finalized_at"],
            acknowledgements=tuple(json.loads(row["acknowledgements"])),
            review_state_digest=row["review_state_digest"],
            profile_sha256=row["profile_sha256"],
            attestation_path=row["attestation_path"],
            attestation_sha256=row["attestation_sha256"],
            report_paths=json.loads(row["report_paths"]),
        )

    def get_signoff(self, run_id: int) -> RunSignoff | None:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM run_signoffs WHERE run_id = ?", (run_id,)
            ).fetchone()
        return None if row is None else self._signoff_from_row(row)

    def assert_mutable(self, run_id: int) -> None:
        if self.get_signoff(run_id) is not None:
            raise RunFinalizedError(
                f"run #{run_id} is finalized; submit a Re-QC run to make corrections"
            )

    def record_signoff(self, signoff: RunSignoff) -> None:
        with self._connect() as conn:
            try:
                conn.execute(
                    """
                    INSERT INTO run_signoffs (
                        run_id, finalized_at, acknowledgements,
                        review_state_digest, profile_sha256, attestation_path,
                        attestation_sha256, report_paths
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        signoff.run_id,
                        signoff.finalized_at,
                        json.dumps(list(signoff.acknowledgements)),
                        signoff.review_state_digest,
                        signoff.profile_sha256,
                        signoff.attestation_path,
                        signoff.attestation_sha256,
                        json.dumps(signoff.report_paths),
                    ),
                )
            except sqlite3.IntegrityError as exc:
                raise RunFinalizedError(
                    f"run #{signoff.run_id} is already finalized"
                ) from exc

    def get_annotation_lineage(self, run_id: int) -> dict[str, AnnotationLineage]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM annotation_lineage WHERE run_id = ?", (run_id,)
            ).fetchall()
        return {
            row["finding_id"]: AnnotationLineage(
                finding_id=row["finding_id"],
                source_run_id=row["source_run_id"],
                source_finding_id=row["source_finding_id"],
                evidence_version=row["evidence_version"],
                evidence_digest=row["evidence_digest"],
                applied_at=row["applied_at"],
            )
            for row in rows
        }

    def apply_carried_annotations(
        self,
        run_id: int,
        source_run_id: int,
        candidates: list[CarryForwardCandidate],
    ) -> int:
        if not candidates:
            return 0
        applied_at = dt.datetime.now(dt.UTC).isoformat(timespec="seconds")
        with self._connect() as conn:
            if conn.execute(
                "SELECT 1 FROM run_signoffs WHERE run_id = ?", (run_id,)
            ).fetchone():
                raise RunFinalizedError(
                    f"run #{run_id} is finalized; submit a Re-QC run to make corrections"
                )
            rerun = conn.execute(
                "SELECT rerun_of FROM runs WHERE id = ?", (run_id,)
            ).fetchone()
            if rerun is None:
                raise KeyError(f"no QC run with id {run_id}")
            if rerun["rerun_of"] != source_run_id:
                raise ValueError("carry-forward source is not this run's direct predecessor")
            conn.executemany(
                """
                INSERT INTO annotations (run_id, finding_id, severity, comment, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(run_id, finding_id) DO UPDATE SET
                    severity = excluded.severity,
                    comment = excluded.comment,
                    updated_at = excluded.updated_at
                """,
                [
                    (
                        run_id,
                        candidate.finding_id,
                        candidate.severity,
                        candidate.comment,
                        applied_at,
                    )
                    for candidate in candidates
                ],
            )
            conn.executemany(
                """
                INSERT INTO annotation_lineage (
                    run_id, finding_id, source_run_id, source_finding_id,
                    evidence_version, evidence_digest, applied_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(run_id, finding_id) DO UPDATE SET
                    source_run_id = excluded.source_run_id,
                    source_finding_id = excluded.source_finding_id,
                    evidence_version = excluded.evidence_version,
                    evidence_digest = excluded.evidence_digest,
                    applied_at = excluded.applied_at
                """,
                [
                    (
                        run_id,
                        candidate.finding_id,
                        source_run_id,
                        candidate.source_finding_id,
                        candidate.evidence_version,
                        candidate.evidence_digest,
                        applied_at,
                    )
                    for candidate in candidates
                ],
            )
        return len(candidates)

    @staticmethod
    def _bounded_session_end(started_at: dt.datetime, observed: dt.datetime) -> dt.datetime:
        return min(observed, started_at + dt.timedelta(hours=4))

    def pause_review_sessions(self, *, now: dt.datetime | None = None) -> int:
        observed = now or dt.datetime.now(dt.UTC)
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT id, started_at FROM review_sessions WHERE ended_at IS NULL"
            ).fetchall()
            for row in rows:
                started = dt.datetime.fromisoformat(row["started_at"])
                ended = self._bounded_session_end(started, observed)
                seconds = max(0.0, (ended - started).total_seconds())
                conn.execute(
                    """
                    UPDATE review_sessions
                    SET ended_at = ?, active_seconds = ? WHERE id = ?
                    """,
                    (ended.isoformat(timespec="seconds"), seconds, row["id"]),
                )
        return len(rows)

    def start_review_session(
        self,
        run_id: int,
        *,
        now: dt.datetime | None = None,
    ) -> int:
        observed = now or dt.datetime.now(dt.UTC)
        self.pause_review_sessions(now=observed)
        with self._connect() as conn:
            if conn.execute("SELECT 1 FROM runs WHERE id = ?", (run_id,)).fetchone() is None:
                raise KeyError(f"no QC run with id {run_id}")
            cursor = conn.execute(
                "INSERT INTO review_sessions (run_id, started_at) VALUES (?, ?)",
                (run_id, observed.isoformat(timespec="seconds")),
            )
        if cursor.lastrowid is None:
            raise RuntimeError("sqlite did not return a review-session id")
        return int(cursor.lastrowid)

    def active_review_run(self) -> int | None:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT run_id FROM review_sessions
                WHERE ended_at IS NULL ORDER BY id DESC LIMIT 1
                """
            ).fetchone()
        return None if row is None else int(row["run_id"])

    def review_seconds(
        self,
        run_id: int,
        *,
        now: dt.datetime | None = None,
    ) -> float | None:
        observed = now or dt.datetime.now(dt.UTC)
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT started_at, ended_at, active_seconds
                FROM review_sessions WHERE run_id = ?
                """,
                (run_id,),
            ).fetchall()
        if not rows:
            return None
        total = 0.0
        for row in rows:
            if row["ended_at"] is not None:
                total += float(row["active_seconds"] or 0.0)
                continue
            started = dt.datetime.fromisoformat(row["started_at"])
            ended = self._bounded_session_end(started, observed)
            total += max(0.0, (ended - started).total_seconds())
        return total

    def carried_annotation_count(self, run_id: int) -> int:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT COUNT(*) FROM annotation_lineage WHERE run_id = ?", (run_id,)
            ).fetchone()
        return int(row[0]) if row is not None else 0


def export_runs_archive(
    records: list[RunRecord], destination: Path, *, managed_root: Path
) -> Path:
    """Bundle the stored reports for several runs into one private zip.

    Entry names are generated here, never taken from stored paths, and a source
    file is included only when it resolves inside ``managed_root``. The private
    focus-target sidecar is deliberately absent from both the entries and the
    manifest: it is local operational metadata, not shareable evidence.
    """
    root = managed_root.resolve()
    manifest: list[dict[str, object]] = []
    private_directory(destination.parent)
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as bundle:
        for record in records:
            included: list[str] = []
            artifact_paths = (
                record.signoff.report_paths
                if record.signoff is not None
                else record.report_paths
            )
            if record.signoff is not None:
                artifact_paths = {
                    **artifact_paths,
                    "attestation": record.signoff.attestation_path,
                }
            for kind, raw in sorted(artifact_paths.items()):
                source = Path(raw)
                try:
                    resolved = source.resolve()
                    resolved.relative_to(root)
                except (OSError, ValueError):
                    logger.warning("skipping a report outside the managed directory")
                    continue
                if not resolved.is_file():
                    continue
                entry = f"run-{record.run_id}/{kind}{resolved.suffix}"
                bundle.write(resolved, entry)
                included.append(entry)
            manifest.append(
                {
                    "run_id": record.run_id,
                    "started_at": record.started_at.isoformat(timespec="seconds"),
                    "mode": record.mode.value,
                    "profile": record.profile,
                    "files": record.files,  # display names only, never paths
                    "counts": record.counts,
                    "pattern_review_counts": record.pattern_review_counts,
                    "archived": record.archived,
                    "finalized_at": (
                        record.signoff.finalized_at
                        if record.signoff is not None
                        else None
                    ),
                    "reports": included,
                }
            )
        bundle.writestr(
            "manifest.json",
            json.dumps(
                {
                    "exported_at": dt.datetime.now(dt.UTC).isoformat(
                        timespec="seconds"
                    ),
                    "runs": manifest,
                },
                indent=2,
            ),
        )
    private_file(destination)
    return destination
