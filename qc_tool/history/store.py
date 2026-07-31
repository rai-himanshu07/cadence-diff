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
from dataclasses import dataclass, field
from pathlib import Path

from qc_tool.coverage import CoverageItem, MappingCoverage, QCRunMode
from qc_tool.crosscheck.trace import MappingSuggestion
from qc_tool.engine import QCRunResult
from qc_tool.findings import Finding, Severity
from qc_tool.review import build_review_groups
from qc_tool.review import review_counts as count_review_groups
from qc_tool.security import private_directory, private_file

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
    review_counts TEXT NOT NULL DEFAULT '{}'
);
CREATE TABLE IF NOT EXISTS annotations (
    run_id INTEGER NOT NULL,
    finding_id TEXT NOT NULL,
    severity TEXT,
    comment TEXT NOT NULL DEFAULT '',
    updated_at TEXT NOT NULL,
    PRIMARY KEY (run_id, finding_id)
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
    file_paths: dict[str, str] = field(default_factory=dict)  # role -> stored path
    rerun_of: int | None = None
    coverage: list[CoverageItem] = field(default_factory=list)
    mapping_coverage: MappingCoverage | None = None
    mapping_suggestions: list[MappingSuggestion] = field(default_factory=list)
    findings: list[Finding] = field(default_factory=list)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


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
        conn = sqlite3.connect(self._db_path)
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
    ) -> int:
        started_at = dt.datetime.now(dt.UTC).isoformat(timespec="seconds")
        counts = {sev.value: count for sev, count in result.counts.items()}
        grouped = count_review_groups(build_review_groups(result.findings))
        grouped_counts = {
            severity.value: count
            for severity, count in grouped.review_items.items()
        }
        findings_payload = json.dumps(
            [finding.model_dump(mode="json") for finding in result.findings]
        )
        with self._connect() as conn:
            cursor = conn.execute(
                """
                INSERT INTO runs (
                    started_at, profile, files, file_hashes, counts,
                    disclosures, verified_crosschecks, findings, report_paths,
                    file_paths, rerun_of, mode, coverage, mapping_coverage,
                    mapping_suggestions, review_counts
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            disclosures=json.loads(row["disclosures"]),
            verified_crosschecks=row["verified_crosschecks"],
            report_paths=json.loads(row["report_paths"]),
            file_paths=json.loads(row["file_paths"]),
            rerun_of=row["rerun_of"],
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
        )

    def list_runs(self, limit: int = 50) -> list[RunRecord]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM runs ORDER BY id DESC LIMIT ?", (limit,)
            ).fetchall()
        return [self._record_from_row(row, with_findings=False) for row in rows]

    def get_run(self, run_id: int) -> RunRecord:
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM runs WHERE id = ?", (run_id,)).fetchone()
        if row is None:
            raise KeyError(f"no QC run with id {run_id}")
        record = self._record_from_row(row, with_findings=True)
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
