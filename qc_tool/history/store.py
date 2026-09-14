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
import time
import zipfile
import zlib
from collections import defaultdict
from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from pydantic import ValidationError

from qc_tool.config.profile import (
    DeliverableProfile,
    ResolvedOutputPolicy,
    canonical_profile_json,
    profile_sha256,
)
from qc_tool.config.resolved_input import ResolvedInputConfigurationV1
from qc_tool.coverage import CoverageItem, FindingOutputMode, MappingCoverage, QCRunMode
from qc_tool.crosscheck.trace import MappingSuggestion
from qc_tool.engine import QCRunResult
from qc_tool.excel.align import AlignmentTrustPayload, decode_alignment_trust_payload
from qc_tool.findings import (
    Finding,
    NumericCounterfactualBasis,
    SeriesAnchor,
    SeriesAnchorV1,
    SeriesAnchorV2,
    Severity,
)
from qc_tool.findings_store import (
    BLOCK_FINDINGS,
    BlockInfo,
    FindingSequence,
    FindingsStoreError,
    encode_block,
    finding_by_id,
    finding_ordinal,
    finding_payload,
    public_payload,
)
from qc_tool.focus.model import (
    EMPTY_SIDECAR_JSON,
    FocusTargetSeed,
    FocusTargetSidecar,
    decode_focus_targets,
    encode_focus_targets,
)
from qc_tool.focus.targets import (
    focus_seeds_for_finding,
)
from qc_tool.history.longitudinal import (
    DecisionOrigin,
    DossierEntry,
    DossierResult,
    DossierStatus,
    RecurrenceEligibility,
    canonical_anchor_digest,
    contract_scope_for_profile,
    decision_anchor,
)
from qc_tool.history.review_state import (
    FINDING_EVIDENCE_VERSION,
    AnnotationLineage,
    AnnotationLineageOutcome,
    AnnotationLineageRelation,
    CarryForwardCandidate,
    PopulationCarryForwardCandidate,
    RunFinalizedError,
    RunSignoff,
    finding_evidence_digest,
)
from qc_tool.package import PackageManifest
from qc_tool.review_series import (
    anchor_matches_finding,
    canonical_series_aggregate_digest,
    canonical_series_anchor_digest,
)
from qc_tool.review_stream import (
    GroupPriorityAggregate,
    GroupSummary,
    RecordSummaryAccumulators,
    counts_from_summaries,
    decode_view_summaries,
    encode_view_summaries,
)
from qc_tool.scope import ComparisonScope
from qc_tool.security import private_directory, private_file
from qc_tool.story import ChangeStory
from qc_tool.triage.preview import canonical_aggregate_digest, canonical_basis_digest

try:  # parse-side only: every encode stays stdlib json (digest-pinned)
    from orjson import loads as _json_loads
except ImportError:  # pragma: no cover - orjson ships in the project env
    _json_loads = json.loads

logger = logging.getLogger(__name__)

#: Largest sidecar payload the UI process will decode into seed models.
#: 32 MB of JSON decodes to roughly 300 MB of Python objects — a bounded,
#: sub-second cost. A monster run's sidecar (hundreds of MB) degrades to
#: "focus unavailable" instead of a multi-GB decode inside the UI server.
FOCUS_SIDECAR_DECODE_CAP = 32 * 1024 * 1024

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
    alignment_trust TEXT NOT NULL DEFAULT 'null',
    mapping_coverage TEXT NOT NULL DEFAULT 'null',
    mapping_suggestions TEXT NOT NULL DEFAULT '[]',
    review_counts TEXT NOT NULL DEFAULT '{}',
    pattern_review_counts TEXT NOT NULL DEFAULT '{}',
    story_counts TEXT NOT NULL DEFAULT '{}',
    comparison_scope TEXT NOT NULL DEFAULT '{}',
    package_manifest TEXT NOT NULL DEFAULT 'null',
    focus_targets TEXT NOT NULL DEFAULT '{}'
    ,profile_snapshot TEXT NOT NULL DEFAULT 'null'
    ,profile_sha256 TEXT NOT NULL DEFAULT ''
    ,formula_engines TEXT NOT NULL DEFAULT '{}'
    ,values_engines TEXT NOT NULL DEFAULT '{}'
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
CREATE TABLE IF NOT EXISTS counterfactual_bases (
    run_id INTEGER NOT NULL,
    finding_id TEXT NOT NULL,
    version INTEGER NOT NULL,
    payload TEXT NOT NULL,
    digest TEXT NOT NULL,
    PRIMARY KEY (run_id, finding_id)
);
CREATE TABLE IF NOT EXISTS run_finding_blocks (
    run_id INTEGER NOT NULL,
    block_ordinal INTEGER NOT NULL,
    first_ordinal INTEGER NOT NULL,
    count INTEGER NOT NULL,
    blob BLOB NOT NULL,
    PRIMARY KEY (run_id, block_ordinal)
);
CREATE TABLE IF NOT EXISTS run_focus_blocks (
    run_id INTEGER NOT NULL,
    block_ordinal INTEGER NOT NULL,
    first_ordinal INTEGER NOT NULL,
    count INTEGER NOT NULL,
    blob BLOB NOT NULL,
    PRIMARY KEY (run_id, block_ordinal)
);
CREATE TABLE IF NOT EXISTS run_view_summaries (
    run_id INTEGER PRIMARY KEY,
    blob BLOB NOT NULL
);
CREATE TABLE IF NOT EXISTS series_anchor_sidecars (
    run_id INTEGER NOT NULL,
    finding_id TEXT NOT NULL,
    version INTEGER NOT NULL,
    payload TEXT NOT NULL,
    digest TEXT NOT NULL,
    PRIMARY KEY (run_id, finding_id)
);
CREATE TABLE IF NOT EXISTS annotation_lineage (
    run_id INTEGER NOT NULL,
    finding_id TEXT NOT NULL,
    source_run_id INTEGER NOT NULL,
    source_finding_id TEXT NOT NULL,
    relation TEXT NOT NULL,
    outcome TEXT NOT NULL,
    evidence_version INTEGER NOT NULL,
    source_digest TEXT NOT NULL,
    applied_at TEXT NOT NULL,
    PRIMARY KEY (run_id, finding_id, source_run_id, source_finding_id)
);
CREATE TABLE IF NOT EXISTS review_sessions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id INTEGER NOT NULL,
    started_at TEXT NOT NULL,
    ended_at TEXT,
    active_seconds REAL
);
CREATE TABLE IF NOT EXISTS decision_occurrences (
    run_id INTEGER NOT NULL,
    finding_id TEXT NOT NULL,
    contract_scope TEXT NOT NULL,
    anchor_version INTEGER NOT NULL,
    anchor_digest TEXT NOT NULL,
    evidence_version INTEGER NOT NULL,
    evidence_digest TEXT NOT NULL,
    PRIMARY KEY (run_id, finding_id)
);
CREATE INDEX IF NOT EXISTS decision_occurrences_anchor_idx
ON decision_occurrences(contract_scope, anchor_digest, run_id);
CREATE TABLE IF NOT EXISTS decision_origins (
    run_id INTEGER NOT NULL,
    finding_id TEXT NOT NULL,
    origin TEXT NOT NULL,
    decided_at TEXT NOT NULL,
    PRIMARY KEY (run_id, finding_id)
);
CREATE TABLE IF NOT EXISTS contract_promotions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    run_id INTEGER NOT NULL,
    finding_id TEXT NOT NULL,
    kind TEXT NOT NULL,
    profile_sha256 TEXT NOT NULL,
    promoted_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS contract_promotions_finding_idx
ON contract_promotions(run_id, finding_id, id);
"""

#: Columns added after the first release; applied to legacy databases.
_MIGRATIONS = {
    "file_paths": "ALTER TABLE runs ADD COLUMN file_paths TEXT NOT NULL DEFAULT '{}'",
    "rerun_of": "ALTER TABLE runs ADD COLUMN rerun_of INTEGER",
    "mode": (
        "ALTER TABLE runs ADD COLUMN mode TEXT NOT NULL DEFAULT 'cycle_comparison'"
    ),
    "coverage": "ALTER TABLE runs ADD COLUMN coverage TEXT NOT NULL DEFAULT '[]'",
    "alignment_trust": "ALTER TABLE runs ADD COLUMN alignment_trust TEXT NOT NULL DEFAULT 'null'",
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
    "package_manifest": (
        "ALTER TABLE runs ADD COLUMN package_manifest TEXT NOT NULL DEFAULT 'null'"
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
    "counterfactual_digest": (
        "ALTER TABLE runs ADD COLUMN counterfactual_digest TEXT NOT NULL DEFAULT ''"
    ),
    "series_anchor_digest": (
        "ALTER TABLE runs ADD COLUMN series_anchor_digest TEXT NOT NULL DEFAULT ''"
    ),
    #: NULL means "not yet measured" (legacy row); backfilled lazily.
    "storage_bytes": "ALTER TABLE runs ADD COLUMN storage_bytes INTEGER",
    #: Resolved formula-engine/adapter-fingerprint per excel role; "" for
    #: legacy runs recorded before this disclosure existed.
    "formula_engines": (
        "ALTER TABLE runs ADD COLUMN formula_engines TEXT NOT NULL DEFAULT '{}'"
    ),
    "values_engines": (
        "ALTER TABLE runs ADD COLUMN values_engines TEXT NOT NULL DEFAULT '{}'"
    ),
    #: Run-level finding-output contract (plan-20260910); "profile" and
    #: "null" (no resolved policy recorded) for any run committed before
    #: this pair of columns existed -- exactly today's behavior.
    "requested_output_mode": (
        "ALTER TABLE runs ADD COLUMN requested_output_mode TEXT "
        "NOT NULL DEFAULT 'profile'"
    ),
    "resolved_output_policy": (
        "ALTER TABLE runs ADD COLUMN resolved_output_policy TEXT "
        "NOT NULL DEFAULT 'null'"
    ),
    #: Run-level scope-affecting overrides (plan-20260913); defaults exactly
    #: match `run_qc()`'s own pre-existing defaults, so a legacy run reads
    #: as "nothing was overridden" -- its own historical truth.
    "allow_dependency_indexing": (
        "ALTER TABLE runs ADD COLUMN allow_dependency_indexing INTEGER "
        "NOT NULL DEFAULT 0"
    ),
    "acceptance_absolute": (
        "ALTER TABLE runs ADD COLUMN acceptance_absolute REAL NOT NULL DEFAULT 0.0"
    ),
    "acceptance_relative": (
        "ALTER TABLE runs ADD COLUMN acceptance_relative REAL NOT NULL DEFAULT 0.0"
    ),
    #: Exact primitive `ResolvedInputConfigurationV1` payload and its
    #: canonical digest (plan-20260913); `null`/"" for any run committed
    #: before this pair of columns existed -- the explicit legacy-absent
    #: default the compatibility service treats as "no logical contract".
    "resolved_input_configuration": (
        "ALTER TABLE runs ADD COLUMN resolved_input_configuration TEXT "
        "NOT NULL DEFAULT 'null'"
    ),
    "resolved_input_digest": (
        "ALTER TABLE runs ADD COLUMN resolved_input_digest TEXT NOT NULL DEFAULT ''"
    ),
}


def _migrate_annotation_lineage_v2(conn: sqlite3.Connection) -> None:
    """Rebuild a legacy strictly-1:1 `annotation_lineage` into the versioned
    many-to-one shape (A4). SQLite cannot alter a table's primary key in
    place, so this copies every legacy row forward, then replaces the table.

    Every legacy row represents a clean, already-applied 1:1 carry, so each
    becomes one `relation="identity"`, `outcome="inherited"` row under the
    new shape -- a faithful, lossless reinterpretation, not a guess.
    """
    columns = {row[1] for row in conn.execute("PRAGMA table_info(annotation_lineage)")}
    if "relation" in columns:
        return  # already the new shape (new database, or already migrated)
    legacy_rows = conn.execute(
        "SELECT run_id, finding_id, source_run_id, source_finding_id,"
        " evidence_version, evidence_digest, applied_at FROM annotation_lineage"
    ).fetchall()
    conn.execute("DROP TABLE annotation_lineage")
    conn.execute(
        """
        CREATE TABLE annotation_lineage (
            run_id INTEGER NOT NULL,
            finding_id TEXT NOT NULL,
            source_run_id INTEGER NOT NULL,
            source_finding_id TEXT NOT NULL,
            relation TEXT NOT NULL,
            outcome TEXT NOT NULL,
            evidence_version INTEGER NOT NULL,
            source_digest TEXT NOT NULL,
            applied_at TEXT NOT NULL,
            PRIMARY KEY (run_id, finding_id, source_run_id, source_finding_id)
        )
        """
    )
    conn.executemany(
        """
        INSERT INTO annotation_lineage (
            run_id, finding_id, source_run_id, source_finding_id,
            relation, outcome, evidence_version, source_digest, applied_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        [
            (
                row["run_id"],
                row["finding_id"],
                row["source_run_id"],
                row["source_finding_id"],
                AnnotationLineageRelation.IDENTITY.value,
                AnnotationLineageOutcome.INHERITED.value,
                row["evidence_version"],
                row["evidence_digest"],
                row["applied_at"],
            )
            for row in legacy_rows
        ],
    )

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
    package_manifest: PackageManifest | None = None
    file_paths: dict[str, str] = field(default_factory=dict)  # role -> stored path
    rerun_of: int | None = None
    #: Retired from the default history view; the record itself is retained.
    archived: bool = False
    coverage: list[CoverageItem] = field(default_factory=list)
    mapping_coverage: MappingCoverage | None = None
    mapping_suggestions: list[MappingSuggestion] = field(default_factory=list)
    findings: Sequence[Finding] = field(default_factory=list)
    #: Private desktop-focus locators; never exported, reported, or attested.
    focus_targets: FocusTargetSidecar = field(default_factory=FocusTargetSidecar)
    #: Undecoded sidecar payload; ``focus_sidecar()`` decodes it on first use.
    focus_targets_raw: str | None = None
    #: Block-stored seeds (runs recorded after 2026-08-09); preferred source.
    focus_blocks: "FocusBlockSource | None" = None
    #: Set when the legacy sidecar exceeded the decode cap; drives the UI note.
    focus_decode_capped: bool = False
    profile_snapshot: DeliverableProfile | None = None
    profile_sha256: str = ""
    signoff: RunSignoff | None = None
    counterfactual_digest: str = ""
    #: Aggregate digest of the private logical-series anchor sidecar; "" is a
    #: legacy run recorded before the sidecar existed.
    series_anchor_digest: str = ""
    alignment_trust: AlignmentTrustPayload | None = None
    #: Bytes this run occupies (database row + report files); None until measured.
    storage_bytes: int | None = None
    #: Resolved formula-engine/adapter-fingerprint string per excel role
    #: (e.g. "native-biff12:1.2.3"); "{}" for legacy runs or roles where
    #: formula enrichment never ran.
    formula_engines: dict[str, str] = field(default_factory=dict)
    #: Resolved cached-values decoder per excel role; ``{}`` for legacy rows.
    values_engines: dict[str, str] = field(default_factory=dict)
    #: Run-level finding-output contract request (plan-20260910); "profile"
    #: for any run recorded before this contract existed -- exactly today's
    #: legacy behavior.
    requested_output_mode: FindingOutputMode = FindingOutputMode.PROFILE
    #: The effective population policy this run actually used, and why;
    #: `None` for any run recorded before this contract existed.
    resolved_output_policy: ResolvedOutputPolicy | None = None
    #: Run-level scope-affecting overrides (plan-20260913); `False`/`0.0`
    #: for any run recorded before these columns existed -- exactly the
    #: pre-existing `run_qc()` defaults, so a legacy row reads as its own
    #: historical truth.
    allow_dependency_indexing: bool = False
    acceptance_absolute: float = 0.0
    acceptance_relative: float = 0.0
    #: Exact per-run resolved logical configuration and its canonical
    #: digest (plan-20260913); `None`/"" is the explicit legacy-absent
    #: default for any run with no saved `input_contract`.
    resolved_input_configuration: ResolvedInputConfigurationV1 | None = None
    resolved_input_digest: str = ""

    def focus_sidecar(self) -> FocusTargetSidecar:
        """Decode the private sidecar on first focus use, never at row read.

        A monster run's sidecar reaches hundreds of MB of JSON (~1.7M seed
        models decoded); eager decoding at every record fetch is what let a
        948k-finding run OOM the UI server on run completion. Payloads past
        ``FOCUS_SIDECAR_DECODE_CAP`` degrade to an empty sidecar (focus
        unavailable for that run) instead of a multi-GB decode.
        """
        if self.focus_targets_raw is not None:
            raw = self.focus_targets_raw
            self.focus_targets_raw = None
            if len(raw) > FOCUS_SIDECAR_DECODE_CAP:
                self.focus_decode_capped = True
                logger.warning(
                    "focus sidecar for run %s exceeds the decode cap "
                    "(%d bytes); desktop focus is unavailable for this run",
                    self.run_id,
                    len(raw),
                )
            else:
                self.focus_targets = decode_focus_targets(raw)
        return self.focus_targets

    def focus_seeds(self, finding_id: str) -> tuple[FocusTargetSeed, ...]:
        """One finding's seeds: block storage first, legacy sidecar fallback."""
        if self.focus_blocks is not None:
            return self.focus_blocks.seeds(finding_id)
        return self.focus_sidecar().seeds(finding_id)

    def focus_degraded(self) -> bool:
        """True when focus is unavailable only because of the legacy cap."""
        if self.focus_blocks is not None:
            return False
        if self.focus_targets_raw is not None:
            return len(self.focus_targets_raw) > FOCUS_SIDECAR_DECODE_CAP
        return self.focus_decode_capped

def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()

def _decode_series_anchor(payload: object) -> SeriesAnchor:
    """Decode a stored anchor by its own version so V1 rows keep working."""
    if not isinstance(payload, dict):
        raise KeyError("series anchor payload is not an object")
    models: dict[int, type[SeriesAnchorV1] | type[SeriesAnchorV2]] = {
        1: SeriesAnchorV1,
        2: SeriesAnchorV2,
    }
    model = models.get(int(payload.get("version", 0)))
    if model is None:
        raise KeyError("unknown series anchor version")
    return model.model_validate(payload)

def _reject_series_anchors(code: str) -> dict[str, SeriesAnchor]:
    """Log a fixed rejection code and fall back to the canonical queue."""
    logger.warning("series anchor sidecar rejected (%s)", code)
    return {}

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

def _report_directory_bytes(report_paths: dict[str, str], root: Path) -> int:
    """Total size of the managed run directories referenced by ``report_paths``."""
    directories: set[Path] = set()
    for raw in report_paths.values():
        try:
            resolved = Path(raw).resolve()
            resolved.relative_to(root)
        except (OSError, ValueError):
            continue
        directories.add(resolved.parent)
    total = 0
    for directory in directories:
        try:
            for item in directory.rglob("*"):
                if item.is_file():
                    total += item.stat().st_size
        except OSError:
            continue
    return total

class _RunBlockSource:
    """``BlockSource`` over one run's compressed finding blocks in SQLite."""

    def __init__(self, db_path: Path, run_id: int, infos: list[BlockInfo]) -> None:
        self._db_path = db_path
        self._run_id = run_id
        self._infos = infos

    def block_infos(self) -> Sequence[BlockInfo]:
        return self._infos

    def read_block(self, index: int) -> bytes:
        conn = sqlite3.connect(self._db_path, timeout=30.0)
        try:
            row = conn.execute(
                "SELECT blob FROM run_finding_blocks"
                " WHERE run_id = ? AND block_ordinal = ?",
                (self._run_id, index),
            ).fetchone()
        finally:
            conn.close()
        if row is None:
            raise FindingsStoreError(
                f"run {self._run_id} finding block {index} is missing"
            )
        return bytes(row[0])

class FocusBlockSource:
    """Per-finding focus seeds from block storage; one blob decode per hit.

    Each blob is zlib-compressed JSONL, one line per finding in production
    order, so a finding's seeds are the ``ordinal - first_ordinal``-th line.
    Anything unreadable degrades to no seeds, matching the sidecar decoder.
    """

    def __init__(self, db_path: Path, run_id: int) -> None:
        self._db_path = db_path
        self._run_id = run_id
        self._cached: tuple[int, int, list[str]] | None = None

    def seeds(self, finding_id: str) -> tuple[FocusTargetSeed, ...]:
        ordinal = finding_ordinal(finding_id)
        if ordinal < 1:
            return ()
        position = ordinal - 1  # F%04d ids are 1-based; lines are 0-based
        lines = self._block_lines(position)
        if lines is None:
            return ()
        first, _count, payload = lines
        index = position - first
        if index < 0 or index >= len(payload):
            return ()
        try:
            entries = _json_loads(payload[index])
            return tuple(
                FocusTargetSeed.model_validate(entry) for entry in entries
            )
        except (json.JSONDecodeError, ValueError, ValidationError):
            logger.warning("focus-block-unreadable")
            return ()

    def _block_lines(self, position: int) -> tuple[int, int, list[str]] | None:
        if self._cached is not None:
            first, count, payload = self._cached
            if first <= position < first + count:
                return self._cached
        conn = sqlite3.connect(self._db_path, timeout=30.0)
        try:
            row = conn.execute(
                "SELECT first_ordinal, count, blob FROM run_focus_blocks"
                " WHERE run_id = ? AND first_ordinal <= ?"
                " ORDER BY first_ordinal DESC LIMIT 1",
                (self._run_id, position),
            ).fetchone()
        finally:
            conn.close()
        if row is None or position >= row[0] + row[1]:
            return None
        try:
            payload = zlib.decompress(bytes(row[2])).decode("utf-8").split("\n")
        except (zlib.error, UnicodeDecodeError):
            logger.warning("focus-block-unreadable")
            return None
        self._cached = (row[0], row[1], payload)
        return self._cached

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
            _migrate_annotation_lineage_v2(conn)
            private_file(db_path)

    @property
    def work_dir(self) -> Path:
        """The managed data directory this history is stored under
        (plan-20260913, Step 5) -- the same directory `RunQueueManager`/
        `get_exclusive_slot` key on for this session's shared exclusive slot.
        """
        return self._db_path.parent

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
        focus_targets: dict[str, tuple[FocusTargetSeed, ...]] | str | None = None,
        profile_snapshot: DeliverableProfile | None = None,
        on_subphase: Callable[[str, float], None] | None = None,
    ) -> int:
        """Record one run. ``on_subphase(name, elapsed_seconds)`` -- when
        given -- fires once per fixed-code subphase (``main_pass``,
        ``story_classify_and_replay``, ``sqlite_write``,
        ``storage_measurement``, ``total``); a diagnostic hook only, never
        persisted and never required for correctness. Aggregate timings
        only -- no path, cell value, or formula text ever reaches it.
        """
        def mark(name: str, start: float) -> None:
            if on_subphase is not None:
                on_subphase(name, time.perf_counter() - start)

        total_start = time.perf_counter()
        started_at = dt.datetime.now(dt.UTC).isoformat(timespec="seconds")
        counts = {sev.value: count for sev, count in result.counts.items()}

        # Record-time passes may use the trusted constructor: the findings
        # container was written by THIS process moments ago, already
        # validated. Shared read surfaces keep the validating iterator.
        def record_pass() -> Iterable[Finding]:
            if isinstance(result.findings, FindingSequence):
                return result.findings.iter_trusted()
            return iter(result.findings)

        # One streamed pass: review/pattern summaries, story pass 1, private
        # sidecar bindings, decision occurrences, focus seeds, and the
        # block-compressed public findings payloads. Block payloads strip
        # the private fields so stored rows match the legacy
        # ``runs.findings`` projection exactly.
        summary_accumulators = RecordSummaryAccumulators()
        contract_scope = contract_scope_for_profile(
            (
                profile_snapshot.name
                if profile_snapshot is not None
                else result.profile_name
            ),
            (
                profile_snapshot.contract_id
                if profile_snapshot is not None
                else None
            ),
        )
        bases: dict[str, NumericCounterfactualBasis] = {}
        anchor_bindings: dict[str, tuple[str, str, SeriesAnchor]] = {}
        occurrence_seeds: list[tuple[str, int, str, str]] = []
        blocks: list[tuple[int, int, bytes]] = []  # (first_ordinal, count, blob)
        batch: list[dict[str, object]] = []
        first_ordinal = 0
        # Seeds generate inline with the same block boundaries as findings,
        # so a focus lookup decodes one small blob instead of one giant row.
        generate_focus = focus_targets is None
        focus_blocks: list[tuple[int, int, bytes]] = []
        focus_batch: list[str] = []
        focus_seeded = False

        def flush_block() -> None:
            nonlocal first_ordinal
            if not batch:
                return
            if generate_focus:
                focus_blocks.append(
                    (
                        first_ordinal,
                        len(batch),
                        zlib.compress("\n".join(focus_batch).encode("utf-8"), 6),
                    )
                )
                focus_batch.clear()
            blocks.append((first_ordinal, len(batch), encode_block(batch)))
            first_ordinal += len(batch)
            batch.clear()

        main_pass_start = time.perf_counter()
        for finding in record_pass():
            summary_accumulators.observe(finding)
            basis = finding.counterfactual_basis
            if basis is not None:
                bases[finding.finding_id] = basis
            anchor = finding.series_anchor
            if anchor is not None and anchor_matches_finding(finding, anchor):
                anchor_bindings[finding.finding_id] = (
                    finding.artifact_member,
                    finding.location or "",
                    anchor,
                )
            decision = decision_anchor(finding)
            occurrence_seeds.append(
                (
                    finding.finding_id,
                    decision.version,
                    canonical_anchor_digest(decision),
                    finding_evidence_digest(finding),
                )
            )
            if generate_focus:
                try:
                    seeds = focus_seeds_for_finding(
                        finding, mode=result.mode, file_hashes=file_hashes
                    )
                    focus_batch.append(
                        json.dumps(
                            [seed.model_dump(mode="json") for seed in seeds]
                        )
                    )
                    focus_seeded = focus_seeded or bool(seeds)
                except Exception:
                    # Fixed code only: never log locators, paths, or values.
                    # A focus failure never fails the run; drop all seeds.
                    logger.warning("focus-target-generation-failed")
                    generate_focus = False
                    focus_blocks.clear()
                    focus_batch.clear()
            batch.append(public_payload(finding_payload(finding)))
            if len(batch) >= BLOCK_FINDINGS:
                flush_block()
        flush_block()
        mark("main_pass", main_pass_start)

        grouped = counts_from_summaries(summary_accumulators.review.finish())
        grouped_counts = {
            severity.value: count
            for severity, count in grouped.review_items.items()
        }
        pattern_summaries, priority_aggregates = (
            summary_accumulators.pattern.finish()
        )
        patterned = counts_from_summaries(pattern_summaries)
        pattern_counts = {
            severity.value: count
            for severity, count in patterned.review_items.items()
        }
        story_counts: dict[str, dict[str, int]] = {}
        story_start = time.perf_counter()
        stories = summary_accumulators.stories.finish(
            record_pass(), record_pass()
        )
        mark("story_classify_and_replay", story_start)
        for story in stories:
            bucket = story_counts.setdefault(
                story.kind.value, {"stories": 0, "members": 0}
            )
            bucket["stories"] += 1
            bucket["members"] += story.member_count
        view_summaries_blob = encode_view_summaries(
            pattern_summaries, priority_aggregates, stories
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
        # New runs bind even an empty eligible population. Legacy rows retain
        # the migration default "", which lets the UI distinguish the states.
        # Each basis/anchor's model_dump(mode="json") is computed exactly
        # once here and reused for both its own sidecar row and the
        # aggregate digest -- previously recomputed 2-3x per item (once via
        # canonical_aggregate_digest/canonical_series_aggregate_digest here,
        # again for the row payload below, again inside canonical_basis_
        # digest/canonical_series_anchor_digest), the dominant real cost of
        # this phase at large finding counts.
        base_rows: list[tuple[object, ...]] = []
        aggregate_payloads: dict[str, object] = {}
        for finding_id, basis in sorted(bases.items()):
            payload = basis.model_dump(mode="json")
            payload_json = json.dumps(payload, sort_keys=True, separators=(",", ":"))
            digest = hashlib.sha256(payload_json.encode("utf-8")).hexdigest()
            base_rows.append((finding_id, basis.version, payload_json, digest))
            aggregate_payloads[finding_id] = payload
        aggregate_digest = hashlib.sha256(
            json.dumps(
                aggregate_payloads, sort_keys=True, separators=(",", ":")
            ).encode("utf-8")
        ).hexdigest()

        anchor_rows: list[tuple[object, ...]] = []
        series_payloads: dict[str, object] = {}
        for finding_id, (member, location, anchor) in sorted(anchor_bindings.items()):
            anchor_payload = anchor.model_dump(mode="json")
            anchor_payload_json = json.dumps(
                anchor_payload, sort_keys=True, separators=(",", ":")
            )
            binding_payload = {
                "finding_id": finding_id,
                "artifact_member": member,
                "location": location,
                "anchor": anchor_payload,
            }
            row_digest = hashlib.sha256(
                json.dumps(
                    binding_payload, sort_keys=True, separators=(",", ":")
                ).encode("utf-8")
            ).hexdigest()
            anchor_rows.append(
                (finding_id, anchor.version, anchor_payload_json, row_digest)
            )
            series_payloads[finding_id] = binding_payload
        series_digest = hashlib.sha256(
            json.dumps(
                series_payloads, sort_keys=True, separators=(",", ":")
            ).encode("utf-8")
        ).hexdigest()

        sqlite_start = time.perf_counter()
        with self._connect() as conn:
            cursor = conn.execute(
                """
                INSERT INTO runs (
                    started_at, profile, files, file_hashes, counts,
                    disclosures, verified_crosschecks, findings, report_paths,
                    file_paths, rerun_of, mode, coverage, alignment_trust, mapping_coverage,
                    mapping_suggestions, review_counts, pattern_review_counts,
                    story_counts, comparison_scope, package_manifest, focus_targets,
                    profile_snapshot, profile_sha256, counterfactual_digest,
                    series_anchor_digest, formula_engines, values_engines,
                    requested_output_mode, resolved_output_policy,
                    allow_dependency_indexing, acceptance_absolute, acceptance_relative,
                    resolved_input_configuration, resolved_input_digest
                ) VALUES (
                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                    ?, ?, ?, ?, ?
                )
                """,
                (
                    started_at,
                    result.profile_name,
                    json.dumps(result.files),
                    json.dumps(file_hashes),
                    json.dumps(counts),
                    json.dumps(result.disclosures),
                    result.verified_crosschecks,
                    "[]",  # findings live in run_finding_blocks for new runs
                    json.dumps(report_paths),
                    json.dumps(file_paths or {}),
                    rerun_of,
                    result.mode.value,
                    json.dumps([item.model_dump(mode="json") for item in result.coverage]),
                    json.dumps(
                        result.alignment_trust.model_dump(mode="json")
                        if result.alignment_trust is not None
                        else None
                    ),
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
                    json.dumps(
                        result.package_manifest.model_dump(mode="json")
                        if result.package_manifest is not None
                        else None
                    ),
                    (
                        focus_targets
                        if isinstance(focus_targets, str)
                        else encode_focus_targets(focus_targets)
                        if focus_targets
                        else EMPTY_SIDECAR_JSON
                    ),
                    profile_payload,
                    profile_digest,
                    aggregate_digest,
                    series_digest,
                    json.dumps(result.formula_engines),
                    json.dumps(result.values_engines),
                    result.requested_output_mode.value,
                    json.dumps(
                        result.resolved_output_policy.model_dump(mode="json")
                        if result.resolved_output_policy is not None
                        else None
                    ),
                    result.allow_dependency_indexing,
                    result.acceptance_absolute,
                    result.acceptance_relative,
                    json.dumps(
                        result.resolved_input_configuration.model_dump(mode="json")
                        if result.resolved_input_configuration is not None
                        else None
                    ),
                    result.resolved_input_digest,
                ),
            )
            run_id = cursor.lastrowid
            if run_id is not None and blocks:
                conn.executemany(
                    """
                    INSERT INTO run_finding_blocks (
                        run_id, block_ordinal, first_ordinal, count, blob
                    ) VALUES (?, ?, ?, ?, ?)
                    """,
                    [
                        (run_id, ordinal, first, count, blob)
                        for ordinal, (first, count, blob) in enumerate(blocks)
                    ],
                )
            if run_id is not None and focus_blocks and focus_seeded:
                conn.executemany(
                    """
                    INSERT INTO run_focus_blocks (
                        run_id, block_ordinal, first_ordinal, count, blob
                    ) VALUES (?, ?, ?, ?, ?)
                    """,
                    [
                        (run_id, ordinal, first, count, blob)
                        for ordinal, (first, count, blob) in enumerate(focus_blocks)
                    ],
                )
            if run_id is not None:
                conn.execute(
                    "INSERT INTO run_view_summaries (run_id, blob) VALUES (?, ?)",
                    (run_id, view_summaries_blob),
                )
            # Insert decision occurrences for longitudinal queries in the
            # same transaction so the run and its derived occurrences are
            # atomically visible.
            if run_id is not None and occurrence_seeds:
                conn.executemany(
                    """
                    INSERT INTO decision_occurrences (
                        run_id, finding_id, contract_scope, anchor_version,
                        anchor_digest, evidence_version, evidence_digest
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    [
                        (
                            run_id,
                            finding_id,
                            contract_scope,
                            anchor_version,
                            anchor_digest,
                            FINDING_EVIDENCE_VERSION,
                            evidence_digest,
                        )
                        for finding_id, anchor_version, anchor_digest, evidence_digest
                        in occurrence_seeds
                    ],
                )
            # Insert private counterfactual bases sidecar rows
            if run_id is not None and bases:
                conn.executemany(
                    """
                    INSERT INTO counterfactual_bases (
                        run_id, finding_id, version, payload, digest
                    )
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    [(run_id, *row) for row in base_rows],
                )
            # Insert private logical-series anchor sidecar rows
            if run_id is not None and anchor_bindings:
                conn.executemany(
                    """
                    INSERT INTO series_anchor_sidecars (
                        run_id, finding_id, version, payload, digest
                    )
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    [(run_id, *row) for row in anchor_rows],
                )
        mark("sqlite_write", sqlite_start)
        if run_id is None:  # pragma: no cover - sqlite always returns a rowid
            raise RuntimeError("sqlite did not return a run id")
        storage_start = time.perf_counter()
        with self._connect() as conn:
            conn.execute(
                "UPDATE runs SET storage_bytes = ? WHERE id = ?",
                (self._measure_run_storage(conn, run_id), run_id),
            )
        mark("storage_measurement", storage_start)
        mark("total", total_start)
        logger.info("recorded QC run %d (%s)", run_id, result.profile_name)
        return run_id

    #: Tables whose rows belong to one run and count toward its footprint.
    _RUN_SIDE_TABLES = (
        "run_finding_blocks",
        "run_focus_blocks",
        "run_view_summaries",
        "decision_occurrences",
        "counterfactual_bases",
        "series_anchor_sidecars",
        "annotations",
    )

    @staticmethod
    def _stored_value_bytes(value: object) -> int:
        if value is None:
            return 1
        if isinstance(value, str):
            return len(value.encode("utf-8"))
        if isinstance(value, (bytes, memoryview)):
            return len(value)
        return 8

    def _measure_run_storage(self, conn: sqlite3.Connection, run_id: int) -> int:
        """Bytes held by one run: its rows plus its managed report directory."""
        total = 0
        row = conn.execute("SELECT * FROM runs WHERE id = ?", (run_id,)).fetchone()
        if row is None:
            return 0
        total += sum(self._stored_value_bytes(value) for value in tuple(row))
        for table in self._RUN_SIDE_TABLES:
            for side_row in conn.execute(
                f"SELECT * FROM {table} WHERE run_id = ?", (run_id,)
            ):
                total += sum(self._stored_value_bytes(value) for value in tuple(side_row))
        report_paths = json.loads(row["report_paths"])
        managed_root = (self._db_path.parent / "runs").resolve()
        total += _report_directory_bytes(report_paths, managed_root)
        return total

    def backfill_storage_bytes(self, limit: int = 200) -> int:
        """Measure legacy rows recorded before the size column existed."""
        with self._connect() as conn:
            ids = [
                r[0]
                for r in conn.execute(
                    "SELECT id FROM runs WHERE storage_bytes IS NULL"
                    " ORDER BY id DESC LIMIT ?",
                    (limit,),
                )
            ]
            for run_id in ids:
                conn.execute(
                    "UPDATE runs SET storage_bytes = ? WHERE id = ?",
                    (self._measure_run_storage(conn, run_id), run_id),
                )
        return len(ids)

    def storage_summary(self) -> tuple[int, int, int]:
        """(total_bytes, archived_bytes, unmeasured_rows) across all runs."""
        with self._connect() as conn:
            total, archived, unmeasured = conn.execute(
                "SELECT COALESCE(SUM(storage_bytes), 0),"
                " COALESCE(SUM(CASE WHEN archived = 1 THEN storage_bytes END), 0),"
                " SUM(CASE WHEN storage_bytes IS NULL THEN 1 ELSE 0 END)"
                " FROM runs"
            ).fetchone()
        return int(total), int(archived), int(unmeasured or 0)

    def _block_infos(self, conn: sqlite3.Connection, run_id: int) -> list[BlockInfo]:
        rows = conn.execute(
            "SELECT block_ordinal, count, LENGTH(blob) AS length"
            " FROM run_finding_blocks WHERE run_id = ? ORDER BY block_ordinal",
            (run_id,),
        ).fetchall()
        return [
            BlockInfo(
                offset=row["block_ordinal"],
                length=row["length"],
                count=row["count"],
            )
            for row in rows
        ]

    def _load_findings(
        self,
        conn: sqlite3.Connection,
        row: sqlite3.Row,
        *,
        decorate: Callable[[Finding], None] | None,
    ) -> Sequence[Finding]:
        """Findings for one run: lazy block pages, or the legacy JSON column."""
        infos = self._block_infos(conn, row["id"])
        if infos:
            return FindingSequence(
                _RunBlockSource(self._db_path, row["id"], infos),
                decorate=decorate,
            )
        findings = [
            Finding.model_validate(item) for item in json.loads(row["findings"])
        ]
        if decorate is not None:
            for finding in findings:
                decorate(finding)
        return findings

    def _record_from_row(
        self, row: sqlite3.Row, *, include_focus_raw: bool = True
    ) -> RunRecord:
        """Metadata-only record; ``_load_findings`` attaches findings lazily."""
        alignment_trust_payload = json.loads(row["alignment_trust"])
        package_manifest_payload = json.loads(row["package_manifest"])
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
            package_manifest=(
                PackageManifest.model_validate(package_manifest_payload)
                if package_manifest_payload is not None
                else None
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
            findings=[],
            # A legacy monster sidecar reaches hundreds of MB; listings never
            # use focus, so they skip carrying the raw payload entirely.
            focus_targets_raw=(
                row["focus_targets"] if include_focus_raw else None
            ),
            profile_snapshot=(
                DeliverableProfile.model_validate(json.loads(row["profile_snapshot"]))
                if json.loads(row["profile_snapshot"]) is not None
                else None
            ),
            profile_sha256=row["profile_sha256"],
            counterfactual_digest=row["counterfactual_digest"],
            series_anchor_digest=row["series_anchor_digest"],
            storage_bytes=row["storage_bytes"],
            alignment_trust=decode_alignment_trust_payload(alignment_trust_payload),
            formula_engines=json.loads(row["formula_engines"]),
            values_engines=json.loads(row["values_engines"]),
            requested_output_mode=FindingOutputMode(row["requested_output_mode"]),
            resolved_output_policy=(
                ResolvedOutputPolicy.model_validate(resolved_output_policy_payload)
                if (
                    resolved_output_policy_payload := json.loads(
                        row["resolved_output_policy"]
                    )
                )
                is not None
                else None
            ),
            allow_dependency_indexing=bool(row["allow_dependency_indexing"]),
            acceptance_absolute=row["acceptance_absolute"],
            acceptance_relative=row["acceptance_relative"],
            resolved_input_configuration=(
                ResolvedInputConfigurationV1.model_validate(
                    resolved_input_configuration_payload
                )
                if (
                    resolved_input_configuration_payload := json.loads(
                        row["resolved_input_configuration"]
                    )
                )
                is not None
                else None
            ),
            resolved_input_digest=row["resolved_input_digest"],
        )

    def list_runs(self, limit: int = 50, *, include_archived: bool = True) -> list[RunRecord]:
        clause = "" if include_archived else " WHERE archived = 0"
        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT * FROM runs{clause} ORDER BY id DESC LIMIT ?", (limit,)
            ).fetchall()
        records = [
            self._record_from_row(row, include_focus_raw=False) for row in rows
        ]
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
                f"DELETE FROM run_finding_blocks WHERE run_id IN ({placeholders})",
                ids,
            )
            conn.execute(
                f"DELETE FROM run_focus_blocks WHERE run_id IN ({placeholders})",
                ids,
            )
            conn.execute(
                f"DELETE FROM run_view_summaries WHERE run_id IN ({placeholders})",
                ids,
            )
            conn.execute(
                f"DELETE FROM decision_occurrences WHERE run_id IN ({placeholders})",
                ids,
            )
            conn.execute(
                f"DELETE FROM decision_origins WHERE run_id IN ({placeholders})",
                ids,
            )
            conn.execute(
                f"DELETE FROM contract_promotions WHERE run_id IN ({placeholders})",
                ids,
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
            conn.execute(
                f"DELETE FROM counterfactual_bases WHERE run_id IN ({placeholders})",
                ids,
            )
            conn.execute(
                f"DELETE FROM series_anchor_sidecars WHERE run_id IN ({placeholders})",
                ids,
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
        annotations = self.get_annotations(run_id)

        def decorate(finding: Finding) -> None:
            annotation = annotations.get(finding.finding_id)
            if annotation is None:
                return
            severity, comment = annotation
            finding.analyst_comment = comment
            if severity is not None:
                finding.severity = Severity(severity)
                finding.severity_overridden = True

        with self._connect() as conn:
            row = conn.execute("SELECT * FROM runs WHERE id = ?", (run_id,)).fetchone()
            if row is None:
                raise KeyError(f"no QC run with id {run_id}")
            record = self._record_from_row(row)
            record.findings = self._load_findings(
                conn, row, decorate=decorate if annotations else None
            )
            record.focus_blocks = self._focus_blocks(conn, run_id)
        record.signoff = self.get_signoff(run_id)
        return record

    def get_raw_run(self, run_id: int) -> RunRecord:
        """Rehydrate immutable engine findings without analyst annotations."""
        with self._connect() as conn:
            row = conn.execute("SELECT * FROM runs WHERE id = ?", (run_id,)).fetchone()
            if row is None:
                raise KeyError(f"no QC run with id {run_id}")
            record = self._record_from_row(row)
            record.findings = self._load_findings(conn, row, decorate=None)
            record.focus_blocks = self._focus_blocks(conn, run_id)
        record.signoff = self.get_signoff(run_id)
        return record

    def _focus_blocks(
        self, conn: sqlite3.Connection, run_id: int
    ) -> "FocusBlockSource | None":
        present = conn.execute(
            "SELECT 1 FROM run_focus_blocks WHERE run_id = ? LIMIT 1", (run_id,)
        ).fetchone()
        if present is None:
            return None
        return FocusBlockSource(self._db_path, run_id)

    def get_view_summaries(
        self, run_id: int
    ) -> tuple[
        list[GroupSummary], dict[str, GroupPriorityAggregate], list[ChangeStory]
    ] | None:
        """Stored pattern-group summaries and stories, or ``None`` for legacy runs."""
        with self._connect() as conn:
            row = conn.execute(
                "SELECT blob FROM run_view_summaries WHERE run_id = ?", (run_id,)
            ).fetchone()
        if row is None:
            return None
        return decode_view_summaries(bytes(row[0]))

    def set_report_paths(self, run_id: int, report_paths: dict[str, str]) -> None:
        """Attach on-demand-generated report files to a recorded run."""
        with self._connect() as conn:
            cursor = conn.execute(
                "UPDATE runs SET report_paths = ? WHERE id = ?",
                (json.dumps(report_paths), run_id),
            )
            if cursor.rowcount == 0:
                raise KeyError(f"no QC run with id {run_id}")
            conn.execute(
                "UPDATE runs SET storage_bytes = ? WHERE id = ?",
                (self._measure_run_storage(conn, run_id), run_id),
            )

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
            # Record MANUAL origins for these annotations in the longitudinal
            # table; update decided_at when existing.
            conn.executemany(
                """
                INSERT INTO decision_origins (run_id, finding_id, origin, decided_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(run_id, finding_id) DO UPDATE SET
                    origin = excluded.origin,
                    decided_at = excluded.decided_at
                """,
                [
                    (run_id, finding_id, DecisionOrigin.MANUAL.value, updated_at)
                    for finding_id, _, _ in updates
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

    def record_promotion(
        self,
        run_id: int,
        finding_id: str,
        kind: str,
        profile_sha256: str,
    ) -> None:
        promoted_at = dt.datetime.now(dt.UTC).isoformat(timespec="seconds")
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO contract_promotions (
                    run_id, finding_id, kind, profile_sha256, promoted_at
                ) VALUES (?, ?, ?, ?, ?)
                """,
                (run_id, finding_id, kind, profile_sha256, promoted_at),
            )

    def get_annotation_lineage(self, run_id: int) -> dict[str, tuple[AnnotationLineage, ...]]:
        """Every lineage row for this run, grouped by current `finding_id`.

        A finding may have many rows (a population absorbing many prior
        atomic sources); each value is a tuple, never a single row, so
        callers cannot silently drop sources by iterating `.values()`.
        """
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM annotation_lineage WHERE run_id = ?"
                " ORDER BY finding_id, source_finding_id",
                (run_id,),
            ).fetchall()
        grouped: dict[str, list[AnnotationLineage]] = defaultdict(list)
        for row in rows:
            grouped[row["finding_id"]].append(
                AnnotationLineage(
                    finding_id=row["finding_id"],
                    source_run_id=row["source_run_id"],
                    source_finding_id=row["source_finding_id"],
                    relation=row["relation"],
                    outcome=row["outcome"],
                    evidence_version=row["evidence_version"],
                    source_digest=row["source_digest"],
                    applied_at=row["applied_at"],
                )
            )
        return {finding_id: tuple(lineages) for finding_id, lineages in grouped.items()}

    def get_counterfactual_bases(self, run_id: int) -> dict[str, NumericCounterfactualBasis]:
        with self._connect() as conn:
            run = conn.execute(
                "SELECT counterfactual_digest FROM runs WHERE id = ?", (run_id,)
            ).fetchone()
            if run is None:
                raise KeyError(f"no QC run with id {run_id}")
            rows = conn.execute(
                """
                SELECT finding_id, version, payload, digest
                FROM counterfactual_bases WHERE run_id = ?
                """,
                (run_id,),
            ).fetchall()
        aggregate_digest = str(run["counterfactual_digest"] or "")
        if not rows:
            if not aggregate_digest:
                return {}
            if aggregate_digest != canonical_aggregate_digest({}):
                raise ValueError("counterfactual sidecar is missing")
            return {}
        if not aggregate_digest:
            raise ValueError("counterfactual aggregate digest is missing")
        result: dict[str, NumericCounterfactualBasis] = {}
        for row in rows:
            finding_id = str(row["finding_id"])
            try:
                payload = json.loads(row["payload"])
                basis = NumericCounterfactualBasis.model_validate(payload)
            except (json.JSONDecodeError, ValidationError) as exc:
                raise ValueError(
                    f"invalid counterfactual payload for finding {finding_id}"
                ) from exc
            if int(row["version"]) != basis.version:
                raise ValueError(
                    f"counterfactual version mismatch for finding {finding_id}"
                )
            if canonical_basis_digest(basis) != row["digest"]:
                raise ValueError(
                    f"counterfactual digest mismatch for finding {finding_id}"
                )
            result[finding_id] = basis
        if canonical_aggregate_digest(result) != aggregate_digest:
            raise ValueError("counterfactual aggregate digest mismatch")
        return result

    def get_series_anchors(self, run_id: int) -> dict[str, SeriesAnchor]:
        """Validated private series anchors for one run; fails closed to ``{}``.

        Missing, legacy, malformed or tampered evidence yields no anchors so the
        UI falls back to the canonical queue instead of guessing a cluster.
        """
        with self._connect() as conn:
            run = conn.execute(
                "SELECT * FROM runs WHERE id = ?",
                (run_id,),
            ).fetchone()
            if run is None:
                raise KeyError(f"no QC run with id {run_id}")
            rows = conn.execute(
                """
                SELECT finding_id, version, payload, digest
                FROM series_anchor_sidecars WHERE run_id = ?
                """,
                (run_id,),
            ).fetchall()
            aggregate_digest = str(run["series_anchor_digest"] or "")
            if not aggregate_digest:
                if rows:
                    _reject_series_anchors("legacy_run_with_rows")
                return {}
            if not rows:
                if aggregate_digest != canonical_series_aggregate_digest({}):
                    _reject_series_anchors("missing_rows")
                return {}

            needed = {str(row["finding_id"]) for row in rows}
            try:
                by_id = {
                    finding.finding_id: finding
                    for finding in self._load_findings(conn, run, decorate=None)
                    if finding.finding_id in needed
                }
            except (json.JSONDecodeError, ValidationError, FindingsStoreError):
                return _reject_series_anchors("unreadable_findings")

        bindings: dict[str, tuple[str, str, SeriesAnchor]] = {}
        anchors: dict[str, SeriesAnchor] = {}
        for row in rows:
            finding_id = str(row["finding_id"])
            finding = by_id.get(finding_id)
            if finding is None:
                return _reject_series_anchors("unknown_finding")
            try:
                anchor = _decode_series_anchor(json.loads(row["payload"]))
            except (json.JSONDecodeError, ValidationError, KeyError):
                return _reject_series_anchors("invalid_payload")
            if int(row["version"]) != anchor.version:
                return _reject_series_anchors("version_mismatch")
            member = finding.artifact_member
            location = finding.location or ""
            if canonical_series_anchor_digest(
                finding_id, member, location, anchor
            ) != row["digest"]:
                return _reject_series_anchors("entry_digest_mismatch")
            if not anchor_matches_finding(finding, anchor):
                return _reject_series_anchors("locator_mismatch")
            bindings[finding_id] = (member, location, anchor)
            anchors[finding_id] = anchor

        if canonical_series_aggregate_digest(bindings) != aggregate_digest:
            return _reject_series_anchors("aggregate_digest_mismatch")
        return anchors

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
                    relation, outcome, evidence_version, source_digest, applied_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(run_id, finding_id, source_run_id, source_finding_id)
                DO UPDATE SET
                    relation = excluded.relation,
                    outcome = excluded.outcome,
                    evidence_version = excluded.evidence_version,
                    source_digest = excluded.source_digest,
                    applied_at = excluded.applied_at
                """,
                [
                    (
                        run_id,
                        candidate.finding_id,
                        source_run_id,
                        candidate.source_finding_id,
                        AnnotationLineageRelation.IDENTITY.value,
                        AnnotationLineageOutcome.INHERITED.value,
                        candidate.evidence_version,
                        candidate.evidence_digest,
                        applied_at,
                    )
                    for candidate in candidates
                ],
            )
            # Record CARRIED origins for these applied annotations.
            conn.executemany(
                """
                INSERT INTO decision_origins (run_id, finding_id, origin, decided_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(run_id, finding_id) DO UPDATE SET
                    origin = excluded.origin,
                    decided_at = excluded.decided_at
                """,
                [
                    (
                        run_id,
                        candidate.finding_id,
                        DecisionOrigin.CARRIED.value,
                        applied_at,
                    )
                    for candidate in candidates
                ],
            )
        return len(candidates)

    def apply_population_carry_forward(
        self,
        run_id: int,
        source_run_id: int,
        candidate: PopulationCarryForwardCandidate,
    ) -> int:
        """Apply one population's carried decision plus its full member
        lineage. Only `outcome == inherited` is accepted -- every other
        outcome means the sources disagreed, were incomplete, or were
        unfinalized, so the analyst must decide fresh (Criterion 8: a
        population accepts exactly one decision, never a silent merge).
        """
        if candidate.outcome is not AnnotationLineageOutcome.INHERITED:
            raise ValueError(
                f"population {candidate.finding_id!r} outcome "
                f"{candidate.outcome.value!r} is not auto-appliable"
            )
        if not candidate.sources:
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
            conn.execute(
                """
                INSERT INTO annotations (run_id, finding_id, severity, comment, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(run_id, finding_id) DO UPDATE SET
                    severity = excluded.severity,
                    comment = excluded.comment,
                    updated_at = excluded.updated_at
                """,
                (run_id, candidate.finding_id, candidate.severity, candidate.comment, applied_at),
            )
            conn.executemany(
                """
                INSERT INTO annotation_lineage (
                    run_id, finding_id, source_run_id, source_finding_id,
                    relation, outcome, evidence_version, source_digest, applied_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(run_id, finding_id, source_run_id, source_finding_id)
                DO UPDATE SET
                    relation = excluded.relation,
                    outcome = excluded.outcome,
                    evidence_version = excluded.evidence_version,
                    source_digest = excluded.source_digest,
                    applied_at = excluded.applied_at
                """,
                [
                    (
                        run_id,
                        candidate.finding_id,
                        source_run_id,
                        source.source_finding_id,
                        AnnotationLineageRelation.MEMBER.value,
                        AnnotationLineageOutcome.INHERITED.value,
                        FINDING_EVIDENCE_VERSION,
                        source.source_digest,
                        applied_at,
                    )
                    for source in candidate.sources
                ],
            )
            conn.execute(
                """
                INSERT INTO decision_origins (run_id, finding_id, origin, decided_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(run_id, finding_id) DO UPDATE SET
                    origin = excluded.origin,
                    decided_at = excluded.decided_at
                """,
                (run_id, candidate.finding_id, DecisionOrigin.CARRIED.value, applied_at),
            )
        return len(candidate.sources)

    @staticmethod
    def _bounded_session_end(started_at: dt.datetime, observed: dt.datetime) -> dt.datetime:
        return min(observed, started_at + dt.timedelta(hours=4))

    def pause_review_sessions(
        self,
        *,
        now: dt.datetime | None = None,
        run_id: int | None = None,
    ) -> int:
        observed = now or dt.datetime.now(dt.UTC)
        with self._connect() as conn:
            if run_id is None:
                rows = conn.execute(
                    "SELECT id, started_at FROM review_sessions WHERE ended_at IS NULL"
                ).fetchall()
            else:
                rows = conn.execute(
                    """
                    SELECT id, started_at FROM review_sessions
                    WHERE ended_at IS NULL AND run_id = ?
                    """,
                    (run_id,),
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
        """Distinct findings carrying a decision -- not lineage row count,
        since one population finding can have many member lineage rows.
        """
        with self._connect() as conn:
            row = conn.execute(
                "SELECT COUNT(DISTINCT finding_id) FROM annotation_lineage"
                " WHERE run_id = ?",
                (run_id,),
            ).fetchone()
        return int(row[0]) if row is not None else 0

    def _backfill_longitudinal_rows(
        self,
        conn: sqlite3.Connection,
        rows: Iterable[sqlite3.Row],
    ) -> int:
        added = 0
        for row in rows:
            run_id = int(row["id"])
            annotations = {
                item["finding_id"]: item
                for item in conn.execute(
                    """
                    SELECT finding_id, updated_at FROM annotations
                    WHERE run_id = ?
                    """,
                    (run_id,),
                ).fetchall()
            }
            lineages = {
                item["finding_id"]: item
                for item in conn.execute(
                    """
                    SELECT finding_id, applied_at FROM annotation_lineage
                    WHERE run_id = ?
                    """,
                    (run_id,),
                ).fetchall()
            }
            existing_occurrences = {
                item["finding_id"]
                for item in conn.execute(
                    """
                    SELECT finding_id FROM decision_occurrences WHERE run_id = ?
                    """,
                    (run_id,),
                ).fetchall()
            }
            existing_origins = {
                item["finding_id"]
                for item in conn.execute(
                    "SELECT finding_id FROM decision_origins WHERE run_id = ?",
                    (run_id,),
                ).fetchall()
            }
            try:
                expected_occurrences = sum(
                    int(value) for value in json.loads(row["counts"]).values()
                )
            except (AttributeError, json.JSONDecodeError, TypeError, ValueError) as exc:
                raise ValueError("stored run counts cannot be indexed") from exc

            if len(existing_occurrences) < expected_occurrences:
                try:
                    profile_payload = json.loads(row["profile_snapshot"])
                    profile_snapshot = (
                        DeliverableProfile.model_validate(profile_payload)
                        if profile_payload is not None
                        else None
                    )
                    findings = self._load_findings(conn, row, decorate=None)
                except (
                    json.JSONDecodeError,
                    ValidationError,
                    FindingsStoreError,
                ) as exc:
                    raise ValueError("stored run evidence cannot be indexed") from exc
                contract_scope = contract_scope_for_profile(
                    (
                        profile_snapshot.name
                        if profile_snapshot is not None
                        else row["profile"]
                    ),
                    (
                        profile_snapshot.contract_id
                        if profile_snapshot is not None
                        else None
                    ),
                )
                for finding in findings:
                    if finding.finding_id in existing_occurrences:
                        continue
                    anchor = decision_anchor(finding)
                    conn.execute(
                        """
                        INSERT INTO decision_occurrences (
                            run_id, finding_id, contract_scope, anchor_version,
                            anchor_digest, evidence_version, evidence_digest
                        ) VALUES (?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            run_id,
                            finding.finding_id,
                            contract_scope,
                            anchor.version,
                            canonical_anchor_digest(anchor),
                            FINDING_EVIDENCE_VERSION,
                            finding_evidence_digest(finding),
                        ),
                    )
                    existing_occurrences.add(finding.finding_id)
                    added += 1

            for finding_id, annotation in annotations.items():
                if finding_id not in existing_occurrences:
                    continue
                if finding_id in existing_origins:
                    continue
                lineage = lineages.get(finding_id)
                origin = (
                    DecisionOrigin.CARRIED
                    if lineage is not None
                    and annotation["updated_at"] == lineage["applied_at"]
                    else DecisionOrigin.LEGACY_MANUAL
                )
                conn.execute(
                    """
                    INSERT INTO decision_origins (
                        run_id, finding_id, origin, decided_at
                    ) VALUES (?, ?, ?, ?)
                    """,
                    (
                        run_id,
                        finding_id,
                        origin.value,
                        annotation["updated_at"],
                    ),
                )
                added += 1
        return added

    def backfill_longitudinal(self, limit: int = 500) -> int:
        """Idempotently index at most ``limit`` recent legacy runs."""
        if limit < 1 or limit > 500:
            raise ValueError("longitudinal backfill limit must be between 1 and 500")
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT id, profile, counts, findings, profile_snapshot
                FROM runs ORDER BY id DESC LIMIT ?
                """,
                (limit,),
            ).fetchall()
            return self._backfill_longitudinal_rows(conn, rows)

    def _dossier_chain(self, current_run_id: int, limit: int) -> list[RunRecord]:
        if limit < 1 or limit > 50:
            raise ValueError("dossier limit must be between 1 and 50")
        current = self.get_raw_run(current_run_id)
        chain = [current]
        while chain[-1].rerun_of is not None and len(chain) < limit:
            try:
                chain.append(self.get_raw_run(int(chain[-1].rerun_of)))
            except KeyError as exc:
                raise ValueError("the Re-QC lineage is incomplete") from exc
        chain.reverse()
        return chain

    def get_dossier(
        self,
        current_run_id: int,
        current_finding_id: str,
        limit: int = 50,
    ) -> DossierResult:
        """Return typed oldest-to-current evidence on one unbroken Re-QC chain."""
        chain = self._dossier_chain(current_run_id, limit)
        current = chain[-1]
        current_finding = finding_by_id(current.findings, current_finding_id)
        if current_finding is None:
            raise KeyError("current finding not present in run")
        run_ids = [record.run_id for record in chain]
        placeholders = ",".join("?" for _ in run_ids)
        with self._connect() as conn:
            legacy_rows = conn.execute(
                f"""
                SELECT id, profile, counts, findings, profile_snapshot FROM runs
                WHERE id IN ({placeholders})
                """,
                run_ids,
            ).fetchall()
            self._backfill_longitudinal_rows(conn, legacy_rows)
            # A later read failure must not roll back a valid idempotent backfill.
            conn.commit()
            current_occurrence = conn.execute(
                """
                SELECT contract_scope, anchor_digest, evidence_digest
                FROM decision_occurrences
                WHERE run_id = ? AND finding_id = ?
                """,
                (current_run_id, current_finding_id),
            ).fetchone()
            if current_occurrence is None:
                raise RuntimeError("longitudinal occurrence index is incomplete")
            contract_scope = str(current_occurrence["contract_scope"])
            anchor_digest = str(current_occurrence["anchor_digest"])
            current_evidence = str(current_occurrence["evidence_digest"])
            entries: list[DossierEntry] = []
            for record in chain:
                occurrences = conn.execute(
                    """
                    SELECT finding_id, evidence_digest FROM decision_occurrences
                    WHERE contract_scope = ? AND anchor_digest = ? AND run_id = ?
                    ORDER BY finding_id
                    """,
                    (contract_scope, anchor_digest, record.run_id),
                ).fetchall()
                status = DossierStatus.NO_STORED_OBSERVATION
                stored_finding: Finding | None = None
                stored_finding_id: str | None = None
                if len(occurrences) > 1:
                    status = DossierStatus.AMBIGUOUS
                elif len(occurrences) == 1:
                    occurrence = occurrences[0]
                    stored_finding_id = str(occurrence["finding_id"])
                    status = (
                        DossierStatus.EXACT
                        if occurrence["evidence_digest"] == current_evidence
                        else DossierStatus.CHANGED
                    )
                    stored_finding = finding_by_id(
                        record.findings,
                        stored_finding_id,
                    )
                    if stored_finding is None:
                        status = DossierStatus.AMBIGUOUS
                        stored_finding_id = None

                analyst_severity: str | None = None
                analyst_comment = ""
                origin: DecisionOrigin | None = None
                carry_source_run_id: int | None = None
                carry_source_finding_id: str | None = None
                promotion_kind: str | None = None
                promotion_profile_sha256: str | None = None
                promoted_at: str | None = None
                if stored_finding_id is not None:
                    annotation = conn.execute(
                        """
                        SELECT severity, comment FROM annotations
                        WHERE run_id = ? AND finding_id = ?
                        """,
                        (record.run_id, stored_finding_id),
                    ).fetchone()
                    if annotation is not None:
                        analyst_severity = annotation["severity"]
                        analyst_comment = str(annotation["comment"] or "")
                    origin_row = conn.execute(
                        """
                        SELECT origin FROM decision_origins
                        WHERE run_id = ? AND finding_id = ?
                        """,
                        (record.run_id, stored_finding_id),
                    ).fetchone()
                    if origin_row is not None:
                        origin = DecisionOrigin(str(origin_row["origin"]))
                    lineage = conn.execute(
                        """
                        SELECT source_run_id, source_finding_id
                        FROM annotation_lineage
                        WHERE run_id = ? AND finding_id = ?
                        """,
                        (record.run_id, stored_finding_id),
                    ).fetchone()
                    if lineage is not None:
                        carry_source_run_id = int(lineage["source_run_id"])
                        carry_source_finding_id = str(lineage["source_finding_id"])
                    promotion = conn.execute(
                        """
                        SELECT kind, profile_sha256, promoted_at
                        FROM contract_promotions
                        WHERE run_id = ? AND finding_id = ?
                        ORDER BY id DESC LIMIT 1
                        """,
                        (record.run_id, stored_finding_id),
                    ).fetchone()
                    if promotion is not None:
                        promotion_kind = str(promotion["kind"])
                        promotion_profile_sha256 = str(promotion["profile_sha256"])
                        promoted_at = str(promotion["promoted_at"])
                entries.append(
                    DossierEntry(
                        run_id=record.run_id,
                        started_at=record.started_at.isoformat(timespec="seconds"),
                        status=status,
                        finalized=record.signoff is not None,
                        finding_id=stored_finding_id,
                        engine_severity=(
                            stored_finding.severity.value
                            if stored_finding is not None
                            and stored_finding.severity is not None
                            else None
                        ),
                        analyst_severity=analyst_severity,
                        analyst_comment=analyst_comment,
                        origin=origin,
                        carry_source_run_id=carry_source_run_id,
                        carry_source_finding_id=carry_source_finding_id,
                        baseline_value=(
                            stored_finding.baseline_value
                            if stored_finding is not None
                            else None
                        ),
                        current_value=(
                            stored_finding.current_value
                            if stored_finding is not None
                            else None
                        ),
                        waiver_reason=(
                            stored_finding.waiver_reason
                            if stored_finding is not None
                            else ""
                        ),
                        waiver_expires=(
                            stored_finding.waiver_expires
                            if stored_finding is not None
                            else ""
                        ),
                        promotion_kind=promotion_kind,
                        promotion_profile_sha256=promotion_profile_sha256,
                        promoted_at=promoted_at,
                    )
                    )
        return DossierResult(
            contract_scope=contract_scope,
            anchor_digest=anchor_digest,
            entries=tuple(entries),
        )

    def recurrence_eligibility(
        self,
        current_run_id: int,
        current_finding_id: str,
        required_occurrences: int = 3,
    ) -> RecurrenceEligibility | None:
        """Return a conservative promotion advisory, never a profile mutation."""
        if required_occurrences < 3 or required_occurrences > 50:
            raise ValueError("required occurrences must be between 3 and 50")
        dossier = self.get_dossier(
            current_run_id,
            current_finding_id,
            limit=required_occurrences,
        )
        return self.recurrence_eligibility_from_dossier(
            dossier,
            required_occurrences=required_occurrences,
        )

    @staticmethod
    def recurrence_eligibility_from_dossier(
        dossier: DossierResult,
        required_occurrences: int = 3,
    ) -> RecurrenceEligibility | None:
        """Derive the advisory from an already-loaded dossier."""
        if required_occurrences < 3 or required_occurrences > 50:
            raise ValueError("required occurrences must be between 3 and 50")
        entries = dossier.entries[-required_occurrences:]
        if len(entries) < required_occurrences:
            return None
        if any(entry.promotion_kind is not None for entry in entries):
            return None
        analyst_severity: str | None = None
        manual_count = 0
        run_ids: list[int] = []
        for entry in entries:
            if entry.status is not DossierStatus.EXACT or not entry.finalized:
                return None
            if entry.analyst_severity is None:
                return None
            if analyst_severity is None:
                analyst_severity = entry.analyst_severity
            elif analyst_severity != entry.analyst_severity:
                return None
            run_ids.append(entry.run_id)
            if entry.origin is DecisionOrigin.MANUAL:
                manual_count += 1
        if analyst_severity is None or manual_count < 2:
            return None
        return RecurrenceEligibility(
            run_ids=tuple(run_ids),
            manual_count=manual_count,
            analyst_severity=analyst_severity,
        )

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
                    "package_manifest": (
                        record.package_manifest.model_dump(mode="json")
                        if record.package_manifest is not None
                        else None
                    ),
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
