"""Compact cross-sheet formula dependency index and downstream-impact annotation.

Nodes are cells ``(sheet, row, col)``. A precedent -> dependent relationship
means the dependent's formula references the precedent. Single-cell references
are stored as direct edges; range, table, named, and supported dynamic
references are stored as rectangular descriptors indexed by sheet and bounded
row interval, so a 365-cell range costs one descriptor instead of 365 edges.
Findings on a cell gain ``impacts``: every downstream formula cell that
transitively consumes it.

``tests/dependency_oracle.py`` keeps the historical per-cell NetworkX expansion
as an independent parity oracle for supported synthetic cases. It is test-only
and must never be imported by production code.
"""

import logging
import re
from collections import Counter
from dataclasses import dataclass

from openpyxl.utils import get_column_letter

from qc_tool.coverage import CoverageState
from qc_tool.excel.formula_tokens import (
    FormulaPatternKey,
    FormulaPrecedentExtraction,
    FormulaPrecedentKind,
    extract_formula_precedents,
    formula_pattern_key,
)
from qc_tool.excel.references import (
    ReferenceStatus,
    ResolvedRange,
    build_reference_index,
    reference_reason_code,
    resolve_reference,
)
from qc_tool.findings import Finding, FindingClass
from qc_tool.io.model import WorkbookSnapshot
from qc_tool.progress import CancellationToken, check_cancelled

logger = logging.getLogger(__name__)

_CELL_REF_RE = re.compile(r"^[A-Z]{1,3}\d+$")
_MAX_RANGE_CELLS = 50_000
_MAX_FAILURE_SAMPLES = 8
_ROW_BUCKET = 64
_COLUMN_BUCKET = 16
_ANNOTATED_CLASSES = frozenset(
    {
        FindingClass.VALUE_CHANGED,
        FindingClass.FORMULA_ERROR,
        FindingClass.FORMULA_HARDCODED,
        FindingClass.FORMULA_REMOVED,
        FindingClass.FORMULA_MISSING,
        FindingClass.FORMULA_NOT_EXTENDED,
        FindingClass.FORMULA_LOGIC_CHANGED,
        FindingClass.FORMULA_INCONSISTENT,
    }
)

Node = tuple[str, int, int]
RangeDescriptor = tuple[int, int, int, int, int]


@dataclass(frozen=True, slots=True)
class SymbolicReference:
    source: ResolvedRange
    dependent: Node
    formula_reference: str

    def contains(self, node: Node) -> bool:
        sheet, row, column = node
        return (
            sheet == self.source.sheet
            and self.source.min_row <= row <= self.source.max_row
            and self.source.min_col <= column <= self.source.max_col
        )


def _display(node: Node) -> str:
    sheet, row, col = node
    return f"{sheet}!{get_column_letter(col)}{row}"


def _strip_sheet_quotes(name: str) -> str:
    if name.startswith("'") and name.endswith("'"):
        return name[1:-1].replace("''", "'")
    return name


class DependencyGraph:
    """Compact integer-keyed dependency index with rectangular descriptors."""

    def __init__(self) -> None:
        self.unsupported_references: list[str] = []
        self.invalid_references: list[str] = []
        self.parse_errors: list[str] = []
        self.reference_status_counts: Counter[ReferenceStatus] = Counter()
        self.unsupported_reason_counts: Counter[str] = Counter()
        self.invalid_reason_counts: Counter[str] = Counter()
        self.parse_error_counts: Counter[str] = Counter()
        self.local_symbol_count = 0
        self._unsupported_samples: set[str] = set()
        self._invalid_samples: set[str] = set()
        self._parse_error_samples: set[str] = set()
        self.symbolic_references: list[SymbolicReference] = []
        self._symbolic_reference_set: set[SymbolicReference] = set()
        self.bounded_references: list[str] = []
        self.bounded_reference_count = 0
        self._bounded_reference_samples: set[str] = set()
        self._node_ids: dict[Node, int] = {}
        self._nodes: list[Node] = []
        self._direct: dict[int, set[int]] = {}
        self._ranges: list[RangeDescriptor] = []
        self._range_sheets: list[str] = []
        self._range_seen: set[tuple[str, RangeDescriptor]] = set()
        self._buckets: dict[tuple[str, int, int], list[int]] = {}
        self._closure_cache: dict[Node, frozenset[Node]] = {}

    # --- construction -------------------------------------------------------

    def _node_id(self, node: Node) -> int:
        node_id = self._node_ids.get(node)
        if node_id is None:
            node_id = len(self._nodes)
            self._node_ids[node] = node_id
            self._nodes.append(node)
        return node_id

    def add_direct(self, precedent: Node, dependent: Node) -> None:
        if precedent == dependent:
            return
        self._direct.setdefault(self._node_id(precedent), set()).add(
            self._node_id(dependent)
        )

    def add_range(self, resolved: ResolvedRange, dependent: Node) -> None:
        """Store one rectangular precedent without materializing its members."""
        if resolved.size == 1:
            self.add_direct(
                (resolved.sheet, resolved.min_row, resolved.min_col), dependent
            )
            return
        descriptor: RangeDescriptor = (
            resolved.min_row,
            resolved.max_row,
            resolved.min_col,
            resolved.max_col,
            self._node_id(dependent),
        )
        if (resolved.sheet, descriptor) in self._range_seen:
            return
        self._range_seen.add((resolved.sheet, descriptor))
        index = len(self._ranges)
        self._ranges.append(descriptor)
        self._range_sheets.append(resolved.sheet)
        for row_bucket in range(
            resolved.min_row // _ROW_BUCKET, resolved.max_row // _ROW_BUCKET + 1
        ):
            for column_bucket in range(
                resolved.min_col // _COLUMN_BUCKET,
                resolved.max_col // _COLUMN_BUCKET + 1,
            ):
                self._buckets.setdefault(
                    (resolved.sheet, row_bucket, column_bucket), []
                ).append(index)

    # --- queries ------------------------------------------------------------

    @property
    def direct_edge_count(self) -> int:
        return sum(len(dependents) for dependents in self._direct.values())

    @property
    def range_descriptor_count(self) -> int:
        return len(self._ranges)

    def direct_dependents(self, node: Node) -> set[Node]:
        """Cells whose formula references ``node`` without transitive closure."""
        sheet, row, column = node
        dependent_ids: set[int] = set()
        node_id = self._node_ids.get(node)
        if node_id is not None:
            dependent_ids.update(self._direct.get(node_id, ()))
        bucket = (sheet, row // _ROW_BUCKET, column // _COLUMN_BUCKET)
        for index in self._buckets.get(bucket, ()):
            min_row, max_row, min_col, max_col, dependent_id = self._ranges[index]
            if min_row <= row <= max_row and min_col <= column <= max_col:
                dependent_ids.add(dependent_id)
        dependents = {self._nodes[dependent_id] for dependent_id in dependent_ids}
        dependents.update(
            reference.dependent
            for reference in self.symbolic_references
            if reference.contains(node)
        )
        dependents.discard(node)
        return dependents

    # --- telemetry ----------------------------------------------------------

    @staticmethod
    def _sample(target: list[str], seen: set[str], reason: str) -> None:
        if reason in seen:
            return
        seen.add(reason)
        if len(target) < _MAX_FAILURE_SAMPLES:
            target.append(reason)

    def record_reference(self, status: ReferenceStatus, detail: str) -> None:
        self.reference_status_counts[status] += 1
        if status is ReferenceStatus.INVALID:
            reason = reference_reason_code(status, detail)
            self.invalid_reason_counts[reason] += 1
            self._sample(self.invalid_references, self._invalid_samples, reason)
        elif status is ReferenceStatus.UNSUPPORTED:
            reason = reference_reason_code(status, detail)
            self.unsupported_reason_counts[reason] += 1
            self._sample(
                self.unsupported_references,
                self._unsupported_samples,
                reason,
            )

    def record_parse_error(self) -> None:
        reason = "formula_parse_error"
        self.parse_error_counts[reason] += 1
        self._sample(self.parse_errors, self._parse_error_samples, reason)

    def record_extraction_unsupported(self, reason: str) -> None:
        self.reference_status_counts[ReferenceStatus.UNSUPPORTED] += 1
        self.unsupported_reason_counts[reason] += 1
        self._sample(
            self.unsupported_references,
            self._unsupported_samples,
            reason,
        )

    def record_bounded_reference(self) -> None:
        reason = "bounded_whole_row_or_column"
        self.bounded_reference_count += 1
        self._sample(
            self.bounded_references,
            self._bounded_reference_samples,
            reason,
        )

    def record_symbolic_reference(
        self, resolved: ResolvedRange, dependent: Node, formula_reference: str
    ) -> None:
        symbolic = SymbolicReference(resolved, dependent, formula_reference)
        if symbolic in self._symbolic_reference_set:
            return
        self._symbolic_reference_set.add(symbolic)
        self.symbolic_references.append(symbolic)

    @property
    def coverage_state(self) -> CoverageState:
        return (
            CoverageState.DEGRADED
            if (
                self.reference_status_counts[ReferenceStatus.UNSUPPORTED]
                or self.reference_status_counts[ReferenceStatus.INVALID]
                or self.parse_error_counts
            )
            else CoverageState.CHECKED
        )

    @property
    def coverage_detail(self) -> str:
        details: list[str] = []
        if self.symbolic_references:
            details.append(
                f"{len(self.symbolic_references)} symbolic aggregate references"
            )
        if self.bounded_reference_count:
            details.append(
                f"{self.bounded_reference_count} bounded whole-row/column references"
            )
        unsupported = self.reference_status_counts[ReferenceStatus.UNSUPPORTED]
        if unsupported:
            details.append(
                f"{unsupported} unsupported references"
            )
        invalid = self.reference_status_counts[ReferenceStatus.INVALID]
        if invalid:
            details.append(f"{invalid} invalid references")
        parse_errors = sum(self.parse_error_counts.values())
        if parse_errors:
            details.append(f"{parse_errors} unparseable formulas")
        return "; ".join(details) or "All parsed formula references were resolved"


def build_dependency_graph(
    workbook: WorkbookSnapshot,
    *,
    max_range_cells: int = _MAX_RANGE_CELLS,
    cancellation_token: CancellationToken | None = None,
) -> DependencyGraph:
    graph = DependencyGraph()
    reference_index = build_reference_index(workbook)
    formula_templates: dict[FormulaPatternKey, FormulaPrecedentExtraction] = {}
    for sheet in workbook.sheets:
        check_cancelled(cancellation_token)
        for index, ((row, col), cell) in enumerate(sheet.cells.items(), start=1):
            if index % 10_000 == 0:
                check_cancelled(cancellation_token)
            if cell.formula is None:
                continue
            dependent: Node = (sheet.name, row, col)
            try:
                pattern = formula_pattern_key(cell.formula)
                extraction = formula_templates.get(pattern)
                if extraction is None:
                    extraction = extract_formula_precedents(cell.formula)
                    formula_templates[pattern] = extraction
            except Exception:  # malformed formulas must not kill a run
                graph.record_parse_error()
                logger.warning("unparseable formula at %s", _display(dependent))
                continue
            for precedent in extraction.precedents:
                if precedent.kind is FormulaPrecedentKind.LOCAL_SYMBOL:
                    graph.local_symbol_count += 1
                    continue
                if precedent.kind is FormulaPrecedentKind.UNSUPPORTED:
                    graph.record_extraction_unsupported(precedent.reason)
                    continue
                value = precedent.value
                resolution = resolve_reference(
                    workbook,
                    value,
                    host_sheet=sheet.name,
                    host_cell=(row, col),
                    require_within_sheet=False,
                    index=reference_index,
                )
                graph.record_reference(resolution.status, resolution.detail)
                if resolution.status is not ReferenceStatus.RESOLVED:
                    continue
                if resolution.detail and "bounded whole-row/column" in resolution.detail:
                    graph.record_bounded_reference()
                for resolved in resolution.ranges:
                    if resolved.size > max_range_cells:
                        graph.record_symbolic_reference(resolved, dependent, value)
                        logger.warning(
                            "representing oversized range %s symbolically (%d cells)",
                            value,
                            resolved.size,
                        )
                        continue
                    graph.add_range(resolved, dependent)
    return graph


def dependent_nodes_of(graph: DependencyGraph, node: Node) -> set[Node]:
    """Transitive dependents across direct edges, range descriptors, aggregates.

    Closures are memoized per requested source node only, so bulk impact
    annotation never materializes closures for cells nobody asked about.
    """
    cached = graph._closure_cache.get(node)
    if cached is not None:
        return set(cached)
    discovered: set[Node] = set()
    pending = [node]
    while pending:
        current = pending.pop()
        for dependent in graph.direct_dependents(current):
            if dependent == node or dependent in discovered:
                continue
            discovered.add(dependent)
            pending.append(dependent)
    graph._closure_cache[node] = frozenset(discovered)
    return discovered


def dependents_of(graph: DependencyGraph, sheet: str, ref: str) -> list[str]:
    """Transitive downstream dependents of ``sheet!ref``, as display strings."""
    from openpyxl.utils.cell import coordinate_to_tuple

    row, col = coordinate_to_tuple(ref)
    node: Node = (sheet, row, col)
    return sorted(_display(dependent) for dependent in dependent_nodes_of(graph, node))


def annotate_impacts(findings: list[Finding], graph: DependencyGraph) -> None:
    """Attach downstream impact lists to cell-level findings (in place)."""
    for finding in findings:
        if finding.finding_class not in _ANNOTATED_CLASSES:
            continue
        if finding.sheet is None or finding.location is None:
            continue
        if not _CELL_REF_RE.match(finding.location):
            continue
        impacts = dependents_of(graph, finding.sheet, finding.location)
        finding.impacts = sorted({*finding.impacts, *impacts})


def limit_impacts(findings: list[Finding], *, max_impacts: int = 25) -> None:
    """Apply one accurate display cap after every impact source has merged."""
    for finding in findings:
        impacts = sorted(
            impact
            for impact in finding.impacts
            if not impact.startswith("... ")
        )
        if len(impacts) > max_impacts:
            finding.impacts = [
                *impacts[:max_impacts],
                f"... {len(impacts) - max_impacts} more",
            ]
        else:
            finding.impacts = impacts
