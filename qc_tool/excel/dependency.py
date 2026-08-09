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

import hashlib
import json
import logging
import re
from collections import Counter, defaultdict
from dataclasses import dataclass

from openpyxl.utils import get_column_letter

from qc_tool.coverage import CoverageItem, CoverageState
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
_MAX_CYCLE_EDGES = 2_000_000
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


@dataclass(frozen=True, slots=True)
class CircularReferenceResult:
    findings: tuple[Finding, ...]
    coverage: CoverageItem


class FormulaCycleGraph:
    """Formula-only adjacency; ranges query spatially indexed formula nodes."""

    def __init__(
        self,
        formula_nodes: list[Node],
        *,
        max_edges: int = _MAX_CYCLE_EDGES,
    ) -> None:
        self._nodes = sorted(formula_nodes)
        self._node_ids = {node: index for index, node in enumerate(self._nodes)}
        self._adjacency: list[set[int]] = [set() for _ in self._nodes]
        self._buckets: dict[tuple[str, int, int], list[int]] = defaultdict(list)
        for node_id, (sheet, row, column) in enumerate(self._nodes):
            self._buckets[
                (sheet, row // _ROW_BUCKET, column // _COLUMN_BUCKET)
            ].append(node_id)
        self._max_edges = max_edges
        self._edge_count = 0
        self.omitted_range_references = 0

    @property
    def node_count(self) -> int:
        return len(self._nodes)

    @property
    def edge_count(self) -> int:
        return self._edge_count

    def add_range(self, resolved: ResolvedRange, dependent: Node) -> None:
        dependent_id = self._node_ids.get(dependent)
        if dependent_id is None:
            raise ValueError("cycle dependent is not a formula node")
        precedent_ids: set[int] = set()
        for row_bucket in range(
            resolved.min_row // _ROW_BUCKET,
            resolved.max_row // _ROW_BUCKET + 1,
        ):
            for column_bucket in range(
                resolved.min_col // _COLUMN_BUCKET,
                resolved.max_col // _COLUMN_BUCKET + 1,
            ):
                for node_id in self._buckets.get(
                    (resolved.sheet, row_bucket, column_bucket), ()
                ):
                    _sheet, row, column = self._nodes[node_id]
                    if (
                        resolved.min_row <= row <= resolved.max_row
                        and resolved.min_col <= column <= resolved.max_col
                    ):
                        precedent_ids.add(node_id)
        new_ids = precedent_ids - self._adjacency[dependent_id]
        if self._edge_count + len(new_ids) > self._max_edges:
            self.omitted_range_references += 1
            return
        self._adjacency[dependent_id].update(new_ids)
        self._edge_count += len(new_ids)

    def strongly_connected_components(
        self,
        cancellation_token: CancellationToken | None = None,
    ) -> tuple[tuple[Node, ...], ...]:
        """Return deterministic cyclic SCCs using iterative Kosaraju passes."""
        adjacency = [tuple(sorted(neighbors)) for neighbors in self._adjacency]
        visited: set[int] = set()
        order: list[int] = []
        operations = 0
        for start in range(len(self._nodes)):
            if start in visited:
                continue
            visited.add(start)
            stack: list[tuple[int, int]] = [(start, 0)]
            while stack:
                operations += 1
                if operations % 10_000 == 0:
                    check_cancelled(cancellation_token)
                node_id, neighbor_index = stack[-1]
                neighbors = adjacency[node_id]
                if neighbor_index < len(neighbors):
                    neighbor = neighbors[neighbor_index]
                    stack[-1] = (node_id, neighbor_index + 1)
                    if neighbor not in visited:
                        visited.add(neighbor)
                        stack.append((neighbor, 0))
                    continue
                stack.pop()
                order.append(node_id)

        reverse: list[list[int]] = [[] for _ in self._nodes]
        for source, neighbors in enumerate(adjacency):
            for target in neighbors:
                reverse[target].append(source)
        for neighbors in reverse:
            neighbors.sort()

        visited.clear()
        components: list[tuple[Node, ...]] = []
        for start in reversed(order):
            if start in visited:
                continue
            visited.add(start)
            pending = [start]
            member_ids: list[int] = []
            while pending:
                operations += 1
                if operations % 10_000 == 0:
                    check_cancelled(cancellation_token)
                node_id = pending.pop()
                member_ids.append(node_id)
                for neighbor in reversed(reverse[node_id]):
                    if neighbor not in visited:
                        visited.add(neighbor)
                        pending.append(neighbor)
            is_self_cycle = (
                len(member_ids) == 1
                and member_ids[0] in self._adjacency[member_ids[0]]
            )
            if len(member_ids) > 1 or is_self_cycle:
                components.append(
                    tuple(sorted(self._nodes[node_id] for node_id in member_ids))
                )
        return tuple(sorted(components, key=lambda component: component[0]))


def _display(node: Node) -> str:
    sheet, row, col = node
    return f"{sheet}!{get_column_letter(col)}{row}"


def _strip_sheet_quotes(name: str) -> str:
    if name.startswith("'") and name.endswith("'"):
        return name[1:-1].replace("''", "'")
    return name


class DependencyGraph:
    """Compact integer-keyed dependency index with rectangular descriptors."""

    def __init__(self, formula_nodes: list[Node] | None = None) -> None:
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
        self.formula_cycles = FormulaCycleGraph(formula_nodes or [])

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
    formula_nodes = [
        (sheet.name, row, column)
        for sheet in workbook.sheets
        for (row, column), cell in sheet.cells.items()
        if cell.formula is not None
    ]
    graph = DependencyGraph(formula_nodes)
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
                    graph.formula_cycles.add_range(resolved, dependent)
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


def detect_circular_references(
    graph: DependencyGraph,
    *,
    cancellation_token: CancellationToken | None = None,
) -> CircularReferenceResult:
    """Report each exact formula SCC once; disclose every incomplete edge source."""
    components = graph.formula_cycles.strongly_connected_components(
        cancellation_token
    )
    findings: list[Finding] = []
    for component in components:
        displays = [_display(node) for node in component]
        canonical = json.dumps(displays, separators=(",", ":"))
        digest = hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:12]
        event_key = f"excel-cycle:{digest}"
        sheet, row, column = component[0]
        sample = displays[:8]
        omitted = len(displays) - len(sample)
        element = "; ".join(sample)
        if omitted:
            element += f"; +{omitted} more cells"
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.CIRCULAR_REFERENCE,
                sheet=sheet,
                location=f"{get_column_letter(column)}{row}",
                element=element,
                current_value=f"{len(component)} formula cells",
                event_key=event_key,
                root_cause_key=event_key,
                message=(
                    f"Circular formula dependency contains {len(component)} "
                    "formula cells"
                ),
            )
        )

    unsupported = graph.reference_status_counts[ReferenceStatus.UNSUPPORTED]
    invalid = graph.reference_status_counts[ReferenceStatus.INVALID]
    parse_errors = sum(graph.parse_error_counts.values())
    omitted_ranges = graph.formula_cycles.omitted_range_references
    incomplete = bool(unsupported or invalid or parse_errors or omitted_ranges)
    details = [
        f"{graph.formula_cycles.node_count} formula nodes",
        f"{graph.formula_cycles.edge_count} exact formula edges",
        f"{len(components)} circular components",
    ]
    if unsupported:
        details.append(f"{unsupported} unsupported references")
    if invalid:
        details.append(f"{invalid} invalid references")
    if parse_errors:
        details.append(f"{parse_errors} unparseable formulas")
    if omitted_ranges:
        details.append(
            f"{omitted_ranges} formula ranges exceeded the cycle-edge budget"
        )
    coverage = CoverageItem(
        check_id="excel-circular-references",
        label="Circular formula references",
        artifact="excel",
        state=CoverageState.DEGRADED if incomplete else CoverageState.CHECKED,
        findings=len(findings),
        detail="; ".join(details),
    )
    return CircularReferenceResult(findings=tuple(findings), coverage=coverage)


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
