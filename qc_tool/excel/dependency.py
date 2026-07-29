"""Cross-sheet formula dependency graph and downstream-impact annotation.

Nodes are cells ``(sheet, row, col)``; an edge precedent -> dependent means
the dependent's formula references the precedent. Range references expand
to their member cells (bounded), sheet-qualified references resolve across
sheets, and named-range operands resolve through the workbook's defined
names. Findings on a cell gain ``impacts``: every downstream formula cell
that transitively consumes it.
"""

import logging
import re
from dataclasses import dataclass

import networkx as nx
from openpyxl.formula import Tokenizer
from openpyxl.utils import get_column_letter

from qc_tool.coverage import CoverageState
from qc_tool.excel.references import ReferenceStatus, ResolvedRange, resolve_reference
from qc_tool.findings import Finding, FindingClass
from qc_tool.io.model import WorkbookSnapshot
from qc_tool.progress import CancellationToken, check_cancelled

logger = logging.getLogger(__name__)

_CELL_REF_RE = re.compile(r"^[A-Z]{1,3}\d+$")
_MAX_RANGE_CELLS = 50_000
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


def _expand_range(resolved: ResolvedRange) -> list[Node]:
    return [
        (resolved.sheet, row, col)
        for row in range(resolved.min_row, resolved.max_row + 1)
        for col in range(resolved.min_col, resolved.max_col + 1)
    ]


class DependencyGraph(nx.DiGraph):
    def __init__(self) -> None:
        super().__init__()
        self.unsupported_references: list[str] = []
        self.invalid_references: list[str] = []
        self.parse_errors: list[str] = []
        self.symbolic_references: list[SymbolicReference] = []
        self.bounded_references: list[str] = []

    @property
    def coverage_state(self) -> CoverageState:
        return (
            CoverageState.DEGRADED
            if (
                self.unsupported_references
                or self.invalid_references
                or self.parse_errors
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
        if self.bounded_references:
            details.append(
                f"{len(self.bounded_references)} bounded whole-row/column references"
            )
        if self.unsupported_references:
            details.append(
                f"{len(self.unsupported_references)} unsupported references"
            )
        if self.invalid_references:
            details.append(f"{len(self.invalid_references)} invalid references")
        if self.parse_errors:
            details.append(f"{len(self.parse_errors)} unparseable formulas")
        return "; ".join(details) or "All parsed formula references were resolved"


def build_dependency_graph(
    workbook: WorkbookSnapshot,
    *,
    max_range_cells: int = _MAX_RANGE_CELLS,
    cancellation_token: CancellationToken | None = None,
) -> DependencyGraph:
    graph = DependencyGraph()
    for sheet in workbook.sheets:
        check_cancelled(cancellation_token)
        for index, ((row, col), cell) in enumerate(sheet.cells.items(), start=1):
            if index % 10_000 == 0:
                check_cancelled(cancellation_token)
            if cell.formula is None:
                continue
            dependent: Node = (sheet.name, row, col)
            try:
                tokens = Tokenizer(cell.formula).items
            except Exception:  # malformed formulas must not kill a run
                detail = f"{_display(dependent)}: {cell.formula}"
                if detail not in graph.parse_errors:
                    graph.parse_errors.append(detail)
                logger.warning(
                    "unparseable formula at %s: %r", _display(dependent), cell.formula
                )
                continue
            for token in tokens:
                if token.type != "OPERAND" or token.subtype != "RANGE":
                    continue
                value = token.value
                resolution = resolve_reference(
                    workbook,
                    value,
                    host_sheet=sheet.name,
                    host_cell=(row, col),
                    require_within_sheet=False,
                )
                if resolution.status is ReferenceStatus.UNSUPPORTED:
                    detail = f"{_display(dependent)}: {value}"
                    if detail not in graph.unsupported_references:
                        graph.unsupported_references.append(detail)
                    continue
                if resolution.status is ReferenceStatus.INVALID:
                    detail = f"{_display(dependent)}: {value}"
                    if detail not in graph.invalid_references:
                        graph.invalid_references.append(detail)
                    continue
                if resolution.detail and "bounded whole-row/column" in resolution.detail:
                    detail = f"{_display(dependent)}: {value}"
                    if detail not in graph.bounded_references:
                        graph.bounded_references.append(detail)
                for resolved in resolution.ranges:
                    if resolved.size > max_range_cells:
                        symbolic = SymbolicReference(resolved, dependent, value)
                        if symbolic not in graph.symbolic_references:
                            graph.symbolic_references.append(symbolic)
                        logger.warning(
                            "representing oversized range %s symbolically (%d cells)",
                            value,
                            resolved.size,
                        )
                        continue
                    for precedent in _expand_range(resolved):
                        if precedent != dependent:
                            graph.add_edge(precedent, dependent)
    return graph


def dependent_nodes_of(graph: "nx.DiGraph[Node]", node: Node) -> set[Node]:
    """Transitive dependents across concrete edges and symbolic aggregate ranges."""
    discovered: set[Node] = set()
    pending = [node]
    symbolic = (
        graph.symbolic_references
        if isinstance(graph, DependencyGraph)
        else []
    )
    while pending:
        current = pending.pop()
        concrete = set(graph.successors(current)) if current in graph else set()
        aggregate = {
            reference.dependent
            for reference in symbolic
            if reference.contains(current)
        }
        for dependent in concrete | aggregate:
            if dependent == node or dependent in discovered:
                continue
            discovered.add(dependent)
            pending.append(dependent)
    return discovered


def dependents_of(graph: "nx.DiGraph[Node]", sheet: str, ref: str) -> list[str]:
    """Transitive downstream dependents of ``sheet!ref``, as display strings."""
    from openpyxl.utils.cell import coordinate_to_tuple

    row, col = coordinate_to_tuple(ref)
    node: Node = (sheet, row, col)
    return sorted(_display(dependent) for dependent in dependent_nodes_of(graph, node))


def annotate_impacts(
    findings: list[Finding], graph: "nx.DiGraph[Node]"
) -> None:
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
