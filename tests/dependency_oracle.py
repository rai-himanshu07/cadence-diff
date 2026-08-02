"""Test-only dependency oracle: the historical per-cell NetworkX expansion.

Production code must never import this module. It exists so the compact index in
``qc_tool.excel.dependency`` can be proven to return identical direct and
transitive impacts for supported synthetic workbooks.
``tests/test_lint.py`` asserts that only the parity tests import it.
"""

from __future__ import annotations

import networkx as nx
from openpyxl.utils import get_column_letter

from qc_tool.excel.formula_tokens import (
    FormulaPrecedentKind,
    extract_formula_precedents,
)
from qc_tool.excel.references import (
    ReferenceStatus,
    ResolvedRange,
    build_reference_index,
    resolve_reference,
)
from qc_tool.io.model import WorkbookSnapshot

Node = tuple[str, int, int]
_MAX_RANGE_CELLS = 50_000


def _expand(resolved: ResolvedRange) -> list[Node]:
    return [
        (resolved.sheet, row, column)
        for row in range(resolved.min_row, resolved.max_row + 1)
        for column in range(resolved.min_col, resolved.max_col + 1)
    ]


def build_oracle_graph(
    workbook: WorkbookSnapshot, *, max_range_cells: int = _MAX_RANGE_CELLS
) -> nx.DiGraph:
    """Expand every resolved range into one concrete edge per member cell."""
    graph: nx.DiGraph = nx.DiGraph()
    graph.graph["symbolic"] = []
    reference_index = build_reference_index(workbook)
    for sheet in workbook.sheets:
        for (row, column), cell in sheet.cells.items():
            if cell.formula is None:
                continue
            dependent: Node = (sheet.name, row, column)
            try:
                extraction = extract_formula_precedents(cell.formula)
            except Exception:
                continue
            for precedent in extraction.precedents:
                if precedent.kind is not FormulaPrecedentKind.REFERENCE:
                    continue
                resolution = resolve_reference(
                    workbook,
                    precedent.value,
                    host_sheet=sheet.name,
                    host_cell=(row, column),
                    require_within_sheet=False,
                    index=reference_index,
                )
                if resolution.status is not ReferenceStatus.RESOLVED:
                    continue
                for resolved in resolution.ranges:
                    if resolved.size > max_range_cells:
                        graph.graph["symbolic"].append((resolved, dependent))
                        continue
                    for member in _expand(resolved):
                        if member != dependent:
                            graph.add_edge(member, dependent)
    return graph


def oracle_dependents(graph: nx.DiGraph, sheet: str, ref: str) -> list[str]:
    """Sorted transitive dependents of ``sheet!ref`` as display strings."""
    from openpyxl.utils.cell import coordinate_to_tuple

    row, column = coordinate_to_tuple(ref)
    node: Node = (sheet, row, column)
    discovered: set[Node] = set()
    pending = [node]
    symbolic = graph.graph["symbolic"]
    while pending:
        current = pending.pop()
        successors = set(graph.successors(current)) if current in graph else set()
        successors.update(
            dependent
            for resolved, dependent in symbolic
            if resolved.sheet == current[0]
            and resolved.min_row <= current[1] <= resolved.max_row
            and resolved.min_col <= current[2] <= resolved.max_col
        )
        for dependent in successors:
            if dependent == node or dependent in discovered:
                continue
            discovered.add(dependent)
            pending.append(dependent)
    return sorted(
        f"{name}!{get_column_letter(column)}{row}"
        for name, row, column in discovered
    )
