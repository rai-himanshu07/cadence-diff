"""Dependency-aware workload complexity for preflight safeguards.

The OOXML loader already refuses pathological raw sizes. This module measures
the cost drivers the raw size cannot see: formula count, lexical formulas,
reference operands, resolved range membership, projected concrete edge cost,
and interaction-rule count. Thresholds are derived from the recorded anonymous
benchmark, where a 35.5k-formula, 741k-operand, 15.8M-projected-edge workload
completes in about 20 seconds at about 350 MiB with the compact index.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl.utils.cell import range_boundaries

from qc_tool.excel.formula_tokens import (
    FormulaPatternKey,
    FormulaPrecedentKind,
    extract_formula_precedents,
    formula_pattern_key,
)
from qc_tool.io.model import WorkbookSnapshot
from qc_tool.progress import CancellationToken, check_cancelled


class WorkbookComplexityError(RuntimeError):
    """A workbook exceeds the dependency-cost refusal limits."""


@dataclass(slots=True)
class WorkbookComplexity:
    formula_count: int = 0
    lexical_formula_count: int = 0
    reference_operands: int = 0
    resolved_range_cells: int = 0
    projected_concrete_edges: int = 0
    interaction_rule_count: int = 0
    warning_reasons: tuple[str, ...] = ()
    override_used: bool = False
    #: Populated only when the caller passed ``sample_limit``.
    sampled_formulas: int = 0
    extras: dict[str, int] = field(default_factory=dict)

    @property
    def degraded(self) -> bool:
        return bool(self.warning_reasons) or self.override_used


#: (attribute, warn limit, refusal limit, human label). Warn limits sit at about
#: twice the measured representative workload; refusal limits at about ten times.
_COMPLEXITY_LIMITS: tuple[tuple[str, int, int, str], ...] = (
    ("formula_count", 75_000, 400_000, "formula cells"),
    ("reference_operands", 1_500_000, 8_000_000, "reference operands"),
    ("projected_concrete_edges", 40_000_000, 250_000_000, "projected cell dependencies"),
    ("interaction_rule_count", 5_000, 40_000, "data-validation and conditional-format rules"),
)


def _estimated_span(operand: str) -> int:
    """Cell count of a literal A1 range, or 1 when it is not a literal range.

    This is deliberately a cheap textual estimate: the safeguard must decide
    before the dependency index is built, so it never resolves names or tables.
    """
    _, _, ref = operand.rpartition("!")
    ref = ref.strip("@#$").replace("$", "")
    if ":" not in ref:
        return 1
    try:
        min_col, min_row, max_col, max_row = range_boundaries(ref)
    except ValueError:
        return 1
    if min_col is None or min_row is None or max_col is None or max_row is None:
        return 1
    return max(0, max_row - min_row + 1) * max(0, max_col - min_col + 1)


@dataclass(frozen=True, slots=True)
class _FormulaCost:
    operands: int
    range_cells: int
    projected_edges: int


def _formula_cost(formula: str) -> _FormulaCost:
    operands = 0
    range_cells = 0
    projected_edges = 0
    for precedent in extract_formula_precedents(formula).precedents:
        if precedent.kind is not FormulaPrecedentKind.REFERENCE:
            continue
        operands += 1
        span = _estimated_span(precedent.value)
        range_cells += span
        projected_edges += max(0, span - 1)
    return _FormulaCost(operands, range_cells, projected_edges)


def assess_workbook_complexity(
    workbook: WorkbookSnapshot,
    *,
    allow_complex_workbook: bool = False,
    cancellation_token: CancellationToken | None = None,
) -> WorkbookComplexity:
    """Estimate dependency cost before the expensive analysis phases run.

    Raises ``WorkbookComplexityError`` above the refusal limits unless the caller
    passes an explicit local override.
    """
    complexity = WorkbookComplexity(
        interaction_rule_count=(
            len(workbook.data_validations) + len(workbook.conditional_formats)
        )
    )
    costs: dict[FormulaPatternKey, _FormulaCost] = {}
    for sheet in workbook.sheets:
        check_cancelled(cancellation_token)
        for index, cell in enumerate(sheet.cells.values(), start=1):
            if index % 10_000 == 0:
                check_cancelled(cancellation_token)
            if cell.formula is None:
                continue
            complexity.formula_count += 1
            lowered = cell.formula.casefold()
            if "let(" in lowered or "lambda(" in lowered:
                complexity.lexical_formula_count += 1
            try:
                pattern = formula_pattern_key(cell.formula)
                cost = costs.get(pattern)
                if cost is None:
                    cost = _formula_cost(cell.formula)
                    costs[pattern] = cost
            except Exception:  # malformed formulas cost nothing to index
                continue
            complexity.reference_operands += cost.operands
            complexity.resolved_range_cells += cost.range_cells
            complexity.projected_concrete_edges += cost.projected_edges

    warnings: list[str] = []
    refusals: list[str] = []
    for attribute, warn_limit, refusal_limit, label in _COMPLEXITY_LIMITS:
        value = int(getattr(complexity, attribute))
        if value >= refusal_limit:
            refusals.append(f"{label} {value:,} >= refusal limit {refusal_limit:,}")
        elif value >= warn_limit:
            warnings.append(f"{label} {value:,} >= warning limit {warn_limit:,}")
    if refusals and not allow_complex_workbook:
        raise WorkbookComplexityError(
            f"{workbook.source_name}: dependency workload refused: "
            f"{'; '.join(refusals)}. Rerun with the explicit local "
            "allow_complex_workbook override only when sufficient memory and time "
            "are available."
        )
    if refusals:
        warnings.extend(f"override accepted: {reason}" for reason in refusals)
        complexity.override_used = True
    complexity.warning_reasons = tuple(warnings)
    return complexity
