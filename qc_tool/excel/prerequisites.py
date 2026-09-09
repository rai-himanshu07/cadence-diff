"""Cycle-comparison prerequisite checks.

A prerequisite pins one profile-configured selector/scenario cell that must
be present, non-blank, and exactly equal between baseline and current before
any alignment, diffing, reports, or history proceed. A mismatch means the
two files were not prepared under the same configuration (for example, a
different dropdown selection), so a value-level comparison would prove
nothing until the analyst aligns them.
"""

from __future__ import annotations

from qc_tool.config.profile import ComparisonPrerequisite
from qc_tool.io.model import WorkbookSnapshot
from qc_tool.run_action import RunActionItem


def _is_blank(value: object) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def check_comparison_prerequisites(
    baseline: WorkbookSnapshot,
    current: WorkbookSnapshot,
    prerequisites: list[ComparisonPrerequisite],
    *,
    member_id: str = "primary",
) -> list[RunActionItem]:
    """Return one bounded mismatch item per failing prerequisite.

    An empty result means every configured prerequisite matched. Only a
    bounded reason category is returned -- never the cell's actual value.
    """
    mismatches: list[RunActionItem] = []
    for prerequisite in prerequisites:
        base_cell = None
        curr_cell = None
        try:
            base_cell = baseline.sheet(prerequisite.sheet).cell(prerequisite.cell)
        except (KeyError, ValueError):
            base_cell = None
        try:
            curr_cell = current.sheet(prerequisite.sheet).cell(prerequisite.cell)
        except (KeyError, ValueError):
            curr_cell = None
        base_value = None if base_cell is None else base_cell.value
        curr_value = None if curr_cell is None else curr_cell.value
        if base_cell is None or curr_cell is None:
            reason = "the configured sheet or cell is missing from one or both files"
        elif _is_blank(base_value) or _is_blank(curr_value):
            reason = "the prerequisite cell is blank in one or both files"
        elif base_value != curr_value:
            reason = "baseline and current do not have the same prerequisite value"
        else:
            continue
        mismatches.append(
            RunActionItem(
                member_id=member_id,
                sheet=prerequisite.sheet,
                cell=prerequisite.cell,
                label=prerequisite.name,
                detail=reason,
            )
        )
    return mismatches
