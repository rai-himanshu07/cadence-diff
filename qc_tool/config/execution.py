"""Engine-native execution bindings resolved from a
``ResolvedInputConfigurationV1`` (plan-20260913, Step 3).

Bridges the per-run resolved configuration to the alignment/structural/
formula comparison engine: confirmed sheet renames and execution-derived
keyed-region identity rules. Keyed regions are adapted into the existing,
already-proven ``RowIdentityRule`` shape so they flow through
``qc_tool.excel.align``'s current keyed-alignment machinery unchanged,
rather than duplicating it.

Every consumer of ``ExecutionBindings`` treats it as fully optional
(``None`` = do exactly what happened before this module existed), so a
legacy run with no saved ``input_contract`` -- and therefore no resolved
member/sheet/region bindings -- stays byte-identical.

Known Step 3 scope boundary: the adapted ``RowIdentityRule`` carries one
shared ``header_row`` for both sides (a legacy-shape limitation), so a
region with genuinely different baseline/current first-data-rows is not
yet fully exploited here -- Step 8 extends alignment to consume
per-side boundaries directly rather than through this adapter.
"""

from __future__ import annotations

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

from qc_tool.config.profile import RowIdentityRule
from qc_tool.config.resolved_input import (
    ResolvedInputConfigurationV1,
    ResolvedMember,
    ResolvedRegion,
)


def _top_left(cell_range: str) -> tuple[int, int] | None:
    """(row, col) of a range or single-cell string's top-left corner, or
    ``None`` if unparseable.
    """
    if not cell_range:
        return None
    try:
        if ":" in cell_range:
            min_col, min_row, _, _ = range_boundaries(cell_range)
            if min_col is None or min_row is None:
                return None
            return (min_row, min_col)
        row, col = coordinate_to_tuple(cell_range)
        return (row, col)
    except ValueError:
        return None


def region_as_row_identity_rule(region: ResolvedRegion) -> RowIdentityRule | None:
    """Adapt one keyed, execution-confirmed ``ResolvedRegion`` into a
    ``RowIdentityRule``, or ``None`` when this region carries no
    execution-confirmed row identity (mode is not ``"keyed"``, or no column
    has ``alignment_role == "identity"``).
    """
    if region.mode != "keyed":
        return None
    identity_columns = [
        column.current_letter
        for column in region.columns
        if column.alignment_role == "identity" and column.current_letter
    ]
    if not identity_columns:
        return None
    ordinal_columns = [
        column.current_letter
        for column in region.columns
        if column.alignment_role == "ordinal" and column.current_letter
    ]
    anchor_source = region.current_data_range or region.current_outer_range
    if not anchor_source:
        return None
    anchor = _top_left(anchor_source)
    if anchor is None:
        return None
    anchor_row, anchor_col = anchor
    anchor_cell = f"{get_column_letter(anchor_col)}{anchor_row}"
    header_row: int | None = None
    if region.header_intent == "first_data_row" and region.current_first_data_row:
        header_row = region.current_first_data_row - 1
    return RowIdentityRule(
        anchor_cell=anchor_cell,
        header_row=header_row,
        identity_columns=identity_columns,
        ordinal_columns=ordinal_columns,
        duplicate_policy=region.duplicate_key_policy,
    )


class ExecutionBindings:
    """Read-only lookup over one run's resolved logical member/sheet/region
    bindings, organized for direct consumption by ``qc_tool.excel.align``,
    ``qc_tool.excel.diff_structure``, and ``qc_tool.excel.formulas``.
    """

    __slots__ = ("_members",)

    def __init__(self, resolved: ResolvedInputConfigurationV1) -> None:
        members: dict[str, ResolvedMember] = {}
        for member in resolved.members:
            members[member.member_id] = member
        self._members = members

    def confirmed_sheet_renames(self, member_id: str) -> dict[str, str]:
        """``{current_sheet_name: baseline_sheet_name}`` for every sheet of
        ``member_id`` whose two physical names both resolved and genuinely
        differ. A sheet sharing one name on both sides is ordinary
        same-name pairing, never a rename.
        """
        member = self._members.get(member_id)
        if member is None:
            return {}
        renames: dict[str, str] = {}
        for sheet in member.sheets:
            if (
                sheet.baseline_sheet_name
                and sheet.current_sheet_name
                and sheet.baseline_sheet_name != sheet.current_sheet_name
            ):
                renames[sheet.current_sheet_name] = sheet.baseline_sheet_name
        return renames

    def row_identity_rules(
        self, member_id: str, current_sheet_name: str
    ) -> tuple[RowIdentityRule, ...]:
        """Execution-confirmed keyed regions on ``current_sheet_name``,
        adapted into ``RowIdentityRule`` so they flow through the existing,
        already-proven keyed-alignment machinery unchanged.
        """
        member = self._members.get(member_id)
        if member is None:
            return ()
        rules: list[RowIdentityRule] = []
        for sheet in member.sheets:
            if sheet.current_sheet_name != current_sheet_name:
                continue
            for region in sheet.regions:
                rule = region_as_row_identity_rule(region)
                if rule is not None:
                    rules.append(rule)
        return tuple(rules)

    def logical_sheet_id(
        self, member_id: str, current_sheet_name: str
    ) -> str | None:
        """The saved logical ``sheet_id`` bound to this physical current
        sheet name, or ``None`` when unresolved. Used to build a bounded,
        content-free ``LogicalFindingAddress``.
        """
        member = self._members.get(member_id)
        if member is None:
            return None
        for sheet in member.sheets:
            if sheet.current_sheet_name == current_sheet_name:
                return sheet.sheet_id
        return None


def build_execution_bindings(
    resolved: ResolvedInputConfigurationV1 | None,
) -> ExecutionBindings | None:
    """``None`` in, ``None`` out -- the universal legacy-preserving guard
    every engine call site uses before doing anything binding-aware.
    """
    if resolved is None or resolved.source == "legacy_default":
        return None
    return ExecutionBindings(resolved)
