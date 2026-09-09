"""Step 8: CLI rendering of a blocked run's action-required payload.

Covers the compatibility contract between ``qc_tool.cli._print_blocked_action``
and ``qc_tool.run_action``'s v1 (flat fields) and v2 (``ranked_table_evidence``)
shapes, plus Criterion 17 (name the package member whenever it is not
``"primary"``, in the CLI surface).
"""

from __future__ import annotations

import pytest

from qc_tool.cli import _print_blocked_action
from qc_tool.run_action import (
    RankedTableEvidence,
    RunActionItem,
    RunActionReason,
    RunActionRequired,
)


def test_print_blocked_action_renders_v2_evidence_without_the_raw_detail_sentence(
    capsys: pytest.CaptureFixture[str],
) -> None:
    action = RunActionRequired(
        version=2,
        reason=RunActionReason.ROW_IDENTITY_CONFIRMATION_REQUIRED,
        items=[
            RunActionItem(
                sheet="Panel",
                cell="A1",
                label="Possible ranked/sorted table: columns B",
                ranked_table_evidence=RankedTableEvidence(
                    sheet="Panel",
                    current_range="A1:E6001",
                    data_row_count=6001,
                    available_columns=("A", "B", "C", "D", "E"),
                    suggested_identity_columns=("B",),
                    suggested_ordinal_columns=("A",),
                    non_blank_coverage=0.999,
                    unique_ratio=0.998,
                    key_overlap=0.95,
                    mismatch_reduction=0.995,
                    projected_positional_mismatches=500_000,
                    projected_avoided_mismatches=497_500,
                ),
            )
        ],
        message="One or more sheets look like a ranked or sorted table.",
    )
    _print_blocked_action(action)
    err = capsys.readouterr().err
    assert "Panel!A1" in err
    assert "6,001 rows" in err
    assert "5 columns available" in err
    assert "suggested identity=B" in err
    assert "ordinal=A" in err
    # No legacy snake_case telemetry sentence leaks through when v2 evidence
    # is present (detail was never populated for this item).
    assert "non_blank_coverage=" not in err
    assert "projected_positional_mismatches=" not in err


def test_print_blocked_action_falls_back_to_v1_detail_when_evidence_is_absent(
    capsys: pytest.CaptureFixture[str],
) -> None:
    """A legacy stored payload (predating the v2 field) still renders."""
    action = RunActionRequired(
        reason=RunActionReason.COMPARISON_PREREQUISITE_MISMATCH,
        items=[
            RunActionItem(
                sheet="Data",
                cell="C3",
                detail="baseline cell is blank",
            )
        ],
        message="Comparison prerequisites were not met.",
    )
    _print_blocked_action(action)
    err = capsys.readouterr().err
    assert "Data!C3" in err
    assert "baseline cell is blank" in err


def test_print_blocked_action_names_the_member_when_not_primary(
    capsys: pytest.CaptureFixture[str],
) -> None:
    action = RunActionRequired(
        reason=RunActionReason.COMPARISON_PREREQUISITE_MISMATCH,
        items=[
            RunActionItem(
                member_id="ops",
                sheet="Data",
                cell="C3",
                detail="baseline cell is blank",
            ),
            RunActionItem(
                sheet="Data",
                cell="D4",
                detail="current cell is blank",
            ),
        ],
        message="Comparison prerequisites were not met.",
    )
    _print_blocked_action(action)
    err = capsys.readouterr().err
    assert "[member=ops]" in err
    lines = err.splitlines()
    primary_line = next(line for line in lines if "D4" in line)
    assert "[member=" not in primary_line
