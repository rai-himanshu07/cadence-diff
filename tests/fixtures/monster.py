"""Deterministic monster workbook pair for memory-behavior probes.

Values-only pair: column A is a stable text key, remaining columns hold
numeric values derived from the coordinates, and the current file shifts
every numeric cell by a fixed delta so the diff population is exactly
``rows * (columns - 1)`` value changes with row-key alignment.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook


def generate_monster_pair(
    dest_dir: Path,
    *,
    rows: int = 50_000,
    columns: int = 8,
    delta: float = 0.01,
) -> tuple[Path, Path]:
    if rows < 1 or columns < 2:
        raise ValueError("monster pair needs at least one row and two columns")
    dest_dir.mkdir(parents=True, exist_ok=True)
    baseline = dest_dir / f"monster-baseline-{rows}x{columns}.xlsx"
    current = dest_dir / f"monster-current-{rows}x{columns}.xlsx"
    for path, shift in ((baseline, 0.0), (current, delta)):
        workbook = Workbook(write_only=True)
        sheet = workbook.create_sheet("Data")
        sheet.append(
            ["Key", *(f"Metric {index}" for index in range(1, columns))]
        )
        for row in range(1, rows + 1):
            sheet.append(
                [
                    f"row-{row}",
                    *(
                        round(row * 31 + column * 7 + shift, 2)
                        for column in range(1, columns)
                    ),
                ]
            )
        workbook.save(path)
    return baseline, current
