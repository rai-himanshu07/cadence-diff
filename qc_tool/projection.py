"""Pre-run diff-volume projection from zip metadata, without parsing pairs.

Byte-identical worksheet parts (CRC + size from the central directory)
cannot produce findings, so the projection sums cell counts only over
sheets that actually differ — an honest upper bound computed in seconds,
used to warn before a monster comparison starts and to offer scoping.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass
from pathlib import Path

from qc_tool.io.ooxml_worksheet import parse_ooxml_worksheet_metadata


@dataclass(frozen=True, slots=True)
class VolumeProjection:
    changed_sheets: tuple[str, ...]
    added_sheets: tuple[str, ...]
    removed_sheets: tuple[str, ...]
    identical_sheets: tuple[str, ...]
    #: Upper bound: findings cannot exceed the differing sheets' cells.
    projected_max_findings: int
    #: Per-sheet upper bounds for changed and added sheets (picker detail).
    sheet_volumes: tuple[tuple[str, int], ...] = ()


def _part_signatures(path: Path) -> dict[str, tuple[int, int]]:
    with zipfile.ZipFile(path) as archive:
        return {info.filename: (info.CRC, info.file_size) for info in archive.infolist()}


def project_cycle_volume(baseline: Path, current: Path) -> VolumeProjection | None:
    """Projection for an Excel cycle pair; ``None`` when it cannot be computed.

    ``None`` (encrypted, XLSB, malformed) means no warning is shown — the
    projection is a courtesy, never a gate.
    """
    try:
        base_meta = parse_ooxml_worksheet_metadata(baseline.read_bytes())
        curr_meta = parse_ooxml_worksheet_metadata(current.read_bytes())
        base_parts = _part_signatures(baseline)
        curr_parts = _part_signatures(current)
    except Exception:
        return None
    base_sheets = {sheet.name: sheet for sheet in base_meta.sheets}
    curr_sheets = {sheet.name: sheet for sheet in curr_meta.sheets}
    changed: list[str] = []
    identical: list[str] = []
    volumes: list[tuple[str, int]] = []
    projected = 0
    for name, curr_sheet in curr_sheets.items():
        base_sheet = base_sheets.get(name)
        if base_sheet is None:
            continue
        if base_parts.get(base_sheet.part) == curr_parts.get(curr_sheet.part):
            identical.append(name)
            continue
        changed.append(name)
        volume = max(base_sheet.cell_count, curr_sheet.cell_count)
        volumes.append((name, volume))
        projected += volume
    added = [name for name in curr_sheets if name not in base_sheets]
    removed = [name for name in base_sheets if name not in curr_sheets]
    for name in added:
        volumes.append((name, curr_sheets[name].cell_count))
    projected += sum(curr_sheets[name].cell_count for name in added)
    projected += sum(base_sheets[name].cell_count for name in removed)
    return VolumeProjection(
        changed_sheets=tuple(changed),
        added_sheets=tuple(added),
        removed_sheets=tuple(removed),
        identical_sheets=tuple(identical),
        projected_max_findings=projected,
        sheet_volumes=tuple(
            sorted(volumes, key=lambda item: item[1], reverse=True)
        ),
    )
