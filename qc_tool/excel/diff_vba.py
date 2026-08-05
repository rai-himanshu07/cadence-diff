"""Deterministic VBA module inventory and module-text comparison.

Evidence is reported as module identity, line counts, changed line ranges and a
content digest. Module source text is never copied into a finding, a report, or
history: macro code routinely carries paths, server names and credentials, and
the analyst can already open the workbook.
"""

from __future__ import annotations

from difflib import SequenceMatcher

from qc_tool.coverage import CoverageItem, CoverageState
from qc_tool.findings import Finding, FindingClass
from qc_tool.io.model import WorkbookSnapshot
from qc_tool.io.vba import VbaModule

#: Enough ranges to locate the change; the workbook itself is the full record.
_MAX_REPORTED_RANGES = 5


def _summary(module: VbaModule) -> str:
    return f"{module.line_count} lines (sha256 {module.digest[:12]})"


def _line_ranges(baseline: VbaModule, current: VbaModule) -> tuple[str, int, int]:
    base_lines = baseline.text.splitlines()
    current_lines = current.text.splitlines()
    matcher = SequenceMatcher(None, base_lines, current_lines, autojunk=False)
    added = removed = 0
    ranges: list[str] = []
    for tag, base_start, base_end, current_start, current_end in matcher.get_opcodes():
        if tag == "equal":
            continue
        added += current_end - current_start
        removed += base_end - base_start
        first = current_start + 1
        last = max(current_end, current_start + 1)
        ranges.append(str(first) if first >= last else f"{first}-{last}")
    shown = ranges[:_MAX_REPORTED_RANGES]
    location = ", ".join(shown)
    if len(ranges) > len(shown):
        location += f", and {len(ranges) - len(shown)} more"
    return location, added, removed


def diff_workbook_vba(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot
) -> list[Finding]:
    """Compare VBA modules by name, then by source text."""
    if not baseline.vba.available or not current.vba.available:
        return []
    base_modules = {module.name: module for module in baseline.vba.modules}
    current_modules = {module.name: module for module in current.vba.modules}
    findings: list[Finding] = []
    for name in sorted(base_modules.keys() | current_modules.keys()):
        before = base_modules.get(name)
        after = current_modules.get(name)
        if before is None and after is not None:
            message = f"VBA module {name!r} added ({after.line_count} lines)"
        elif after is None and before is not None:
            message = f"VBA module {name!r} removed (was {before.line_count} lines)"
        elif before is not None and after is not None:
            if before.digest == after.digest:
                continue
            location, added, removed = _line_ranges(before, after)
            message = (
                f"VBA module {name!r} changed: {added} lines added, "
                f"{removed} removed at {location}"
            )
        else:  # pragma: no cover - one side must exist
            continue
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.VBA_MODULE_CHANGED,
                element=name,
                baseline_value=_summary(before) if before is not None else None,
                current_value=_summary(after) if after is not None else None,
                message=message,
            )
        )
    return findings


def vba_coverage(*workbooks: WorkbookSnapshot) -> CoverageItem:
    """Report whether VBA module text could be read for every supplied workbook."""
    blocked = [book for book in workbooks if book.vba.present and not book.vba.available]
    degraded = [book for book in workbooks if book.vba.available and book.vba.detail]
    modules = sum(len(book.vba.modules) for book in workbooks)
    if blocked:
        state = CoverageState.UNAVAILABLE
        detail = "; ".join(f"{book.source_name}: {book.vba.detail}" for book in blocked)
    elif degraded:
        state = CoverageState.DEGRADED
        detail = "; ".join(f"{book.source_name}: {book.vba.detail}" for book in degraded)
    elif not any(book.vba.present for book in workbooks):
        state = CoverageState.CHECKED
        detail = "no VBA project is present"
    else:
        state = CoverageState.CHECKED
        detail = ""
    return CoverageItem(
        check_id="excel-vba",
        label="VBA module inventory and text",
        artifact="excel",
        state=state,
        findings=modules,
        detail=detail,
    )
