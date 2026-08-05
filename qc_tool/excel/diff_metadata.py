"""Deterministic comparison of comments, Power Query definitions, and connections.

Connection evidence deliberately stops at a classification and a digest. The
connection string, URL, and command text are never read into a finding, so a
server name or credential cannot reach a report, a history row, or a log line.
"""

from __future__ import annotations

from qc_tool.coverage import CoverageItem, CoverageState
from qc_tool.findings import Finding, FindingClass
from qc_tool.io.model import WorkbookSnapshot
from qc_tool.io.ooxml_metadata import CellComment, PowerQuery, WorkbookConnection


def _comment_findings(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot
) -> list[Finding]:
    before = {item.location: item for item in baseline.metadata.comments}
    after = {item.location: item for item in current.metadata.comments}
    findings: list[Finding] = []
    for location in sorted(before.keys() | after.keys()):
        old = before.get(location)
        new = after.get(location)
        if old is not None and new is not None and old.text == new.text:
            continue
        sheet, _, ref = location.partition("!")
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.CELL_COMMENT_CHANGED,
                sheet=sheet,
                location=ref,
                element=location,
                baseline_value=old.text if old is not None else None,
                current_value=new.text if new is not None else None,
                message=_comment_message(location, old, new),
            )
        )
    return findings


def _comment_message(
    location: str, old: CellComment | None, new: CellComment | None
) -> str:
    if old is None:
        return f"comment added at {location}"
    if new is None:
        return f"comment removed at {location}"
    return f"comment text changed at {location}"


def _query_findings(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot
) -> list[Finding]:
    before = {item.name: item for item in baseline.metadata.queries}
    after = {item.name: item for item in current.metadata.queries}
    findings: list[Finding] = []
    for name in sorted(before.keys() | after.keys()):
        old = before.get(name)
        new = after.get(name)
        if old is not None and new is not None and old.digest == new.digest:
            continue
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.POWER_QUERY_CHANGED,
                element=name,
                baseline_value=_query_summary(old),
                current_value=_query_summary(new),
                message=_query_message(name, old, new),
            )
        )
    return findings


def _query_summary(query: PowerQuery | None) -> str | None:
    if query is None:
        return None
    return f"{query.line_count} lines (sha256 {query.digest})"


def _query_message(name: str, old: PowerQuery | None, new: PowerQuery | None) -> str:
    if old is None:
        return f"Power Query {name!r} added"
    if new is None:
        return f"Power Query {name!r} removed"
    return f"Power Query {name!r} definition changed"


def _connection_findings(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot
) -> list[Finding]:
    before = {item.name: item for item in baseline.metadata.connections}
    after = {item.name: item for item in current.metadata.connections}
    findings: list[Finding] = []
    for name in sorted(before.keys() | after.keys()):
        old = before.get(name)
        new = after.get(name)
        if old is not None and new is not None and old == new:
            continue
        findings.append(
            Finding(
                artifact="excel",
                finding_class=FindingClass.CONNECTION_CHANGED,
                element=name,
                baseline_value=_connection_summary(old),
                current_value=_connection_summary(new),
                message=_connection_message(name, old, new),
            )
        )
    return findings


def _connection_summary(connection: WorkbookConnection | None) -> str | None:
    if connection is None:
        return None
    scope = "external" if connection.external else "local"
    return f"{connection.kind} ({scope}, target sha256 {connection.target_digest})"


def _connection_message(
    name: str, old: WorkbookConnection | None, new: WorkbookConnection | None
) -> str:
    if old is None:
        return f"data connection {name!r} added"
    if new is None:
        return f"data connection {name!r} removed"
    if old.target_digest != new.target_digest:
        return f"data connection {name!r} points at a different target"
    return f"data connection {name!r} settings changed"


def diff_workbook_metadata(
    baseline: WorkbookSnapshot, current: WorkbookSnapshot
) -> list[Finding]:
    """Compare comments, query definitions, and connections across the pair."""
    findings: list[Finding] = []
    if baseline.metadata.comments_available and current.metadata.comments_available:
        findings += _comment_findings(baseline, current)
    if baseline.metadata.queries_available and current.metadata.queries_available:
        findings += _query_findings(baseline, current)
    if (
        baseline.metadata.connections_available
        and current.metadata.connections_available
    ):
        findings += _connection_findings(baseline, current)
    return findings


def external_connection_findings(workbook: WorkbookSnapshot) -> list[Finding]:
    """Presence and risk, kept separate from content comparison."""
    return [
        Finding(
            artifact="excel",
            finding_class=FindingClass.EXTERNAL_CONNECTION,
            element=connection.name,
            current_value=connection.kind,
            message=(
                f"data connection {connection.name!r} reaches an external "
                f"{connection.kind} target"
            ),
        )
        for connection in workbook.metadata.connections
        if connection.external
    ]


def _coverage(
    check_id: str,
    label: str,
    workbooks: tuple[WorkbookSnapshot, ...],
    *,
    available: str,
    detail_field: str,
    counted: str,
) -> CoverageItem:
    blocked = [
        book for book in workbooks if not getattr(book.metadata, available)
    ]
    degraded = [
        book
        for book in workbooks
        if getattr(book.metadata, available) and getattr(book.metadata, detail_field)
    ]
    if blocked:
        state = CoverageState.UNAVAILABLE
        source = blocked
    elif degraded:
        state = CoverageState.DEGRADED
        source = degraded
    else:
        state = CoverageState.CHECKED
        source = []
    return CoverageItem(
        check_id=check_id,
        label=label,
        artifact="excel",
        state=state,
        findings=sum(len(getattr(book.metadata, counted)) for book in workbooks),
        detail="; ".join(
            f"{book.source_name}: {getattr(book.metadata, detail_field)}"
            for book in source
        ),
    )


def comment_coverage(*workbooks: WorkbookSnapshot) -> CoverageItem:
    return _coverage(
        "excel-comments",
        "Cell comments and notes",
        workbooks,
        available="comments_available",
        detail_field="comments_detail",
        counted="comments",
    )


def power_query_coverage(*workbooks: WorkbookSnapshot) -> CoverageItem:
    return _coverage(
        "excel-power-query",
        "Power Query definitions",
        workbooks,
        available="queries_available",
        detail_field="queries_detail",
        counted="queries",
    )


def connection_coverage(*workbooks: WorkbookSnapshot) -> CoverageItem:
    return _coverage(
        "excel-connections",
        "Workbook data connections",
        workbooks,
        available="connections_available",
        detail_field="connections_detail",
        counted="connections",
    )
