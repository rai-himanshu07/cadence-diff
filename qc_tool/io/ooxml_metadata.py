"""Cell comments, Power Query definitions, and workbook connections.

Read straight from the package. Nothing here opens a network connection, follows
an external reference, or evaluates a query: connections are described by a
fixed classification vocabulary and a digest, never by the connection string,
URL, or command text, so a server name or credential cannot reach a finding, a
report, or a log line.
"""

from __future__ import annotations

import hashlib
import io
import re
import zipfile
from dataclasses import dataclass
from xml.etree import ElementTree

from qc_tool.io.opc import (
    UnsafeRelationshipTargetError,
    resolve_internal_relationship_target,
)

_WORKBOOK_PART = "xl/workbook.xml"
_CONNECTIONS_PART = "xl/connections.xml"
#: A comment is analyst-prose like a cell value, but it is bounded in evidence.
_MAX_COMMENT_CHARS = 240
_CONNECTION_KINDS = {
    "1": "odbc",
    "2": "dao",
    "3": "file",
    "4": "web",
    "5": "oledb",
    "6": "text",
    "7": "ado",
    "8": "dsp",
}
_SHARED_QUERY_RE = re.compile(
    r'^\s*shared\s+(?P<name>"(?:[^"]|"")*"|[^\s=]+)\s*=', re.MULTILINE
)


@dataclass(frozen=True, slots=True)
class CellComment:
    sheet: str
    ref: str
    author: str
    text: str
    threaded: bool = False

    @property
    def location(self) -> str:
        return f"{self.sheet}!{self.ref}"


@dataclass(frozen=True, slots=True)
class PowerQuery:
    name: str
    digest: str
    line_count: int


@dataclass(frozen=True, slots=True)
class WorkbookConnection:
    """A connection described without ever retaining what it points at."""

    name: str
    kind: str
    target_digest: str
    external: bool = False
    refresh_on_load: bool | None = None
    background_refresh: bool | None = None


@dataclass(frozen=True, slots=True)
class WorkbookMetadataScan:
    comments: tuple[CellComment, ...] = ()
    queries: tuple[PowerQuery, ...] = ()
    connections: tuple[WorkbookConnection, ...] = ()
    comments_available: bool = False
    queries_available: bool = False
    connections_available: bool = False
    comments_detail: str = ""
    queries_detail: str = ""
    connections_detail: str = ""


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _attribute(element: ElementTree.Element, name: str) -> str | None:
    for key, value in element.attrib.items():
        if _local_name(key) == name:
            return value
    return None


def _digest(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="replace")).hexdigest()[:16]


def _xml_bool(value: str | None) -> bool | None:
    if value is None:
        return None
    return value.strip().lower() in {"1", "true", "on"}


def _relationships(archive: zipfile.ZipFile, source_part: str) -> list[tuple[str, str, str]]:
    """Return ``(id, type, resolved target)`` for every internal relationship."""
    source_directory, _, source_name = source_part.rpartition("/")
    part = "/".join(
        item for item in (source_directory, "_rels", f"{source_name}.rels") if item
    )
    if part not in archive.namelist():
        return []
    try:
        root = ElementTree.fromstring(archive.read(part))
    except ElementTree.ParseError:
        return []
    found: list[tuple[str, str, str]] = []
    for relationship in root:
        if _local_name(relationship.tag) != "Relationship":
            continue
        identifier = relationship.get("Id")
        target = relationship.get("Target")
        if not identifier or not target or relationship.get("TargetMode") == "External":
            continue
        resolved = resolve_internal_relationship_target(source_part, target)
        kind = (relationship.get("Type") or "").rsplit("/", 1)[-1]
        found.append((identifier, kind, resolved))
    return found


def _relationship_targets(archive: zipfile.ZipFile, source_part: str) -> dict[str, str]:
    return {
        identifier: target
        for identifier, _, target in _relationships(archive, source_part)
    }


def _sheet_parts(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    if _WORKBOOK_PART not in archive.namelist():
        return []
    try:
        root = ElementTree.fromstring(archive.read(_WORKBOOK_PART))
    except ElementTree.ParseError:
        return []
    relationships = _relationship_targets(archive, _WORKBOOK_PART)
    pairs: list[tuple[str, str]] = []
    for element in root.iter():
        if _local_name(element.tag) != "sheet":
            continue
        name = element.get("name")
        identifier = _attribute(element, "id")
        if name and identifier and identifier in relationships:
            pairs.append((name, relationships[identifier]))
    return pairs


# --- comments ------------------------------------------------------------


def _element_text(element: ElementTree.Element) -> str:
    return "".join(element.itertext()).strip()[:_MAX_COMMENT_CHARS]


def _legacy_comments(sheet: str, payload: bytes) -> list[CellComment]:
    root = ElementTree.fromstring(payload)
    authors: list[str] = []
    comments: list[CellComment] = []
    for block in root:
        tag = _local_name(block.tag)
        if tag == "authors":
            authors = [(item.text or "").strip() for item in block]
        elif tag == "commentList":
            for comment in block:
                ref = comment.get("ref") or ""
                index = comment.get("authorId")
                author = ""
                if index is not None and index.isdigit() and int(index) < len(authors):
                    author = authors[int(index)]
                comments.append(
                    CellComment(
                        sheet=sheet,
                        ref=ref,
                        author=author,
                        text=_element_text(comment),
                    )
                )
    return comments


def _threaded_comments(sheet: str, payload: bytes) -> list[CellComment]:
    root = ElementTree.fromstring(payload)
    return [
        CellComment(
            sheet=sheet,
            ref=comment.get("ref") or "",
            author=comment.get("personId") or "",
            text=_element_text(comment),
            threaded=True,
        )
        for comment in root
        if _local_name(comment.tag) == "threadedComment"
    ]


def _scan_comments(archive: zipfile.ZipFile) -> tuple[list[CellComment], bool, str]:
    comments: list[CellComment] = []
    failures = 0
    for sheet, part in _sheet_parts(archive):
        for _, kind, target in _relationships(archive, part):
            # Producers disagree on part names, so the relationship type decides.
            if kind not in {"comments", "threadedComment"}:
                continue
            try:
                payload = archive.read(target)
                if kind == "threadedComment":
                    comments.extend(_threaded_comments(sheet, payload))
                else:
                    comments.extend(_legacy_comments(sheet, payload))
            except (KeyError, ElementTree.ParseError):
                failures += 1
    detail = f"{failures} comment parts could not be read" if failures else ""
    return comments, failures == 0, detail


# --- power query ---------------------------------------------------------


def _mashup_blobs(archive: zipfile.ZipFile) -> list[str]:
    blobs: list[str] = []
    for name in archive.namelist():
        if not name.casefold().startswith("customxml/item"):
            continue
        try:
            root = ElementTree.fromstring(archive.read(name))
        except (KeyError, ElementTree.ParseError):
            continue
        if _local_name(root.tag) == "DataMashup" and root.text:
            blobs.append(root.text.strip())
    return blobs


def _section_queries(text: str) -> list[PowerQuery]:
    matches = list(_SHARED_QUERY_RE.finditer(text))
    if not matches:
        return [
            PowerQuery(
                name="Section1",
                digest=_digest(text),
                line_count=len(text.splitlines()),
            )
        ]
    queries: list[PowerQuery] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        body = text[match.start() : end]
        name = match.group("name").strip('"').replace('""', '"')
        queries.append(
            PowerQuery(name=name, digest=_digest(body), line_count=len(body.splitlines()))
        )
    return queries


def _scan_queries(archive: zipfile.ZipFile) -> tuple[list[PowerQuery], bool, str]:
    import base64
    import binascii

    blobs = _mashup_blobs(archive)
    if not blobs:
        return [], True, ""
    queries: list[PowerQuery] = []
    decoded_sections = 0
    for index, blob in enumerate(blobs, start=1):
        try:
            payload = base64.b64decode(blob, validate=True)
        except (binascii.Error, ValueError):
            queries.append(PowerQuery(f"DataMashup{index}", _digest(blob), 0))
            continue
        section = _mashup_section(payload)
        if section is None:
            # The mashup envelope could not be opened; the blob digest still
            # detects that the query definitions changed.
            queries.append(
                PowerQuery(f"DataMashup{index}", hashlib.sha256(payload).hexdigest()[:16], 0)
            )
            continue
        decoded_sections += 1
        queries.extend(_section_queries(section))
    if decoded_sections == len(blobs):
        return queries, True, ""
    return (
        queries,
        True,
        "query text could not be opened for "
        f"{len(blobs) - decoded_sections} of {len(blobs)} mashups; "
        "comparison falls back to a definition digest",
    )


def _mashup_section(payload: bytes) -> str | None:
    """Open the mashup package parts and return the M section document."""
    if len(payload) < 8:
        return None
    length = int.from_bytes(payload[4:8], "little")
    if not 0 < length <= len(payload) - 8:
        return None
    try:
        package = zipfile.ZipFile(io.BytesIO(payload[8 : 8 + length]))
    except zipfile.BadZipFile:
        return None
    with package:
        part = next(
            (name for name in package.namelist() if name.casefold().endswith("section1.m")),
            None,
        )
        if part is None:
            return None
        return package.read(part).decode("utf-8", errors="replace")


# --- connections ---------------------------------------------------------


def _scan_connections(
    archive: zipfile.ZipFile,
) -> tuple[list[WorkbookConnection], bool, str]:
    if _CONNECTIONS_PART not in archive.namelist():
        return [], True, ""
    try:
        root = ElementTree.fromstring(archive.read(_CONNECTIONS_PART))
    except ElementTree.ParseError as exc:
        return [], False, f"connection metadata is not readable XML: {exc}"

    connections: list[WorkbookConnection] = []
    for element in root:
        if _local_name(element.tag) != "connection":
            continue
        targets: list[str] = []
        kind = _CONNECTION_KINDS.get(element.get("type") or "", "unknown")
        for child in element:
            child_tag = _local_name(child.tag)
            if child_tag == "dbPr":
                targets.extend(
                    value
                    for value in (child.get("connection"), child.get("command"))
                    if value
                )
            elif child_tag == "webPr":
                targets.extend(value for value in (child.get("url"),) if value)
                kind = "web"
            elif child_tag == "textPr":
                kind = "text"
            elif child_tag == "olapPr":
                kind = "olap"
        joined = "\u0000".join(targets)
        lowered = joined.casefold()
        if "microsoft.mashup" in lowered:
            kind = "power_query"
        connections.append(
            WorkbookConnection(
                name=element.get("name") or f"connection{element.get('id') or ''}",
                kind=kind,
                target_digest=_digest(joined),
                external=any(
                    marker in lowered
                    for marker in ("http://", "https://", "ftp://", "\\\\")
                ),
                refresh_on_load=_xml_bool(element.get("refreshOnLoad")),
                background_refresh=_xml_bool(element.get("background")),
            )
        )
    return connections, True, ""


def scan_workbook_metadata(data: bytes) -> WorkbookMetadataScan:
    """Read comments, Power Query definitions, and connections from a package."""
    try:
        archive = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        detail = "workbook is not an OOXML package"
        return WorkbookMetadataScan(
            comments_detail=detail, queries_detail=detail, connections_detail=detail
        )
    with archive:
        try:
            comments, comments_ok, comments_detail = _scan_comments(archive)
            queries, queries_ok, queries_detail = _scan_queries(archive)
            connections, connections_ok, connections_detail = _scan_connections(archive)
        except UnsafeRelationshipTargetError:
            detail = "unsafe relationship target in workbook package"
            return WorkbookMetadataScan(
                comments_detail=detail,
                queries_detail=detail,
                connections_detail=detail,
            )
    return WorkbookMetadataScan(
        comments=tuple(sorted(comments, key=lambda item: (item.sheet, item.ref))),
        queries=tuple(sorted(queries, key=lambda item: item.name)),
        connections=tuple(sorted(connections, key=lambda item: item.name)),
        comments_available=comments_ok,
        queries_available=queries_ok,
        connections_available=connections_ok,
        comments_detail=comments_detail,
        queries_detail=queries_detail,
        connections_detail=connections_detail,
    )
