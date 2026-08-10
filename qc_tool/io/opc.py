"""Shared, bounded contracts for Office Open Packaging Convention parts."""

from __future__ import annotations

import io
import posixpath
import zipfile
from urllib.parse import urlsplit
from xml.etree import ElementTree


class InvalidOfficePackageError(ValueError):
    """An input has an Office extension/signature but is not a readable package."""


class UnsafeRelationshipTargetError(ValueError):
    """An internal relationship target escapes the package root."""


def validate_office_package(data: bytes, *, source_name: str) -> None:
    """Validate the ZIP directory without expanding package members."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            archive.infolist()
            names = set(archive.namelist())
            suffix = source_name.casefold().rsplit(".", 1)[-1]
            required = {"[Content_Types].xml", "_rels/.rels"}
            if suffix in {"xlsx", "xlsm"}:
                required.add("xl/workbook.xml")
            elif suffix == "xlsb":
                required.add("xl/workbook.bin")
            elif suffix == "pptx":
                required.add("ppt/presentation.xml")
            if not required <= names:
                raise InvalidOfficePackageError(
                    f"{source_name}: invalid Office package; required metadata is missing"
                )
            for relationship_part in (
                name for name in names if name.endswith(".rels")
            ):
                try:
                    root = ElementTree.fromstring(archive.read(relationship_part))
                except ElementTree.ParseError as exc:
                    raise InvalidOfficePackageError(
                        f"{source_name}: invalid Office package relationship metadata"
                    ) from exc
                source_part = _source_part_for_relationships(relationship_part)
                for relationship in root:
                    target = relationship.get("Target")
                    if not target or relationship.get("TargetMode") == "External":
                        continue
                    resolve_internal_relationship_target(source_part, target)
    except UnsafeRelationshipTargetError as exc:
        raise InvalidOfficePackageError(
            f"{source_name}: invalid Office package relationship target"
        ) from exc
    except (OSError, zipfile.BadZipFile) as exc:
        raise InvalidOfficePackageError(
            f"{source_name}: invalid Office package"
        ) from exc


def _source_part_for_relationships(relationship_part: str) -> str:
    if relationship_part == "_rels/.rels":
        return ""
    directory, separator, filename = relationship_part.rpartition("/_rels/")
    if not separator or not filename.endswith(".rels"):
        return relationship_part.removesuffix(".rels")
    source_name = filename.removesuffix(".rels")
    return f"{directory}/{source_name}" if directory else source_name


def resolve_internal_relationship_target(source_part: str, target: str) -> str:
    """Resolve one internal OPC target and reject package-root traversal."""
    normalized_target = target.strip()
    if not normalized_target or "\\" in normalized_target:
        raise UnsafeRelationshipTargetError("unsafe relationship target")
    parsed = urlsplit(normalized_target)
    if parsed.scheme or parsed.netloc:
        raise UnsafeRelationshipTargetError("unsafe relationship target")
    resolved = (
        posixpath.normpath(normalized_target.lstrip("/"))
        if normalized_target.startswith("/")
        else posixpath.normpath(
            posixpath.join(posixpath.dirname(source_part), normalized_target)
        )
    )
    if resolved in {"", ".", ".."} or resolved.startswith("../"):
        raise UnsafeRelationshipTargetError("unsafe relationship target")
    return resolved
