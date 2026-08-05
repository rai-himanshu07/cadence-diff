"""Fail-closed focus-risk scan for an already-open Office package.

Focus changes Office selection, which can invoke event handlers, add-ins, and
embedded content. A document is eligible only when its saved package contains
none of the risk vocabulary below. Anything ambiguous, malformed, or
unrecognised refuses.

The Excel part/relationship vocabulary is the same one the XLSB BIFF12 scanner
and the OOXML loader already use; ``tests/test_focus_binding.py`` pins that
equivalence. PowerPoint adds slide actions and OLE packages.
"""

from __future__ import annotations

import posixpath
import zipfile
from collections.abc import Iterable
from dataclasses import dataclass
from enum import StrEnum
from pathlib import Path
from xml.etree import ElementTree

#: Focus never scans an unbounded package; larger sources refuse instead.
MAX_PACKAGE_BYTES = 512 * 1024 * 1024

#: Only these formats may be focused. Macro-enabled formats are never eligible.
SUPPORTED_SUFFIXES = frozenset({".xlsx", ".xlsb", ".pptx"})
ACTIVE_CONTENT_SUFFIXES = frozenset({".xlsm", ".pptm", ".xls", ".ppt", ".xlam", ".ppam"})


class PackageRiskKind(StrEnum):
    VBA_PROJECT = "vba_project"
    ACTIVEX_CONTROL = "activex_control"
    EMBEDDED_OLE = "embedded_ole"
    CUSTOM_OFFICE_UI = "custom_office_ui"
    EXTERNAL_DATA = "external_data"
    EXTERNAL_RELATIONSHIP = "external_relationship"
    MACRO_SHEET = "macro_sheet"
    DIALOG_SHEET = "dialog_sheet"
    QUERY_TABLE = "query_table"
    CONTROL_CONTENT = "control_content"
    SLIDE_ACTION = "slide_action"
    UNKNOWN_EXECUTABLE_RELATIONSHIP = "unknown_executable_relationship"
    PATH_TRAVERSAL = "path_traversal"
    UNREADABLE_RELATIONSHIP_METADATA = "unreadable_relationship_metadata"


class PackageScanError(StrEnum):
    """Fixed refusal codes; never a message, path, or exception text."""

    UNSUPPORTED_FORMAT = "unsupported_format"
    ACTIVE_CONTENT_FORMAT = "active_content_format"
    ENCRYPTED_PACKAGE = "encrypted_package"
    MALFORMED_PACKAGE = "malformed_package"
    PACKAGE_TOO_LARGE = "package_too_large"
    UNREADABLE_PACKAGE = "unreadable_package"


_RISKY_PARTS = {
    "xl/vbaproject.bin": PackageRiskKind.VBA_PROJECT,
    "ppt/vbaproject.bin": PackageRiskKind.VBA_PROJECT,
    "xl/connections.xml": PackageRiskKind.EXTERNAL_DATA,
    "xl/connections.bin": PackageRiskKind.EXTERNAL_DATA,
}
_RISKY_PATH_SEGMENTS = {
    "/activex/": PackageRiskKind.ACTIVEX_CONTROL,
    "/ctrlprops/": PackageRiskKind.CONTROL_CONTENT,
    "/dialogsheets/": PackageRiskKind.DIALOG_SHEET,
    "/embeddings/": PackageRiskKind.EMBEDDED_OLE,
    "/externalconnections/": PackageRiskKind.EXTERNAL_DATA,
    "/externallinks/": PackageRiskKind.EXTERNAL_DATA,
    "/macrosheets/": PackageRiskKind.MACRO_SHEET,
    "/querytables/": PackageRiskKind.QUERY_TABLE,
    "/customui/": PackageRiskKind.CUSTOM_OFFICE_UI,
    "/vbaproject.bin": PackageRiskKind.VBA_PROJECT,
}
_RISKY_RELATIONSHIP_KINDS = {
    "activexcontrol": PackageRiskKind.ACTIVEX_CONTROL,
    "attachedtoolbars": PackageRiskKind.CUSTOM_OFFICE_UI,
    "connections": PackageRiskKind.EXTERNAL_DATA,
    "ctrlprop": PackageRiskKind.CONTROL_CONTENT,
    "customui": PackageRiskKind.CUSTOM_OFFICE_UI,
    "dialogsheet": PackageRiskKind.DIALOG_SHEET,
    "externallink": PackageRiskKind.EXTERNAL_DATA,
    "externallinkpath": PackageRiskKind.EXTERNAL_DATA,
    "macrosheet": PackageRiskKind.MACRO_SHEET,
    "oleobject": PackageRiskKind.EMBEDDED_OLE,
    "package": PackageRiskKind.EMBEDDED_OLE,
    "querytable": PackageRiskKind.QUERY_TABLE,
    "vbaproject": PackageRiskKind.VBA_PROJECT,
}
#: PowerPoint click/hover actions can run a macro or launch a program.
_ACTION_ELEMENTS = frozenset({"hlinkClick", "hlinkHover"})
_ACTION_ATTRIBUTES = frozenset({"action"})
_SAFE_ACTION_PREFIXES = ("ppaction://hlinkshowjump", "ppaction://hlinksldjump")

_EXECUTABLE_SUFFIXES = frozenset(
    {
        ".bat", ".cmd", ".com", ".cpl", ".dll", ".exe", ".hta", ".jar", ".js",
        ".jse", ".lnk", ".msc", ".msi", ".pif", ".ps1", ".reg", ".scr", ".sct",
        ".vb", ".vbe", ".vbs", ".wsf", ".wsh",
    }
)
#: Compound-file magic; an encrypted OOXML package is an OLE container.
_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


@dataclass(frozen=True, slots=True)
class PackageScan:
    """Aggregate-only scan result: fixed codes, never names or targets."""

    risks: tuple[PackageRiskKind, ...] = ()
    error: PackageScanError | None = None

    @property
    def safe_to_focus(self) -> bool:
        return self.error is None and not self.risks


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _escapes_package(target: str) -> bool:
    text = target.replace("\\", "/").strip()
    if not text:
        return True
    if text.startswith("//") or "://" in text:
        return True
    if text.startswith("/"):
        text = text.lstrip("/")
    if len(text) >= 2 and text[1] == ":":
        return True
    return posixpath.normpath(text).startswith("..")


def _relationship_source_part(relationship_part: str) -> str | None:
    normalized = relationship_part.replace("\\", "/").lstrip("/")
    if normalized == "_rels/.rels":
        return ""
    marker = "/_rels/"
    if marker not in normalized:
        return None
    parent, relationship_name = normalized.rsplit(marker, 1)
    if not relationship_name.endswith(".rels") or relationship_name == ".rels":
        return None
    return posixpath.join(parent, relationship_name.removesuffix(".rels"))


def _relationship_target_escapes(relationship_part: str, target: str) -> bool:
    text = target.replace("\\", "/").strip()
    source_part = _relationship_source_part(relationship_part)
    if not text or source_part is None:
        return True
    if text.startswith("//") or "://" in text:
        return True
    if len(text) >= 2 and text[1] == ":":
        return True
    path = text.split("#", 1)[0].split("?", 1)[0]
    if path.startswith("/"):
        resolved = posixpath.normpath(path.lstrip("/"))
    else:
        resolved = posixpath.normpath(
            posixpath.join(posixpath.dirname(source_part), path)
        )
    return resolved in {"", ".", ".."} or resolved.startswith("../")


def _relationship_risks(archive: zipfile.ZipFile) -> set[PackageRiskKind]:
    risks: set[PackageRiskKind] = set()
    for part in archive.namelist():
        if not part.lower().endswith(".rels"):
            continue
        try:
            root = ElementTree.fromstring(archive.read(part))
        except (ElementTree.ParseError, KeyError, OSError, zipfile.BadZipFile):
            risks.add(PackageRiskKind.UNREADABLE_RELATIONSHIP_METADATA)
            continue
        for element in root:
            if _local_name(element.tag) != "Relationship":
                continue
            kind = element.get("Type", "").rsplit("/", 1)[-1].lower()
            risky = _RISKY_RELATIONSHIP_KINDS.get(kind)
            if risky is not None:
                risks.add(risky)
            target = element.get("Target", "")
            external = element.get("TargetMode") == "External"
            if external and kind != "hyperlink":
                risks.add(PackageRiskKind.EXTERNAL_RELATIONSHIP)
            suffix = posixpath.splitext(target.split("?", 1)[0])[1].lower()
            if suffix in _EXECUTABLE_SUFFIXES:
                risks.add(PackageRiskKind.UNKNOWN_EXECUTABLE_RELATIONSHIP)
            if not external and _relationship_target_escapes(part, target):
                risks.add(PackageRiskKind.PATH_TRAVERSAL)
    return risks


def _slide_action_risks(archive: zipfile.ZipFile) -> set[PackageRiskKind]:
    risks: set[PackageRiskKind] = set()
    for part in archive.namelist():
        lowered = part.lower()
        if not lowered.endswith(".xml"):
            continue
        if not (
            lowered.startswith("ppt/slides/")
            or lowered.startswith("ppt/slidemasters/")
            or lowered.startswith("ppt/slidelayouts/")
        ):
            continue
        try:
            root = ElementTree.fromstring(archive.read(part))
        except (ElementTree.ParseError, KeyError, OSError, zipfile.BadZipFile):
            risks.add(PackageRiskKind.UNREADABLE_RELATIONSHIP_METADATA)
            continue
        for element in root.iter():
            if _local_name(element.tag) not in _ACTION_ELEMENTS:
                continue
            for name, value in element.attrib.items():
                if _local_name(name) not in _ACTION_ATTRIBUTES:
                    continue
                action = value.strip().lower()
                if action and not action.startswith(_SAFE_ACTION_PREFIXES):
                    risks.add(PackageRiskKind.SLIDE_ACTION)
    return risks


def _archive_risks(archive: zipfile.ZipFile) -> set[PackageRiskKind]:
    risks: set[PackageRiskKind] = set()
    names = archive.namelist()
    if "[Content_Types].xml" not in names:
        risks.add(PackageRiskKind.PATH_TRAVERSAL)
    for name in names:
        lowered = "/" + name.replace("\\", "/").lower().lstrip("/")
        exact = lowered.lstrip("/")
        if _escapes_package(name):
            risks.add(PackageRiskKind.PATH_TRAVERSAL)
        risky = _RISKY_PARTS.get(exact)
        if risky is not None:
            risks.add(risky)
        if exact.endswith(".rels") or "/_rels/" in lowered:
            continue
        for segment, kind in _RISKY_PATH_SEGMENTS.items():
            if segment in lowered:
                risks.add(kind)
    return risks


def scan_focus_package(path: Path) -> PackageScan:
    """Scan one saved Office package; anything unproved refuses focus."""
    suffix = path.suffix.lower()
    if suffix in ACTIVE_CONTENT_SUFFIXES:
        return PackageScan(error=PackageScanError.ACTIVE_CONTENT_FORMAT)
    if suffix not in SUPPORTED_SUFFIXES:
        return PackageScan(error=PackageScanError.UNSUPPORTED_FORMAT)
    try:
        if path.stat().st_size > MAX_PACKAGE_BYTES:
            return PackageScan(error=PackageScanError.PACKAGE_TOO_LARGE)
        with path.open("rb") as handle:
            if handle.read(len(_OLE_MAGIC)) == _OLE_MAGIC:
                return PackageScan(error=PackageScanError.ENCRYPTED_PACKAGE)
    except OSError:
        return PackageScan(error=PackageScanError.UNREADABLE_PACKAGE)
    risks: set[PackageRiskKind] = set()
    try:
        with zipfile.ZipFile(path) as archive:
            risks |= _archive_risks(archive)
            risks |= _relationship_risks(archive)
            if suffix == ".pptx":
                risks |= _slide_action_risks(archive)
    except zipfile.BadZipFile:
        return PackageScan(error=PackageScanError.MALFORMED_PACKAGE)
    except OSError:
        return PackageScan(error=PackageScanError.UNREADABLE_PACKAGE)
    return PackageScan(risks=tuple(sorted(risks, key=lambda item: item.value)))


def risk_labels(risks: Iterable[PackageRiskKind]) -> tuple[str, ...]:
    return tuple(sorted(risk.value for risk in risks))
