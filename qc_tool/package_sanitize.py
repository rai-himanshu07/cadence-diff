"""Package-aware strict sanitization for a current Excel/PPT deliverable pair."""

import dataclasses
import datetime as dt
import hashlib
from pathlib import Path
from typing import Any, cast

from pptx import Presentation
from pydantic import BaseModel, Field

from qc_tool.config.profile import CrosscheckMapping, DeliverableProfile
from qc_tool.crosscheck.numbers import format_figure_like, replace_figure_ordinals
from qc_tool.crosscheck.trace import (
    extract_deck_figures,
    mapping_identity,
    occurrence_identity,
    verify_mappings,
)
from qc_tool.io.loader import load_workbook_snapshot
from qc_tool.ppt.extract import load_deck_snapshot
from qc_tool.privacy import PrivacyReport, verify_sanitized
from qc_tool.sanitize import SanitizeError, sanitize_file
from qc_tool.security import private_directory, private_file


class PackageMappingResult(BaseModel):
    mapping_digest: str
    source_sheet_ordinal: int | None = None
    source_cell: str
    occurrence_ordinal: int | None = None
    status: str
    detail: str = ""


class PackageSanitizeManifest(BaseModel):
    schema_version: int = 1
    generated_at: str
    seed: int
    privacy_safe: bool
    mappings_verified: int = 0
    mappings_unverifiable: int = 0
    inputs: dict[str, dict[str, str | int]]
    outputs: dict[str, dict[str, str | int]]
    mapping_results: list[PackageMappingResult] = Field(default_factory=list)
    sanitizer_stats: dict[str, dict] = Field(default_factory=dict)
    privacy: dict[str, PrivacyReport]


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _mapping_digest(mapping: CrosscheckMapping) -> str:
    payload = (
        mapping.slide,
        mapping.line_skeleton,
        mapping.figure_index,
        mapping.source_sheet,
        mapping.source_cell,
    )
    return hashlib.sha256(repr(payload).encode()).hexdigest()[:20]


def _replace_deck_figures(path: Path, replacements: dict[int, str]) -> None:
    presentation = Presentation(str(path))
    figure_index = 0
    for slide in presentation.slides:
        title_shape = slide.shapes.title
        for base_shape in slide.shapes:
            shape = cast(Any, base_shape)
            is_title = bool(
                title_shape is not None and base_shape._element is title_shape._element
            )
            if base_shape.has_text_frame and not is_title:
                for paragraph in shape.text_frame.paragraphs:
                    if not paragraph.text.strip():
                        continue
                    updated, figure_index = replace_figure_ordinals(
                        paragraph.text, replacements, start_index=figure_index
                    )
                    if updated != paragraph.text:
                        paragraph.text = updated
            if base_shape.has_table:
                for row in list(shape.table.rows)[1:]:
                    for cell in list(row.cells)[1:]:
                        updated, figure_index = replace_figure_ordinals(
                            cell.text, replacements, start_index=figure_index
                        )
                        if updated != cell.text:
                            cell.text = updated
    presentation.save(str(path))
    private_file(path)


def sanitize_package(
    excel_source: Path,
    ppt_source: Path,
    output_dir: Path,
    profile: DeliverableProfile,
    *,
    seed: int = 0,
    forbidden_tokens: list[str] | None = None,
) -> PackageSanitizeManifest:
    """Strictly sanitize a pair and reconcile all evaluable confirmed mappings."""
    output_dir = private_directory(output_dir)
    excel_output = output_dir / "workbook.sanitized.xlsx"
    ppt_output = output_dir / "deck.sanitized.pptx"
    manifest_path = output_dir / "redaction-manifest.json"

    original_workbook = load_workbook_snapshot(excel_source)
    original_deck = load_deck_snapshot(ppt_source)
    original_occurrences = extract_deck_figures(original_deck)
    occurrence_indices = {
        occurrence_identity(occurrence): index
        for index, occurrence in enumerate(original_occurrences)
    }

    _, excel_stats = sanitize_file(
        excel_source, excel_output, seed=seed, redact_text=True
    )
    _, ppt_stats = sanitize_file(ppt_source, ppt_output, seed=seed, redact_text=True)
    sanitized_workbook = load_workbook_snapshot(excel_output)

    replacements: dict[int, str] = {}
    mapping_results: list[PackageMappingResult] = []
    pending: list[tuple[CrosscheckMapping, int, int, str]] = []
    for mapping in profile.crosscheck.mappings:
        digest = _mapping_digest(mapping)
        ordinal = occurrence_indices.get(mapping_identity(mapping))
        try:
            sheet_ordinal = original_workbook.sheet_names.index(mapping.source_sheet) + 1
        except ValueError:
            sheet_ordinal = None
        if ordinal is None or sheet_ordinal is None:
            mapping_results.append(
                PackageMappingResult(
                    mapping_digest=digest,
                    source_sheet_ordinal=sheet_ordinal,
                    source_cell=mapping.source_cell,
                    occurrence_ordinal=ordinal,
                    status="unverifiable",
                    detail="mapping anchor or source sheet does not resolve",
                )
            )
            continue
        source_sheet = f"Sheet_{sheet_ordinal:03d}"
        cell = sanitized_workbook.sheet(source_sheet).cell(mapping.source_cell)
        value = None if cell is None else cell.value
        if isinstance(value, bool) or not isinstance(value, int | float):
            mapping_results.append(
                PackageMappingResult(
                    mapping_digest=digest,
                    source_sheet_ordinal=sheet_ordinal,
                    source_cell=mapping.source_cell,
                    occurrence_ordinal=ordinal,
                    status="unverifiable",
                    detail="sanitized source cell has no numeric cached value",
                )
            )
            continue
        replacements[ordinal] = format_figure_like(
            original_occurrences[ordinal].figure, float(value)
        )
        pending.append((mapping, ordinal, sheet_ordinal, digest))

    if replacements:
        _replace_deck_figures(ppt_output, replacements)

    sanitized_deck = load_deck_snapshot(ppt_output)
    sanitized_occurrences = extract_deck_figures(sanitized_deck)
    sanitized_mappings: list[CrosscheckMapping] = []
    pending_by_digest: dict[str, tuple[int, int]] = {}
    for original_mapping, ordinal, sheet_ordinal, digest in pending:
        if ordinal >= len(sanitized_occurrences):
            mapping_results.append(
                PackageMappingResult(
                    mapping_digest=digest,
                    source_sheet_ordinal=sheet_ordinal,
                    source_cell=original_mapping.source_cell,
                    occurrence_ordinal=ordinal,
                    status="unverifiable",
                    detail="sanitized deck occurrence ordering changed",
                )
            )
            continue
        occurrence = sanitized_occurrences[ordinal]
        sanitized_mappings.append(
            CrosscheckMapping(
                slide=occurrence.slide,
                line_skeleton=occurrence.line_skeleton,
                figure_index=occurrence.figure_index,
                label=f"mapping_{digest}",
                source_sheet=f"Sheet_{sheet_ordinal:03d}",
                source_cell=original_mapping.source_cell,
            )
        )
        pending_by_digest[digest] = (ordinal, sheet_ordinal)

    sanitized_profile = profile.crosscheck.model_copy(
        update={"mappings": sanitized_mappings}
    )
    verification = verify_mappings(sanitized_deck, sanitized_workbook, sanitized_profile)
    verified_labels = {mapping.label for mapping in verification.verified}
    findings_by_label = {finding.element: finding for finding in verification.findings}
    for mapping in sanitized_mappings:
        digest = mapping.label.removeprefix("mapping_")
        ordinal, sheet_ordinal = pending_by_digest[digest]
        finding = findings_by_label.get(mapping.label)
        verified = mapping.label in verified_labels
        mapping_results.append(
            PackageMappingResult(
                mapping_digest=digest,
                source_sheet_ordinal=sheet_ordinal,
                source_cell=mapping.source_cell,
                occurrence_ordinal=ordinal,
                status="verified" if verified else "unverifiable",
                detail="" if verified else (finding.message if finding else "verification failed"),
            )
        )

    tokens = forbidden_tokens or []
    privacy = {
        "excel": verify_sanitized(excel_output, forbidden_tokens=tokens),
        "ppt": verify_sanitized(ppt_output, forbidden_tokens=tokens),
    }
    privacy_safe = all(report.safe for report in privacy.values())
    verified_count = sum(item.status == "verified" for item in mapping_results)
    unverifiable_count = sum(item.status != "verified" for item in mapping_results)
    manifest = PackageSanitizeManifest(
        generated_at=dt.datetime.now(dt.UTC).isoformat(timespec="seconds"),
        seed=seed,
        privacy_safe=privacy_safe,
        mappings_verified=verified_count,
        mappings_unverifiable=unverifiable_count,
        inputs={
            "excel": {"sha256": _sha256(excel_source), "size": excel_source.stat().st_size},
            "ppt": {"sha256": _sha256(ppt_source), "size": ppt_source.stat().st_size},
        },
        outputs={
            "excel": {"sha256": _sha256(excel_output), "size": excel_output.stat().st_size},
            "ppt": {"sha256": _sha256(ppt_output), "size": ppt_output.stat().st_size},
        },
        mapping_results=sorted(mapping_results, key=lambda item: item.mapping_digest),
        sanitizer_stats={
            "excel": dataclasses.asdict(excel_stats),
            "ppt": dataclasses.asdict(ppt_stats),
        },
        privacy=privacy,
    )
    manifest_path.write_text(manifest.model_dump_json(indent=2), encoding="utf-8")
    private_file(manifest_path)
    if not privacy_safe:
        raise SanitizeError("package privacy verification failed; see redaction-manifest.json")
    return manifest
