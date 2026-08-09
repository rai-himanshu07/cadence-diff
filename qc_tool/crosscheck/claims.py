"""Claim identity and the eligible/unavailable populations behind coverage.

Coverage is only credible when the population it covers is stated exactly. A
figure rendered into a picture is a material presentation claim that cannot be
read at all, so counting only the figures we *can* read would silently shrink
the denominator and make a clean result look complete.

This module keeps the two populations explicit and requires them to reconcile.
It is deliberately independent of the unresolved question of whether any
commercial tool grounds arbitrary deck claims to a nominated workbook: nothing
here asserts closure, only what was and was not inspected.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Literal

from pydantic import BaseModel, Field

from qc_tool.crosscheck.numbers import numeric_skeleton
from qc_tool.crosscheck.trace import FigureOccurrence
from qc_tool.excel.periods import PeriodKind
from qc_tool.ppt.claim_periods import CURRENT_PERIOD_CONTEXT, periods_in_text
from qc_tool.ppt.model import DeckSnapshot

#: Shape types whose content is rasterized or opaque, so any figure inside them
#: is invisible to every reader in this tool.
_UNREADABLE_SHAPE_TYPES = frozenset(
    {
        "PICTURE",
        "LINKED_PICTURE",
        "MEDIA",
        "EMBEDDED_OLE_OBJECT",
        "LINKED_OLE_OBJECT",
        "OLE_OBJECT",
    }
)


@dataclass(frozen=True, slots=True)
class ClaimIdentity:
    """Durable identity of one presentation claim.

    ``slide`` and ``line_skeleton`` are the durable parts: the slide's display
    name survives reordering, and the numeric skeleton survives a figure
    refresh. ``figure_index`` disambiguates repeated figures on one line, which
    is what keeps two equal values on the same line distinguishable.
    """

    slide: str
    line_skeleton: str
    figure_index: int

    @property
    def key(self) -> str:
        return f"{self.slide}|{self.line_skeleton}|{self.figure_index}"


def claim_identity(occurrence: FigureOccurrence) -> ClaimIdentity:
    return ClaimIdentity(
        slide=occurrence.slide,
        line_skeleton=occurrence.line_skeleton,
        figure_index=occurrence.figure_index,
    )


@dataclass(frozen=True, slots=True)
class ClaimPopulation:
    """The readable and unreadable halves of one deck's claim population."""

    readable: int
    unreadable_shapes: int
    unreadable_slides: tuple[int, ...]
    # additive identity availability counts
    identity_available: int = 0
    identity_unavailable: int = 0

    @property
    def total_inspected_surfaces(self) -> int:
        return self.readable + self.unreadable_shapes

    @property
    def complete(self) -> bool:
        """Whether every material surface in the deck could be read."""
        return self.unreadable_shapes == 0

    @property
    def identity_complete(self) -> bool:
        """Whether identities were resolved for every readable occurrence."""
        return self.identity_unavailable == 0 and self.identity_available == self.readable

    def reconciles(self, mapped: int, unmapped: int) -> bool:
        """Whether the mapped split accounts for exactly the readable claims."""
        return mapped + unmapped == self.readable


def claim_population(
    deck: DeckSnapshot,
    occurrences: list[FigureOccurrence],
    assessments: list[ClaimIdentityAssessment] | None = None,
) -> ClaimPopulation:
    """Split a deck into claims that were read and surfaces that could not be."""
    unreadable = 0
    slides: list[int] = []
    for slide in deck.slides:
        opaque = sum(
            1
            for shape in slide.shapes
            if shape.shape_type in _UNREADABLE_SHAPE_TYPES
        )
        if opaque:
            unreadable += opaque
            slides.append(slide.index + 1)
    resolved = (
        assessments
        if assessments is not None
        else classify_claim_identities(deck, occurrences)
    )
    if len(resolved) != len(occurrences):
        raise ValueError("claim identity population does not reconcile")
    identity_available = sum(item.identity is not None for item in resolved)
    identity_unavailable = len(resolved) - identity_available
    return ClaimPopulation(
        readable=len(occurrences),
        unreadable_shapes=unreadable,
        unreadable_slides=tuple(slides),
        identity_available=identity_available,
        identity_unavailable=identity_unavailable,
    )


class SemanticClaimIdentityV1(BaseModel):
    version: Literal[1] = 1
    semantic_skeleton: str = Field(min_length=1)
    period_kind: Literal["month", "week", "quarter", "date"]
    period_key: tuple[int, int, int]
    unit_key: str = Field(min_length=1)
    scale: float = Field(gt=0)
    decimals: int = Field(ge=0)
    surface_kind: Literal["text", "table", "chart"]
    surface_anchor: str = Field(min_length=1)
    occurrence_ordinal: int = Field(default=0, ge=0)

    model_config = {"frozen": True}


@dataclass(frozen=True, slots=True)
class ClaimIdentityAssessment:
    occurrence: FigureOccurrence
    identity: SemanticClaimIdentityV1 | None
    unavailable_reasons: tuple[str, ...]


def classify_claim_identities(
    deck: DeckSnapshot, occurrences: list[FigureOccurrence]
) -> list[ClaimIdentityAssessment]:
    """Classify readable occurrences without guessing missing identity axes."""

    assessments: list[ClaimIdentityAssessment] = []
    slide_by_index = {slide.index + 1: slide for slide in deck.slides}

    def period_candidates(
        texts: list[str],
    ) -> dict[tuple[PeriodKind, tuple[int, int, int]], set[str]]:
        found: dict[tuple[PeriodKind, tuple[int, int, int]], set[str]] = {}
        for text in texts:
            for label, period in periods_in_text(text):
                found.setdefault((period.kind, period.sort_key), set()).add(label)
        return found

    def normalized_anchor(text: str) -> str:
        rendered = text
        for label, _period in periods_in_text(text):
            rendered = re.sub(
                re.escape(label),
                " ",
                rendered,
                flags=re.IGNORECASE,
            )
        return " ".join(numeric_skeleton(rendered).casefold().split())

    for occurrence in occurrences:
        reasons: list[str] = []
        direct_periods = period_candidates(
            [
                occurrence.line,
                occurrence.line_skeleton,
                occurrence.surface_anchor,
            ]
        )
        chosen_period: tuple[PeriodKind, tuple[int, int, int]] | None = None
        if len(direct_periods) == 1:
            chosen_period = next(iter(direct_periods))
        elif len(direct_periods) > 1:
            reasons.append("ambiguous_period")
        else:
            slide = slide_by_index.get(occurrence.slide_index)
            contextual_texts: list[str] = []
            if slide is not None:
                contextual_texts.append(slide.title or "")
                contextual_texts.extend(
                    text
                    for text in slide.texts
                    if CURRENT_PERIOD_CONTEXT.search(text)
                )
            contextual_periods = period_candidates(contextual_texts)
            if len(contextual_periods) == 1:
                chosen_period = next(iter(contextual_periods))
            elif len(contextual_periods) > 1:
                reasons.append("ambiguous_period")
            else:
                reasons.append("period_unavailable")

        if chosen_period is None:
            assessments.append(
                ClaimIdentityAssessment(
                    occurrence=occurrence,
                    identity=None,
                    unavailable_reasons=tuple(reasons),
                )
            )
            continue

        period_kind, period_key = chosen_period
        source_anchor = (
            occurrence.line
            if occurrence.surface_kind == "text"
            else occurrence.surface_anchor
        )
        skeleton = normalized_anchor(source_anchor)
        if not skeleton:
            reasons.append("skeleton_unavailable")
        surface_anchor = normalized_anchor(source_anchor)
        if not surface_anchor:
            reasons.append("anchor_unavailable")
        unit_key = occurrence.figure.unit_key
        scale = float(occurrence.figure.scale)
        decimals = int(occurrence.figure.decimals)
        if not math.isfinite(scale) or scale <= 0:
            reasons.append("scale_invalid")
        if decimals < 0:
            reasons.append("decimals_invalid")

        if reasons:
            assessments.append(
                ClaimIdentityAssessment(
                    occurrence=occurrence,
                    identity=None,
                    unavailable_reasons=tuple(reasons),
                )
            )
            continue

        identity = SemanticClaimIdentityV1(
            semantic_skeleton=skeleton,
            period_kind=period_kind,
            period_key=period_key,
            unit_key=unit_key,
            scale=scale,
            decimals=decimals,
            surface_kind=occurrence.surface_kind,
            surface_anchor=surface_anchor,
            occurrence_ordinal=occurrence.figure_index,
        )
        assessments.append(
            ClaimIdentityAssessment(
                occurrence=occurrence,
                identity=identity,
                unavailable_reasons=(),
            )
        )

    return assessments
