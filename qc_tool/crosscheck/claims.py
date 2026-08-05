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

from dataclasses import dataclass

from qc_tool.crosscheck.trace import FigureOccurrence
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

    @property
    def total_inspected_surfaces(self) -> int:
        return self.readable + self.unreadable_shapes

    @property
    def complete(self) -> bool:
        """Whether every material surface in the deck could be read."""
        return self.unreadable_shapes == 0

    def reconciles(self, mapped: int, unmapped: int) -> bool:
        """Whether the mapped split accounts for exactly the readable claims."""
        return mapped + unmapped == self.readable


def claim_population(
    deck: DeckSnapshot, occurrences: list[FigureOccurrence]
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
    return ClaimPopulation(
        readable=len(occurrences),
        unreadable_shapes=unreadable,
        unreadable_slides=tuple(slides),
    )
