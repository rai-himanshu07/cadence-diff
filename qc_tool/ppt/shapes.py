"""Structural PowerPoint shape walkers shared by read and rewrite paths."""

from __future__ import annotations

from collections.abc import Iterator
from typing import Any


def _shape_type_name(shape: Any) -> str:
    shape_type = shape.shape_type
    return str(getattr(shape_type, "name", shape_type))


def iter_text_leaf_shapes(shape: Any) -> Iterator[Any]:
    """Yield non-group text-frame descendants in depth-first shape order."""
    if _shape_type_name(shape) == "GROUP":
        for child in shape.shapes:
            yield from iter_text_leaf_shapes(child)
        return
    if shape.has_text_frame:
        yield shape
