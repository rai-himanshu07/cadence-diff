"""Fuzzy slide pairing that survives add/delete/reorder.

Pair scores combine title similarity, body-content similarity, and a
position prior. Profile pins (baseline title -> current title) are applied
first and always win. Remaining slides pair greedily by descending score
above the profile threshold — at deck scale this matches the optimal
assignment in practice and avoids a scipy dependency; upgrade to Hungarian
if real decks show conflicts.

Reorder detection: paired slides outside the longest increasing
subsequence of current positions (in baseline order) moved relative to
their peers — pure index shifts caused by insertions do not flag.
"""

import logging
from dataclasses import dataclass, field

from rapidfuzz import fuzz

from qc_tool.config.profile import PptProfile
from qc_tool.ppt.extract import DeckSnapshot, SlideContent

logger = logging.getLogger(__name__)

_TITLE_WEIGHT = 0.5
_BODY_WEIGHT = 0.35
_POSITION_WEIGHT = 0.15


@dataclass(slots=True)
class SlideMatching:
    pairs: list[tuple[SlideContent, SlideContent]] = field(default_factory=list)
    removed: list[SlideContent] = field(default_factory=list)  # baseline only
    added: list[SlideContent] = field(default_factory=list)  # current only
    reordered: list[tuple[SlideContent, SlideContent]] = field(default_factory=list)


def _body_text(slide: SlideContent) -> str:
    parts = list(slide.texts)
    for table in slide.tables:
        parts.extend(cell for row in table.rows for cell in row)
    for chart in slide.charts:
        if chart.all_series:
            parts.extend(
                category
                for series in chart.all_series
                for category in series.categories
            )
        else:
            parts.extend(chart.categories)
    return " ".join(parts)


def slide_similarity(base: SlideContent, curr: SlideContent, deck_size: int) -> float:
    title_score = fuzz.ratio(base.title or "", curr.title or "")
    body_score = fuzz.token_set_ratio(_body_text(base), _body_text(curr))
    position_score = 100.0 * (1.0 - abs(base.index - curr.index) / max(deck_size, 1))
    return (
        _TITLE_WEIGHT * title_score
        + _BODY_WEIGHT * body_score
        + _POSITION_WEIGHT * position_score
    )


def _longest_increasing_subsequence(values: list[int]) -> set[int]:
    """Indices (into ``values``) forming one longest increasing subsequence."""
    if not values:
        return set()
    n = len(values)
    lengths = [1] * n
    parents = [-1] * n
    for i in range(1, n):
        for j in range(i):
            if values[j] < values[i] and lengths[j] + 1 > lengths[i]:
                lengths[i] = lengths[j] + 1
                parents[i] = j
    best = max(range(n), key=lambda i: lengths[i])
    keep: set[int] = set()
    while best != -1:
        keep.add(best)
        best = parents[best]
    return keep


def match_slides(
    baseline: DeckSnapshot, current: DeckSnapshot, profile: PptProfile | None = None
) -> SlideMatching:
    profile = profile or PptProfile()
    matching = SlideMatching()
    base_pool = list(baseline.slides)
    curr_pool = list(current.slides)
    deck_size = max(len(base_pool), len(curr_pool))

    # 1. Profile pins always win.
    for base_title, curr_title in profile.slide_pins.items():
        base_slide = next((s for s in base_pool if s.title == base_title), None)
        curr_slide = next((s for s in curr_pool if s.title == curr_title), None)
        if base_slide is None or curr_slide is None:
            logger.warning(
                "slide pin %r -> %r did not match both decks", base_title, curr_title
            )
            continue
        matching.pairs.append((base_slide, curr_slide))
        base_pool.remove(base_slide)
        curr_pool.remove(curr_slide)

    # 2. Greedy fuzzy assignment above the threshold.
    scored = sorted(
        (
            (slide_similarity(b, c, deck_size), b.index, c.index, b, c)
            for b in base_pool
            for c in curr_pool
        ),
        key=lambda item: (-item[0], item[1], item[2]),
    )
    used_base: set[int] = set()
    used_curr: set[int] = set()
    for score, _, _, base_slide, curr_slide in scored:
        if score < profile.match_threshold:
            break
        if base_slide.index in used_base or curr_slide.index in used_curr:
            continue
        matching.pairs.append((base_slide, curr_slide))
        used_base.add(base_slide.index)
        used_curr.add(curr_slide.index)

    matching.removed = [s for s in base_pool if s.index not in used_base]
    matching.added = [s for s in curr_pool if s.index not in used_curr]

    # 3. Relative-order analysis over all pairs.
    matching.pairs.sort(key=lambda pair: pair[0].index)
    current_positions = [curr.index for _, curr in matching.pairs]
    stable = _longest_increasing_subsequence(current_positions)
    matching.reordered = [
        pair for i, pair in enumerate(matching.pairs) if i not in stable
    ]
    return matching
