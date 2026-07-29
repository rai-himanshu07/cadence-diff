"""PowerPoint diff over matched slides and semantically matched elements."""

import difflib

from qc_tool.config.profile import PptProfile
from qc_tool.crosscheck.numbers import numeric_skeleton as _numeric_skeleton
from qc_tool.findings import Finding, FindingClass
from qc_tool.ppt.element_diff import diff_slide_elements
from qc_tool.ppt.match import SlideMatching
from qc_tool.ppt.model import SlideContent


def _text_findings(baseline: SlideContent, current: SlideContent) -> list[Finding]:
    findings: list[Finding] = []
    slide = current.display_name
    matcher = difflib.SequenceMatcher(
        a=baseline.texts,
        b=current.texts,
        autojunk=False,
    )
    for operation, baseline_start, baseline_end, current_start, current_end in (
        matcher.get_opcodes()
    ):
        if operation == "equal":
            continue
        baseline_lines = baseline.texts[baseline_start:baseline_end]
        current_lines = current.texts[current_start:current_end]
        for baseline_line, current_line in zip(
            baseline_lines,
            current_lines,
            strict=False,
        ):
            expected = _numeric_skeleton(baseline_line) == _numeric_skeleton(
                current_line
            )
            findings.append(
                Finding(
                    artifact="ppt",
                    finding_class=FindingClass.SLIDE_TEXT_CHANGED,
                    expected_growth=expected,
                    slide=slide,
                    baseline_value=baseline_line,
                    current_value=current_line,
                    message=(
                        f"{slide}: "
                        + ("figure refreshed" if expected else "text changed")
                    ),
                )
            )
        for line in baseline_lines[len(current_lines) :]:
            findings.append(
                Finding(
                    artifact="ppt",
                    finding_class=FindingClass.SLIDE_TEXT_CHANGED,
                    slide=slide,
                    baseline_value=line,
                    message=f"{slide}: text removed",
                )
            )
        for line in current_lines[len(baseline_lines) :]:
            findings.append(
                Finding(
                    artifact="ppt",
                    finding_class=FindingClass.SLIDE_TEXT_CHANGED,
                    slide=slide,
                    current_value=line,
                    message=f"{slide}: text added",
                )
            )
    return findings


def diff_decks(
    matching: SlideMatching,
    profile: PptProfile | None = None,
) -> list[Finding]:
    """Diff slides and every semantically matched table/chart descendant."""
    profile = profile or PptProfile()
    findings: list[Finding] = []
    for slide in matching.added:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.SLIDE_ADDED,
                slide=slide.display_name,
                location=f"slide {slide.index + 1}",
                message=f"slide {slide.display_name!r} added in current deck",
            )
        )
    for slide in matching.removed:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.SLIDE_REMOVED,
                slide=slide.display_name,
                location=f"slide {slide.index + 1}",
                message=f"slide {slide.display_name!r} removed from current deck",
            )
        )
    for baseline_slide, current_slide in matching.reordered:
        findings.append(
            Finding(
                artifact="ppt",
                finding_class=FindingClass.SLIDE_REORDERED,
                expected_growth=True,
                slide=current_slide.display_name,
                baseline_value=f"position {baseline_slide.index + 1}",
                current_value=f"position {current_slide.index + 1}",
                message=(
                    f"slide {current_slide.display_name!r} moved from position "
                    f"{baseline_slide.index + 1} to {current_slide.index + 1}"
                ),
            )
        )
    for baseline_slide, current_slide in matching.pairs:
        findings.extend(_text_findings(baseline_slide, current_slide))
        findings.extend(
            diff_slide_elements(baseline_slide, current_slide, profile)
        )
    return findings
