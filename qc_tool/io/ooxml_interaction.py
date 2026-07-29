"""Canonical Excel data-validation and conditional-format extraction."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any, cast
from xml.etree.ElementTree import Element

from openpyxl.formatting.formatting import ConditionalFormatting
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.datavalidation import DataValidation

from qc_tool.io.model import ConditionalFormatDescriptor, DataValidationDescriptor

_SUPPORTED_RULE_TYPES = frozenset(
    {
        "aboveAverage",
        "beginsWith",
        "blanks",
        "cellIs",
        "containsBlanks",
        "containsErrors",
        "containsText",
        "duplicateValues",
        "endsWith",
        "errors",
        "expression",
        "notContainsBlanks",
        "notContainsErrors",
        "notContainsText",
        "notErrors",
        "notBlanks",
        "timePeriod",
        "top10",
        "uniqueValues",
    }
)


@dataclass(slots=True)
class InteractionExtraction:
    data_validations: list[DataValidationDescriptor] = field(default_factory=list)
    conditional_formats: list[ConditionalFormatDescriptor] = field(default_factory=list)
    rules_supported: bool = True
    rule_details: list[str] = field(default_factory=list)
    styles_supported: bool = True
    style_details: list[str] = field(default_factory=list)


def _optional_string(value: object) -> str | None:
    return None if value is None else str(value)


def _range_key(cell_range: str) -> tuple[int, int, int, int, str]:
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    except ValueError:
        return 0, 0, 0, 0, cell_range
    return (
        min_row or 0,
        min_col or 0,
        max_row or 0,
        max_col or 0,
        cell_range,
    )


def _target_ranges(sqref: object) -> tuple[str, ...]:
    ranges = getattr(sqref, "ranges", None)
    values = [str(item) for item in ranges] if ranges is not None else [str(sqref)]
    return tuple(sorted(values, key=_range_key))


def _color_value(color: Any, context: str) -> tuple[str | None, str | None]:
    if color is None:
        return None, None
    color_type = getattr(color, "type", None)
    if color_type == "rgb":
        value = getattr(color, "rgb", None)
        if value is None:
            return None, None
        normalized = str(value).upper()
        return normalized[-6:] if len(normalized) in {6, 8} else normalized, None
    if color_type is None:
        return None, None
    return None, f"{context} uses unresolved {color_type} color"


def _font_style(font: Any, details: list[str]) -> dict[str, object] | None:
    if font is None:
        return None
    color, issue = _color_value(getattr(font, "color", None), "font")
    if issue:
        details.append(issue)
    canonical = {
        "name": getattr(font, "name", None),
        "size": getattr(font, "sz", None),
        "bold": bool(getattr(font, "b", False)),
        "italic": bool(getattr(font, "i", False)),
        "underline": getattr(font, "u", None),
        "strike": bool(getattr(font, "strike", False)),
        "vertical_align": getattr(font, "vertAlign", None),
        "outline": bool(getattr(font, "outline", False)),
        "shadow": bool(getattr(font, "shadow", False)),
        "color": color,
    }
    return canonical if any(value not in {None, False} for value in canonical.values()) else None


def _fill_style(fill: Any, details: list[str]) -> dict[str, object] | None:
    if fill is None:
        return None
    fill_type = getattr(fill, "fill_type", None)
    if fill_type is None:
        return None
    foreground, foreground_issue = _color_value(
        getattr(fill, "fgColor", None),
        "fill foreground",
    )
    background, background_issue = _color_value(
        getattr(fill, "bgColor", None),
        "fill background",
    )
    for issue in (foreground_issue, background_issue):
        if issue:
            details.append(issue)
    return {
        "type": fill_type,
        "foreground": foreground,
        "background": None if fill_type == "solid" else background,
    }


def _side_style(side: Any, name: str, details: list[str]) -> dict[str, object]:
    color, issue = _color_value(getattr(side, "color", None), f"{name} border")
    if issue:
        details.append(issue)
    return {"style": getattr(side, "style", None), "color": color}


def _border_style(border: Any, details: list[str]) -> dict[str, object] | None:
    if border is None:
        return None
    sides = {
        name: _side_style(getattr(border, name, None), name, details)
        for name in (
            "left",
            "right",
            "top",
            "bottom",
            "diagonal",
            "vertical",
            "horizontal",
        )
        if getattr(border, name, None) is not None
    }
    canonical = {
        "sides": sides,
        "diagonal_up": bool(getattr(border, "diagonalUp", False)),
        "diagonal_down": bool(getattr(border, "diagonalDown", False)),
        "outline": bool(getattr(border, "outline", False)),
    }
    has_side = any(
        side["style"] is not None or side["color"] is not None for side in sides.values()
    )
    has_border_flag = any(canonical[name] for name in ("diagonal_up", "diagonal_down", "outline"))
    return canonical if has_side or has_border_flag else None


def _alignment_style(alignment: Any) -> dict[str, object] | None:
    if alignment is None:
        return None
    canonical: dict[str, object] = {
        name: getattr(alignment, name, None)
        for name in (
            "horizontal",
            "vertical",
            "textRotation",
            "wrapText",
            "shrinkToFit",
            "indent",
            "relativeIndent",
            "justifyLastLine",
            "readingOrder",
        )
    }
    for name in ("textRotation", "indent", "relativeIndent", "readingOrder"):
        if canonical[name] == 0:
            canonical[name] = None
    return canonical if any(value is not None for value in canonical.values()) else None


def _protection_style(protection: Any) -> dict[str, object] | None:
    if protection is None:
        return None
    return {
        "locked": getattr(protection, "locked", None),
        "hidden": getattr(protection, "hidden", None),
    }


def _differential_style(differential: Any) -> tuple[str | None, bool, str]:
    if differential is None:
        return None, True, ""
    details: list[str] = []
    number_format = getattr(differential, "numFmt", None)
    canonical = {
        "font": _font_style(getattr(differential, "font", None), details),
        "fill": _fill_style(getattr(differential, "fill", None), details),
        "border": _border_style(getattr(differential, "border", None), details),
        "number_format": (
            {
                "id": getattr(number_format, "numFmtId", None),
                "code": getattr(number_format, "formatCode", None),
            }
            if number_format is not None
            else None
        ),
        "alignment": _alignment_style(getattr(differential, "alignment", None)),
        "protection": _protection_style(getattr(differential, "protection", None)),
    }
    if details:
        return None, False, "; ".join(sorted(set(details)))
    if all(value is None for value in canonical.values()):
        return None, True, ""
    return json.dumps(canonical, sort_keys=True, separators=(",", ":")), True, ""


def _validation_descriptor(
    validation: Any,
    *,
    sheet: str,
    source_index: int,
) -> DataValidationDescriptor:
    return DataValidationDescriptor(
        sheet=sheet,
        source_index=source_index,
        source_id=f"{sheet}!validation[{source_index}]",
        target_ranges=_target_ranges(validation.sqref),
        validation_type=_optional_string(getattr(validation, "type", None)),
        operator=_optional_string(getattr(validation, "operator", None)),
        formula1=_optional_string(getattr(validation, "formula1", None)),
        formula2=_optional_string(getattr(validation, "formula2", None)),
        allow_blank=getattr(validation, "allowBlank", None),
        dropdown_suppressed=getattr(validation, "showDropDown", None),
        show_error_message=getattr(validation, "showErrorMessage", None),
        show_input_message=getattr(validation, "showInputMessage", None),
        error_style=_optional_string(getattr(validation, "errorStyle", None)),
        prompt_title=_optional_string(getattr(validation, "promptTitle", None)),
        prompt=_optional_string(getattr(validation, "prompt", None)),
        error_title=_optional_string(getattr(validation, "errorTitle", None)),
        error=_optional_string(getattr(validation, "error", None)),
    )


def _conditional_descriptor(
    rule: Any,
    *,
    sheet: str,
    source_index: int,
    target_ranges: tuple[str, ...],
) -> ConditionalFormatDescriptor:
    rule_type = str(getattr(rule, "type", "unknown"))
    semantic_supported = rule_type in _SUPPORTED_RULE_TYPES
    semantic_detail = "" if semantic_supported else f"{rule_type} rule details are unsupported"
    style_key, style_supported, style_detail = _differential_style(getattr(rule, "dxf", None))
    formulas = tuple(str(item) for item in (getattr(rule, "formula", None) or []))
    return ConditionalFormatDescriptor(
        sheet=sheet,
        source_index=source_index,
        source_id=f"{sheet}!conditional-format[{source_index}]",
        target_ranges=target_ranges,
        rule_type=rule_type,
        operator=_optional_string(getattr(rule, "operator", None)),
        formulas=formulas,
        priority=getattr(rule, "priority", None),
        stop_if_true=getattr(rule, "stopIfTrue", None),
        text=_optional_string(getattr(rule, "text", None)),
        time_period=_optional_string(getattr(rule, "timePeriod", None)),
        rank=getattr(rule, "rank", None),
        percent=getattr(rule, "percent", None),
        bottom=getattr(rule, "bottom", None),
        above_average=getattr(rule, "aboveAverage", None),
        equal_average=getattr(rule, "equalAverage", None),
        standard_deviations=getattr(rule, "stdDev", None),
        style_key=style_key,
        semantic_supported=semantic_supported,
        semantic_detail=semantic_detail,
        style_supported=style_supported,
        style_detail=style_detail,
    )


def extract_worksheet_interactions(worksheet: Any, sheet: str) -> InteractionExtraction:
    """Extract one worksheet's interaction rules through stable public APIs."""
    extraction = InteractionExtraction()
    validations = getattr(
        getattr(worksheet, "data_validations", None),
        "dataValidation",
        [],
    )
    extraction.data_validations.extend(
        _validation_descriptor(validation, sheet=sheet, source_index=index)
        for index, validation in enumerate(validations)
    )
    source_index = 0
    for group in worksheet.conditional_formatting:
        target_ranges = _target_ranges(group.sqref)
        for rule in group.rules:
            descriptor = _conditional_descriptor(
                rule,
                sheet=sheet,
                source_index=source_index,
                target_ranges=target_ranges,
            )
            extraction.conditional_formats.append(descriptor)
            if not descriptor.semantic_supported:
                extraction.rules_supported = False
                extraction.rule_details.append(descriptor.semantic_detail)
            if not descriptor.style_supported:
                extraction.styles_supported = False
                extraction.style_details.append(descriptor.style_detail)
            source_index += 1
    return extraction


def extract_data_validation_element(
    element: Element,
    *,
    sheet: str,
    source_index: int,
) -> DataValidationDescriptor:
    """Parse one raw OOXML data-validation element into the snapshot contract."""
    validation = cast(
        DataValidation | None,
        DataValidation.from_tree(cast(Any, element)),
    )
    if validation is None:
        raise ValueError("data-validation element could not be parsed")
    return _validation_descriptor(
        validation,
        sheet=sheet,
        source_index=source_index,
    )


def extract_conditional_formatting_element(
    element: Element,
    *,
    sheet: str,
    source_index: int,
    differential_styles: list[DifferentialStyle],
) -> InteractionExtraction:
    """Parse one raw OOXML conditional-format group into canonical descriptors."""
    extraction = InteractionExtraction()
    group = cast(
        ConditionalFormatting | None,
        ConditionalFormatting.from_tree(cast(Any, element)),
    )
    if group is None:
        raise ValueError("conditional-format element could not be parsed")
    target_ranges = _target_ranges(group.sqref)
    for offset, rule in enumerate(group.rules):
        style_id = getattr(rule, "dxfId", None)
        if style_id is not None:
            if 0 <= style_id < len(differential_styles):
                rule.dxf = differential_styles[style_id]
            else:
                descriptor = _conditional_descriptor(
                    rule,
                    sheet=sheet,
                    source_index=source_index + offset,
                    target_ranges=target_ranges,
                )
                descriptor.style_key = None
                descriptor.style_supported = False
                descriptor.style_detail = f"differential style {style_id} is unavailable"
                extraction.conditional_formats.append(descriptor)
                extraction.styles_supported = False
                extraction.style_details.append(descriptor.style_detail)
                if not descriptor.semantic_supported:
                    extraction.rules_supported = False
                    extraction.rule_details.append(descriptor.semantic_detail)
                continue
        descriptor = _conditional_descriptor(
            rule,
            sheet=sheet,
            source_index=source_index + offset,
            target_ranges=target_ranges,
        )
        extraction.conditional_formats.append(descriptor)
        if not descriptor.semantic_supported:
            extraction.rules_supported = False
            extraction.rule_details.append(descriptor.semantic_detail)
        if not descriptor.style_supported:
            extraction.styles_supported = False
            extraction.style_details.append(descriptor.style_detail)
    return extraction
