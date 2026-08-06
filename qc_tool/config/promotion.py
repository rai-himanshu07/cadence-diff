"""Safe typed promotion of reviewed findings into reporting-contract rules."""

from __future__ import annotations

import datetime as dt
from enum import StrEnum

from pydantic import BaseModel, Field, model_validator

from qc_tool.config.profile import (
    AcceptanceBand,
    DeliverableProfile,
    FindingWaiver,
    NumericBoundsControl,
    RangeControl,
    SheetProfile,
)
from qc_tool.findings import Finding, FindingClass


class PromotionKind(StrEnum):
    WAIVER = "waiver"
    REQUIRED_RANGE = "required_range"
    NUMERIC_BOUND = "numeric_bound"
    ACCEPTANCE_BAND = "acceptance_band"


class PromotionRequest(BaseModel):
    kind: PromotionKind
    name: str = ""
    reason: str = ""
    expires: dt.date | None = None
    minimum: float | None = None
    maximum: float | None = None
    absolute: float = Field(default=0.0, ge=0.0)
    relative: float = Field(default=0.0, ge=0.0)

    @model_validator(mode="after")
    def validate_kind_fields(self):
        if self.kind is PromotionKind.WAIVER:
            if not self.reason.strip() or self.expires is None:
                raise ValueError("waiver promotion requires a reason and expiry")
        elif not self.name.strip():
            raise ValueError("control promotion requires a name")
        if (
            self.kind is PromotionKind.NUMERIC_BOUND
            and self.minimum is None
            and self.maximum is None
        ):
            raise ValueError("numeric-bound promotion requires a minimum or maximum")
        if (
            self.kind is PromotionKind.ACCEPTANCE_BAND
            and self.absolute <= 0
            and self.relative <= 0
        ):
            raise ValueError("acceptance-band promotion requires a positive bound")
        return self


_PROMOTION_CLASSES: dict[FindingClass, tuple[PromotionKind, ...]] = {
    FindingClass.REQUIRED_VALUE_MISSING: (PromotionKind.REQUIRED_RANGE,),
    FindingClass.NUMERIC_BOUND_VIOLATION: (PromotionKind.NUMERIC_BOUND,),
    FindingClass.VALUE_CHANGED: (PromotionKind.ACCEPTANCE_BAND,),
}


def available_promotions(finding: Finding) -> tuple[PromotionKind, ...]:
    specific = _PROMOTION_CLASSES.get(finding.finding_class, ())
    return (PromotionKind.WAIVER, *specific)


def promote_finding(
    profile: DeliverableProfile,
    finding: Finding,
    request: PromotionRequest,
) -> DeliverableProfile:
    """Return a validated draft; never mutate or save the input profile."""
    if request.kind not in available_promotions(finding):
        raise ValueError("this finding does not support the requested control")
    draft = profile.model_copy(deep=True)
    if request.kind is PromotionKind.WAIVER:
        if request.expires is None:
            raise ValueError("waiver promotion requires an expiry")
        waiver = FindingWaiver(
            finding_class=finding.finding_class,
            reason=request.reason.strip(),
            expires=request.expires,
            sheet=finding.sheet,
            slide=finding.slide,
            location=finding.location or finding.baseline_location,
            element=finding.element,
        )
        identity = waiver.model_dump(mode="json")
        if any(item.model_dump(mode="json") == identity for item in draft.waivers):
            raise ValueError("an identical waiver already exists")
        draft.waivers.append(waiver)
        return DeliverableProfile.model_validate(draft.model_dump(mode="json"))

    if not finding.sheet or not (finding.location or finding.baseline_location):
        raise ValueError("control promotion requires an exact worksheet range")
    location = finding.location or finding.baseline_location
    if location is None:
        raise ValueError("control promotion requires an exact worksheet range")
    if request.kind is PromotionKind.REQUIRED_RANGE:
        draft.excel.controls.required_ranges.append(
            RangeControl.model_validate(
                {
                    "name": request.name.strip(),
                    "sheet": finding.sheet,
                    "range": location,
                }
            )
        )
    elif request.kind is PromotionKind.NUMERIC_BOUND:
        draft.excel.controls.numeric_bounds.append(
            NumericBoundsControl.model_validate(
                {
                    "name": request.name.strip(),
                    "sheet": finding.sheet,
                    "range": location,
                    "minimum": request.minimum,
                    "maximum": request.maximum,
                }
            )
        )
    else:
        sheet_profile = draft.excel.sheets.setdefault(finding.sheet, SheetProfile())
        sheet_profile.acceptance_bands.append(
            AcceptanceBand.model_validate(
                {
                    "range": location,
                    "absolute": request.absolute,
                    "relative": request.relative,
                }
            )
        )
    return DeliverableProfile.model_validate(draft.model_dump(mode="json"))
