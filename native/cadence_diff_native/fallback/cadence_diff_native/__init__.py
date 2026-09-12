"""Pure-Python fallback for platforms without the native cadence-diff engine."""

from typing import NoReturn

__version__ = "2.0.0"
__kernel_api_version__ = 1
__native_available__ = False


def _unavailable(*args: object, **kwargs: object) -> NoReturn:
    del args, kwargs
    raise RuntimeError("the cadence-diff native engine is unavailable")


formula_delta_batch = _unavailable
formula_r1c1_report = _unavailable
raw_values_report = _unavailable
formula_surface_report = _unavailable
