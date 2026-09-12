"""Native engine boundary for the ``cadence-diff-native`` helper package.

The helper distribution always installs for cadence-diff 2.x: supported
platforms receive a Rust/PyO3 wheel, while other platforms receive an
importable pure-Python fallback module. Import success is therefore not proof
that native execution is available. Every caller MUST use
``native_kernel_status()`` or ``native_kernel_available()``.

See ``docs/plans/plan-20260906-group-first-and-native-kernel.md``'s Phase B
for the full design; this module is only the thin, defensive import boundary
``qc_tool/io/loader.py`` calls through.
"""

from __future__ import annotations

import importlib.metadata
from enum import StrEnum
from types import ModuleType
from typing import Any

try:
    import cadence_diff_native as _native_module  # pyright: ignore[reportMissingModuleSource]
except ImportError:  # pragma: no cover - exercised by an unpatched real env
    _native_module = None

#: `(row, col, num, boolean, text)` per decoded cell; exactly one of
#: `num`/`boolean`/`text` is populated, mirroring pyxlsb's own raw `.v` shape.
RawValueCell = tuple[int, int, float | None, bool | None, str | None]
FormulaDeltaInput = tuple[str, str, str, str]
FormulaDeltaOutput = tuple[bool, bool, str | None, bool | None, str, bool]
FormulaSheetRaw = tuple[str, list[int], list[int], list[str], list[int], list[str]]
DefinedNameRaw = tuple[str, str | None]
EXPECTED_KERNEL_API_VERSION = 1
_REQUIRED_NATIVE_SYMBOLS = (
    "raw_values_report",
    "formula_surface_report",
    "formula_r1c1_report",
    "formula_delta_batch",
)


class NativeKernelStatus(StrEnum):
    """Content-free reason the helper is or is not usable as native code."""

    MISSING = "missing"
    FALLBACK_STUB = "fallback_stub"
    API_MISMATCH = "api_mismatch"
    MISSING_SYMBOL = "missing_symbol"
    AVAILABLE = "available"


def native_kernel_status() -> NativeKernelStatus:
    module = _native_module
    if module is None:
        return NativeKernelStatus.MISSING
    native_available = getattr(module, "__native_available__", None)
    if native_available is False:
        return NativeKernelStatus.FALLBACK_STUB
    if native_available is not True:
        return NativeKernelStatus.API_MISMATCH
    if getattr(module, "__kernel_api_version__", None) != EXPECTED_KERNEL_API_VERSION:
        return NativeKernelStatus.API_MISMATCH
    if any(not callable(getattr(module, name, None)) for name in _REQUIRED_NATIVE_SYMBOLS):
        return NativeKernelStatus.MISSING_SYMBOL
    return NativeKernelStatus.AVAILABLE


def native_kernel_available() -> bool:
    return native_kernel_status() is NativeKernelStatus.AVAILABLE


def native_distribution_version() -> str | None:
    """Installed helper version, including fallback wheels, when knowable."""
    module = _native_module
    if module is None:
        return None
    try:
        return importlib.metadata.version("cadence-diff-native")
    except importlib.metadata.PackageNotFoundError:
        version = getattr(module, "__version__", None)
        return version if isinstance(version, str) and version else None


def _require_native_module() -> ModuleType | Any:
    status = native_kernel_status()
    if status is not NativeKernelStatus.AVAILABLE:
        raise RuntimeError(f"cadence-diff native engine unavailable ({status.value})")
    return _native_module


def native_values_fingerprint() -> str | None:
    """Resolved native values-decoder identity, or None when unavailable."""
    if not native_kernel_available():
        return None
    return f"native-biff12:{native_distribution_version() or 'unknown'}"


def raw_values_report(data: bytes) -> list[tuple[str, list[RawValueCell]]]:
    """Per-sheet raw cell values, decoded from in-memory XLSB bytes.

    Raises ``RuntimeError`` if the extension is not installed; callers must
    check ``native_kernel_available()`` first and fall back to pyxlsb instead
    of relying on this exception as control flow.
    """
    module = _require_native_module()
    return module.raw_values_report(data)


def formula_delta_batch(pairs: list[FormulaDeltaInput]) -> list[FormulaDeltaOutput]:
    """Classify a bounded formula-pair batch, with per-row fallback markers."""
    module = _require_native_module()
    return module.formula_delta_batch(pairs)


def formula_surface_report(
    data: bytes,
) -> tuple[list[FormulaSheetRaw], list[DefinedNameRaw]]:
    """Raw formula surface from the compatible native helper."""
    module = _require_native_module()
    return module.formula_surface_report(data)


def formula_r1c1_report(data: bytes) -> list[object]:
    """Raw R1C1 report from the compatible native helper."""
    module = _require_native_module()
    return module.formula_r1c1_report(data)
