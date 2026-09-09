"""Optional native XLSB kernel adapter (Rust/PyO3, ``native/xlsbkernel/``).

The extension is an optional accelerator, not a runtime dependency: it may be
absent on any platform without a locally built or (later, per Phase B's
packaging plan) published matching wheel. Every caller MUST treat
``native_kernel_available()`` as the single source of truth and degrade to
the existing ``pyxlsb`` path without raising -- never assume the import
succeeded just because this module imported cleanly.

See ``docs/plans/plan-20260906-group-first-and-native-kernel.md``'s Phase B
for the full design; this module is only the thin, defensive import boundary
``qc_tool/io/loader.py`` calls through.
"""

from __future__ import annotations

try:
    import xlsbkernel as _xlsbkernel  # pyright: ignore[reportMissingModuleSource]
except ImportError:  # pragma: no cover - exercised by an unpatched real env
    _xlsbkernel = None

#: `(row, col, num, boolean, text)` per decoded cell; exactly one of
#: `num`/`boolean`/`text` is populated, mirroring pyxlsb's own raw `.v` shape.
RawValueCell = tuple[int, int, float | None, bool | None, str | None]


def native_kernel_available() -> bool:
    return _xlsbkernel is not None


def raw_values_report(data: bytes) -> list[tuple[str, list[RawValueCell]]]:
    """Per-sheet raw cell values, decoded from in-memory XLSB bytes.

    Raises ``RuntimeError`` if the extension is not installed; callers must
    check ``native_kernel_available()`` first and fall back to pyxlsb instead
    of relying on this exception as control flow.
    """
    if _xlsbkernel is None:
        raise RuntimeError("native xlsbkernel extension is not installed")
    return _xlsbkernel.raw_values_report(data)  # pyright: ignore[reportAttributeAccessIssue]
