"""Native helper availability and compatibility contracts for 2.0."""

from __future__ import annotations

from types import SimpleNamespace

import pytest

import qc_tool.io.native_kernel as native_kernel_module
from qc_tool.io.native_kernel import (
    NativeKernelStatus,
    native_kernel_available,
    native_kernel_status,
    native_values_fingerprint,
    raw_values_report,
)


def _module(
    *,
    available: bool = True,
    api_version: int = 1,
    omit: str | None = None,
) -> SimpleNamespace:
    values: dict[str, object] = {
        "__version__": "2.0.0",
        "__kernel_api_version__": api_version,
        "__native_available__": available,
        "raw_values_report": lambda data: [],
        "formula_surface_report": lambda data: ([], []),
        "formula_r1c1_report": lambda data: [],
        "formula_delta_batch": lambda pairs: [],
    }
    if omit is not None:
        values.pop(omit)
    return SimpleNamespace(**values)


@pytest.mark.parametrize(
    ("module", "expected"),
    [
        (None, NativeKernelStatus.MISSING),
        (_module(available=False), NativeKernelStatus.FALLBACK_STUB),
        (_module(api_version=2), NativeKernelStatus.API_MISMATCH),
        (_module(omit="formula_delta_batch"), NativeKernelStatus.MISSING_SYMBOL),
        (_module(), NativeKernelStatus.AVAILABLE),
    ],
)
def test_native_status_requires_availability_api_and_symbols(
    monkeypatch: pytest.MonkeyPatch,
    module: SimpleNamespace | None,
    expected: NativeKernelStatus,
) -> None:
    monkeypatch.setattr(native_kernel_module, "_native_module", module)

    assert native_kernel_status() is expected
    assert native_kernel_available() is (expected is NativeKernelStatus.AVAILABLE)


def test_fallback_stub_cannot_be_called_as_native(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(native_kernel_module, "_native_module", _module(available=False))

    with pytest.raises(RuntimeError, match="fallback_stub"):
        raw_values_report(b"unused")


def test_native_fingerprint_uses_project_owned_distribution_name(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(native_kernel_module, "_native_module", _module())
    seen: list[str] = []

    def version(name: str) -> str:
        seen.append(name)
        return "2.0.0"

    monkeypatch.setattr(native_kernel_module.importlib.metadata, "version", version)

    assert native_values_fingerprint() == "native-biff12:2.0.0"
    assert seen == ["cadence-diff-native"]
