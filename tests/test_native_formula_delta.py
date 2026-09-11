"""Parity and performance gates for the optional Rust formula-delta batch."""

from __future__ import annotations

import random
import string
import time

import pytest

import qc_tool.excel.formulas as formulas_module
from qc_tool.excel.formula_tokens import formula_reference_operands
from qc_tool.excel.formulas import (
    FormulaComparisonTelemetry,
    _differs_only_by_extension,
    detect_formula_wrapper,
    diff_workbook_formulas,
)
from qc_tool.io import native_kernel as native_kernel_module
from qc_tool.io.native_kernel import (
    FormulaDeltaInput,
    formula_delta_batch,
    native_kernel_available,
)
from tests.test_formulas import (
    _alignment_for_rows,
    _diverse_construct_rows,
    _formula_column_workbook,
    _repeated_pattern_rows,
)

pytestmark = pytest.mark.skipif(
    not native_kernel_available(),
    reason="native/xlsbkernel is not built in this environment",
)


def _python_classify(pair: FormulaDeltaInput):
    base, current, base_norm, current_norm = pair
    expected = _differs_only_by_extension(base, current)
    wrapper = None if expected else detect_formula_wrapper(base_norm, current_norm)
    base_references = {
        operand.value.casefold() for operand in formula_reference_operands(base)
    }
    current_references = {
        operand.value.casefold() for operand in formula_reference_operands(current)
    }
    return (
        expected,
        wrapper.kind if wrapper is not None else None,
        wrapper.exact if wrapper is not None else None,
        (
            f"formula-wrapper:{wrapper.kind}:{wrapper.skeleton_key}"
            if wrapper is not None
            else ""
        ),
        bool(current_references - base_references),
    )


@pytest.mark.parametrize(
    "pair",
    [
        ("=SUM(C2:C10)", "=SUM(C2:C15)", "=SUM(RC:R[8]C)", "=SUM(RC:R[13]C)"),
        ("=B5", "=B5+C5", "=RC", "=RC+RC[1]"),
        ("=$B$2+B6", "=$B$2+B6*2", "=R2C2+RC", "=R2C2+RC*2"),
        ("=Other!A1+B7", "=Other!A1+B7*2", "=Other!R1C1+RC", "=Other!R1C1+RC*2"),
        ("=SUM(C:C)", "=SUM(C:C)*2", "=SUM(C:C)", "=SUM(C:C)*2"),
        ('="a,b"', '=IF(A1,"a,b","c")', '="a,b"', '=IF(RC[-1],"a,b","c")'),
        ("=#REF!", "=IFERROR(A1,#N/A)", "=#REF!", "=IFERROR(RC[-1],#N/A)"),
        ("={1,2;3,4}", "=SUM({1,2;3,4})", "={1,2;3,4}", "=SUM({1,2;3,4})"),
        (
            "=IF(R[1]C[-2],RC[3],NA())",
            "=IF(R1C8,C[3],IF(R[1]C[-2],RC[3],NA()),NA())",
            "=IF(R[1]C[-2],RC[3],NA())",
            "=IF(R1C8,C[3],IF(R[1]C[-2],RC[3],NA()),NA())",
        ),
        (
            "=IF(R1C8,C[3],IF(R[1]C[-2],RC[3],NA()),NA())",
            "=IF(R[1]C[-2],RC[3],NA())",
            "=IF(R1C8,C[3],IF(R[1]C[-2],RC[3],NA()),NA())",
            "=IF(R[1]C[-2],RC[3],NA())",
        ),
    ],
)
def test_rust_formula_delta_matches_python(pair: FormulaDeltaInput) -> None:
    actual = formula_delta_batch([pair])[0]

    assert actual[0]
    assert actual[1:] == _python_classify(pair)


@pytest.mark.parametrize(
    "pair",
    [
        ("=LET(x,A1,x+1)", "=LET(x,A1,x+2)", "=LET(x,RC[-1],x+1)", "=LET(x,RC[-1],x+2)"),
        (
            "=LAMBDA(x,x+1)(A1)",
            "=LAMBDA(x,x+2)(A1)",
            "=LAMBDA(x,x+1)(RC[-1])",
            "=LAMBDA(x,x+2)(RC[-1])",
        ),
        ("=B2#", "=B2#+1", "=_xlfn.ANCHORARRAY(RC)", "=_xlfn.ANCHORARRAY(RC)+1"),
        ("=@A1", "=@A1+1", "=_xlfn.SINGLE(RC[-1])", "=_xlfn.SINGLE(RC[-1])+1"),
        ("='Δ'!A1", "='Δ'!A1+1", "='Δ'!RC[-1]", "='Δ'!RC[-1]+1"),
        ('="unterminated', "=A1", '="unterminated', "=RC[-1]"),
    ],
)
def test_unsupported_or_malformed_rows_request_python_fallback(
    pair: FormulaDeltaInput,
) -> None:
    assert formula_delta_batch([pair])[0][0] is False


def test_malformed_batch_never_panics_or_changes_cardinality() -> None:
    generator = random.Random(20260911)
    alphabet = string.printable + "\x00\x01\x7f"
    pairs: list[FormulaDeltaInput] = []
    for _ in range(10_000):
        values = []
        for _ in range(4):
            prefix = "=" if generator.random() < 0.75 else ""
            values.append(
                prefix
                + "".join(
                    generator.choices(alphabet, k=generator.randint(0, 96))
                )
            )
        pairs.append((values[0], values[1], values[2], values[3]))
    deeply_nested = "=" + "F(" * 5_000 + "A1" + ")" * 5_000
    pairs.append((deeply_nested, "=A1", deeply_nested, "=RC"))

    actual = formula_delta_batch(pairs)

    assert len(actual) == len(pairs)
    assert all(len(row) == 6 and type(row[0]) is bool for row in actual)


def test_rust_formula_delta_matches_python_on_generated_reference_pairs() -> None:
    generator = random.Random(20260911)
    pairs: list[FormulaDeltaInput] = []
    operators = ("+", "-", "*", "/")
    for row in range(2, 502):
        left = generator.choice(("A", "B", "C", "D"))
        right = generator.choice(("E", "F", "G", "H"))
        operator = generator.choice(operators)
        multiplier = generator.randint(1, 9)
        pairs.append(
            (
                f"={left}{row}{operator}{right}{row}*{multiplier}",
                f"={left}{row}{operator}{right}{row}*{multiplier + 1}",
                f"=RC[-2]{operator}RC[-1]*{multiplier}",
                f"=RC[-2]{operator}RC[-1]*{multiplier + 1}",
            )
        )

    actual = formula_delta_batch(pairs)

    assert all(row[0] for row in actual)
    assert [row[1:] for row in actual] == [_python_classify(pair) for pair in pairs]


def test_rust_batch_is_at_least_three_times_faster_than_current_python_path() -> None:
    pairs: list[FormulaDeltaInput] = [
        (f"=A{row}*2", f"=A{row}*3", "=RC[-1]*2", "=RC[-1]*3")
        for row in range(20, 20_020)
    ]

    def python_run() -> None:
        wrapper_cache = {}
        for pair in pairs:
            base, current, base_norm, current_norm = pair
            expected = _differs_only_by_extension(base, current)
            key = (base_norm, current_norm)
            if not expected and key not in wrapper_cache:
                wrapper_cache[key] = detect_formula_wrapper(base_norm, current_norm)
            base_references = {
                operand.value.casefold()
                for operand in formula_reference_operands(base)
            }
            current_references = {
                operand.value.casefold()
                for operand in formula_reference_operands(current)
            }
            bool(current_references - base_references)

    python_run()
    formula_delta_batch(pairs)
    python_times = []
    rust_times = []
    for _ in range(3):
        started = time.process_time()
        python_run()
        python_times.append(time.process_time() - started)
        started = time.process_time()
        formula_delta_batch(pairs)
        rust_times.append(time.process_time() - started)

    speedup = min(python_times) / min(rust_times)
    assert speedup >= 3.0, f"Rust classifier speedup was only {speedup:.2f}x"


def test_native_integration_is_byte_identical_to_python_oracle() -> None:
    base_formulas, current_formulas = _diverse_construct_rows()
    rows = sorted(base_formulas)
    max_row = max(rows) + 1
    baseline = _formula_column_workbook(base_formulas, max_row=max_row)
    current = _formula_column_workbook(current_formulas, max_row=max_row)
    alignment = _alignment_for_rows(rows, max_row)
    expected = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        _use_native_delta=False,
    )
    telemetry = FormulaComparisonTelemetry()

    actual = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        telemetry=telemetry,
        _use_native_delta=True,
    )

    assert [finding.model_dump(mode="json") for finding in actual] == [
        finding.model_dump(mode="json") for finding in expected
    ]
    assert telemetry.native_delta_supported_pairs == len(rows)
    assert telemetry.native_delta_fallback_pairs == 0
    assert telemetry.native_delta_batches == 1


def test_native_integration_falls_back_per_unsupported_row() -> None:
    base_formulas = {2: "=LET(x,A2,x+1)", 3: "=A3*2"}
    current_formulas = {2: "=LET(x,A2,x+2)", 3: "=A3*3"}
    rows = sorted(base_formulas)
    baseline = _formula_column_workbook(base_formulas, max_row=4)
    current = _formula_column_workbook(current_formulas, max_row=4)
    alignment = _alignment_for_rows(rows, 4)
    expected = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        _use_native_delta=False,
    )
    telemetry = FormulaComparisonTelemetry()

    actual = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        telemetry=telemetry,
        _use_native_delta=True,
    )

    assert [finding.model_dump(mode="json") for finding in actual] == [
        finding.model_dump(mode="json") for finding in expected
    ]
    assert telemetry.native_delta_supported_pairs == 1
    assert telemetry.native_delta_fallback_pairs == 1
    assert telemetry.native_delta_declared_unsupported_pairs == 1


@pytest.mark.parametrize(
    "native_failure",
    ["malformed", "panic", "old_extension", "wrong_length"],
)
def test_native_integration_falls_back_on_invalid_extension_behavior(
    monkeypatch: pytest.MonkeyPatch,
    native_failure: str,
) -> None:
    baseline = _formula_column_workbook({2: "=A2*2"}, max_row=3)
    current = _formula_column_workbook({2: "=A2*3"}, max_row=3)
    alignment = _alignment_for_rows([2], 3)
    expected = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        _use_native_delta=False,
    )

    if native_failure == "malformed":
        monkeypatch.setattr(
            native_kernel_module,
            "formula_delta_batch",
            lambda pairs: [("invalid",)] * len(pairs),
        )
    elif native_failure == "wrong_length":
        monkeypatch.setattr(
            native_kernel_module,
            "formula_delta_batch",
            lambda _pairs: [],
        )
    else:
        class NativePanic(BaseException):
            pass

        def panic(_pairs: list[FormulaDeltaInput]) -> None:
            if native_failure == "old_extension":
                raise RuntimeError("installed extension has no classifier")
            raise NativePanic

        monkeypatch.setattr(native_kernel_module, "formula_delta_batch", panic)
    telemetry = FormulaComparisonTelemetry()

    actual = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        telemetry=telemetry,
        _use_native_delta=True,
    )

    assert [finding.model_dump(mode="json") for finding in actual] == [
        finding.model_dump(mode="json") for finding in expected
    ]
    assert telemetry.native_delta_supported_pairs == 0
    assert telemetry.native_delta_fallback_pairs == 1
    assert telemetry.native_delta_batch_failures == (
        native_failure in {"panic", "old_extension", "wrong_length"}
    )
    assert telemetry.native_delta_invalid_output_pairs == (
        native_failure == "malformed"
    )
    assert telemetry.native_delta_declared_unsupported_pairs == 0
    assert telemetry.native_delta_api_failures == (
        native_failure == "old_extension"
    )
    assert telemetry.native_delta_protocol_failures == (
        native_failure == "wrong_length"
    )
    assert telemetry.native_delta_runtime_failures == (
        native_failure == "panic"
    )


def test_native_integration_never_hides_memory_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    baseline = _formula_column_workbook({2: "=A2*2"}, max_row=3)
    current = _formula_column_workbook({2: "=A2*3"}, max_row=3)
    alignment = _alignment_for_rows([2], 3)

    def out_of_memory(_pairs: list[FormulaDeltaInput]) -> None:
        raise MemoryError

    monkeypatch.setattr(native_kernel_module, "formula_delta_batch", out_of_memory)

    with pytest.raises(MemoryError):
        diff_workbook_formulas(
            baseline,
            current,
            alignment,
            _use_native_delta=True,
        )


def test_single_oversized_occurrence_never_crosses_the_native_boundary(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    baseline = _formula_column_workbook({2: "=A2+A2*2"}, max_row=3)
    current = _formula_column_workbook({2: "=A2+A2*3"}, max_row=3)
    alignment = _alignment_for_rows([2], 3)
    expected = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        _use_native_delta=False,
    )
    native_calls = 0

    def observe(_pairs: list[FormulaDeltaInput]):
        nonlocal native_calls
        native_calls += 1
        return []

    monkeypatch.setattr(
        formulas_module,
        "_NATIVE_FORMULA_DELTA_MAX_STRING_BYTES",
        4,
    )
    monkeypatch.setattr(native_kernel_module, "formula_delta_batch", observe)
    telemetry = FormulaComparisonTelemetry()

    actual = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        telemetry=telemetry,
        _use_native_delta=True,
    )

    assert [finding.model_dump(mode="json") for finding in actual] == [
        finding.model_dump(mode="json") for finding in expected
    ]
    assert native_calls == 0
    assert telemetry.native_delta_batches == 0
    assert telemetry.native_delta_fallback_pairs == 1
    assert telemetry.native_delta_oversized_pairs == 1


def test_native_integration_bounds_each_batch_by_formula_bytes(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    base_formulas = {row: f"=A{row}*2" for row in range(2, 42)}
    current_formulas = {row: f"=A{row}*3" for row in range(2, 42)}
    rows = sorted(base_formulas)
    baseline = _formula_column_workbook(base_formulas, max_row=42)
    current = _formula_column_workbook(current_formulas, max_row=42)
    alignment = _alignment_for_rows(rows, 42)
    byte_budget = 160
    observed_batch_bytes: list[int] = []
    original = native_kernel_module.formula_delta_batch

    def observe(pairs: list[FormulaDeltaInput]):
        observed_batch_bytes.append(
            sum(len(value.encode("utf-8")) for pair in pairs for value in pair)
        )
        return original(pairs)

    monkeypatch.setattr(
        formulas_module,
        "_NATIVE_FORMULA_DELTA_BATCH_BYTES",
        byte_budget,
    )
    monkeypatch.setattr(native_kernel_module, "formula_delta_batch", observe)
    telemetry = FormulaComparisonTelemetry()

    findings = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        telemetry=telemetry,
        _use_native_delta=True,
    )

    assert len(findings) == len(rows)
    assert len(observed_batch_bytes) > 1
    assert max(observed_batch_bytes) <= byte_budget
    assert telemetry.native_delta_batches == len(observed_batch_bytes)
    assert telemetry.native_delta_batch_failures == 0


def test_native_integration_benchmark_preserves_exact_payloads() -> None:
    base_formulas, current_formulas, rows = _repeated_pattern_rows(20_000)
    max_row = max(rows) + 1
    baseline = _formula_column_workbook(base_formulas, max_row=max_row)
    current = _formula_column_workbook(current_formulas, max_row=max_row)
    alignment = _alignment_for_rows(rows, max_row)

    def run(*, native: bool):
        started = time.process_time()
        findings = diff_workbook_formulas(
            baseline,
            current,
            alignment,
            _use_native_delta=native,
        )
        return time.process_time() - started, findings

    run(native=False)
    run(native=True)
    python_times = []
    native_times = []
    expected = actual = []
    for _ in range(3):
        elapsed, expected = run(native=False)
        python_times.append(elapsed)
        elapsed, actual = run(native=True)
        native_times.append(elapsed)

    assert [finding.model_dump(mode="json") for finding in actual] == [
        finding.model_dump(mode="json") for finding in expected
    ]
    assert all(elapsed > 0 for elapsed in python_times + native_times)


def test_production_formula_delta_uses_native_by_default(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    baseline = _formula_column_workbook({2: "=A2*2"}, max_row=3)
    current = _formula_column_workbook({2: "=A2*3"}, max_row=3)
    alignment = _alignment_for_rows([2], 3)

    native_calls = 0
    original = native_kernel_module.formula_delta_batch

    def observe(pairs: list[FormulaDeltaInput]):
        nonlocal native_calls
        native_calls += 1
        return original(pairs)

    monkeypatch.setattr(
        native_kernel_module,
        "formula_delta_batch",
        observe,
    )
    telemetry = FormulaComparisonTelemetry()

    findings = diff_workbook_formulas(
        baseline,
        current,
        alignment,
        telemetry=telemetry,
    )

    assert len(findings) == 1
    assert native_calls == 1
    assert telemetry.native_delta_batches == 1
    assert telemetry.native_delta_supported_pairs == 1
