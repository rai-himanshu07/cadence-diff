"""Lexical LET/LAMBDA precedent extraction contracts."""

from __future__ import annotations

from qc_tool.excel.formula_tokens import (
    FormulaPrecedentKind,
    extract_formula_precedents,
)


def _values(formula: str, kind: FormulaPrecedentKind) -> list[str]:
    result = extract_formula_precedents(formula)
    return [precedent.value for precedent in result.precedents if precedent.kind is kind]


def test_let_bindings_enter_scope_after_each_value_expression() -> None:
    formula = "=LET(rate,A1,scaled,rate*B1,scaled+C1)"

    assert _values(formula, FormulaPrecedentKind.REFERENCE) == ["A1", "B1", "C1"]
    assert _values(formula, FormulaPrecedentKind.LOCAL_SYMBOL) == [
        "rate",
        "scaled",
        "rate",
        "scaled",
    ]


def test_let_binding_does_not_apply_to_its_own_value() -> None:
    formula = "=LET(rate,rate+A1,rate+B1)"

    assert _values(formula, FormulaPrecedentKind.REFERENCE) == ["rate", "A1", "B1"]
    assert _values(formula, FormulaPrecedentKind.LOCAL_SYMBOL) == ["rate", "rate"]


def test_lambda_parameters_apply_only_to_the_body() -> None:
    formula = "=LAMBDA(item,item+A1)(B1)"

    assert _values(formula, FormulaPrecedentKind.REFERENCE) == ["A1", "B1"]
    assert _values(formula, FormulaPrecedentKind.LOCAL_SYMBOL) == ["item", "item"]


def test_nested_scopes_and_shadowing_do_not_leak() -> None:
    formula = "=LET(x,A1,LAMBDA(x,LET(y,x+B1,y+C1))(D1)+x)"

    assert _values(formula, FormulaPrecedentKind.REFERENCE) == [
        "A1",
        "B1",
        "C1",
        "D1",
    ]
    assert _values(formula, FormulaPrecedentKind.LOCAL_SYMBOL) == [
        "x",
        "x",
        "y",
        "x",
        "y",
        "x",
    ]


def test_supported_higher_order_lambda_forms_keep_true_references() -> None:
    map_result = extract_formula_precedents("=MAP(A1:A5,LAMBDA(item,item*B1))")
    reduce_result = extract_formula_precedents(
        "=REDUCE(0,A1:A5,LAMBDA(total,item,total+item+B1))"
    )

    assert [
        item.value
        for item in map_result.precedents
        if item.kind is FormulaPrecedentKind.REFERENCE
    ] == ["A1:A5", "B1"]
    assert [
        item.value
        for item in reduce_result.precedents
        if item.kind is FormulaPrecedentKind.REFERENCE
    ] == ["A1:A5", "B1"]
    assert map_result.unsupported == ()
    assert reduce_result.unsupported == ()


def test_malformed_lexical_syntax_is_one_unsupported_construct() -> None:
    result = extract_formula_precedents("=LET(item,A1)")

    assert result.references == ()
    assert result.local_symbols == ()
    assert [item.reason for item in result.unsupported] == ["malformed_let"]


def test_non_lexical_and_dynamic_references_keep_existing_values() -> None:
    ordinary = extract_formula_precedents("=SUM(A1:A5)+B1")
    dynamic = extract_formula_precedents("=SUM(B2#)+@C1:C5")

    assert [item.value for item in ordinary.references] == ["A1:A5", "B1"]
    assert [item.value for item in dynamic.references] == ["B2#", "@C1:C5"]
    assert ordinary.local_symbols == dynamic.local_symbols == ()
