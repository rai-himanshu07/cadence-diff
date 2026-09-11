//! Bounded formula-delta classification over batches of raw and normalized
//! formula pairs. Unsupported syntax returns a per-row fallback marker; it is
//! never guessed. Python remains authoritative for profile and finding logic.

use pyo3::prelude::*;
use sha2::{Digest, Sha256};
use std::collections::{HashMap, HashSet};

const ERROR_CODES: [&str; 8] = [
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
];
const MAX_FORMULA_BYTES: usize = 32_768;
const MAX_NESTING: usize = 256;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TokenType {
    Operand,
    Func,
    Array,
    Paren,
    Sep,
    OperatorPrefix,
    OperatorInfix,
    OperatorPostfix,
    WhiteSpace,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Subtype {
    None,
    Text,
    Number,
    Logical,
    Error,
    Range,
    Open,
    Close,
    Arg,
    Row,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Token {
    value: String,
    token_type: TokenType,
    subtype: Subtype,
}

impl Token {
    fn operand(value: String) -> Self {
        let subtype = if value.starts_with('"') {
            Subtype::Text
        } else if value.starts_with('#') {
            Subtype::Error
        } else if matches!(value.as_str(), "TRUE" | "FALSE") {
            Subtype::Logical
        } else if value.parse::<f64>().is_ok() {
            Subtype::Number
        } else {
            Subtype::Range
        };
        Self {
            value,
            token_type: TokenType::Operand,
            subtype,
        }
    }

    fn opener(value: String, token_type: TokenType) -> Self {
        Self {
            value,
            token_type,
            subtype: Subtype::Open,
        }
    }

    fn closer(token_type: TokenType) -> Self {
        Self {
            value: if token_type == TokenType::Array {
                "}".to_string()
            } else {
                ")".to_string()
            },
            token_type,
            subtype: Subtype::Close,
        }
    }
}

fn is_scientific_prefix(value: &str) -> bool {
    let Some(prefix) = value.strip_suffix(['E', 'e']) else {
        return false;
    };
    let mut parts = prefix.split('.');
    let Some(integer) = parts.next() else {
        return false;
    };
    if integer.len() != 1 || !matches!(integer.as_bytes()[0], b'1'..=b'9') {
        return false;
    }
    match parts.next() {
        None => true,
        Some(fraction) => {
            !fraction.is_empty()
                && fraction.bytes().all(|byte| byte.is_ascii_digit())
                && parts.next().is_none()
        }
    }
}

fn save_operand(buffer: &mut String, tokens: &mut Vec<Token>) {
    if !buffer.is_empty() {
        tokens.push(Token::operand(std::mem::take(buffer)));
    }
}

fn quoted_end(formula: &[u8], start: usize, delimiter: u8) -> Option<usize> {
    let mut index = start + 1;
    while index < formula.len() {
        if formula[index] != delimiter {
            index += 1;
            continue;
        }
        if index + 1 < formula.len() && formula[index + 1] == delimiter {
            index += 2;
            continue;
        }
        return Some(index + 1);
    }
    None
}

fn tokenize(formula: &str) -> Result<Vec<Token>, ()> {
    if formula.len() > MAX_FORMULA_BYTES || !formula.is_ascii() || !formula.starts_with('=') {
        return Err(());
    }
    let bytes = formula.as_bytes();
    let mut tokens: Vec<Token> = Vec::new();
    let mut stack: Vec<TokenType> = Vec::new();
    let mut buffer = String::new();
    let mut index = 1usize;
    while index < bytes.len() {
        let byte = bytes[index];
        if matches!(byte, b'+' | b'-') && is_scientific_prefix(&buffer) {
            buffer.push(byte as char);
            index += 1;
            continue;
        }
        if b",;}) +-*/^&=><%".contains(&byte) {
            save_operand(&mut buffer, &mut tokens);
        }
        match byte {
            b'"' | b'\'' => {
                if !buffer.is_empty() && !buffer.ends_with(':') {
                    return Err(());
                }
                let end = quoted_end(bytes, index, byte).ok_or(())?;
                let value = formula[index..end].to_string();
                if byte == b'"' {
                    tokens.push(Token::operand(value));
                } else {
                    buffer.push_str(&value);
                }
                index = end;
            }
            b'[' => {
                let mut depth = 0i32;
                let mut end = None;
                for (offset, candidate) in bytes[index..].iter().enumerate() {
                    match candidate {
                        b'[' => depth += 1,
                        b']' => {
                            depth -= 1;
                            if depth == 0 {
                                end = Some(index + offset + 1);
                                break;
                            }
                        }
                        _ => {}
                    }
                }
                let end = end.ok_or(())?;
                buffer.push_str(&formula[index..end]);
                index = end;
            }
            b'#' => {
                if !buffer.is_empty() && !buffer.ends_with('!') {
                    // Dynamic spill syntax is left to the Python fallback.
                    return Err(());
                }
                let rest = &formula[index..];
                let error = ERROR_CODES
                    .iter()
                    .find(|candidate| rest.starts_with(**candidate))
                    .ok_or(())?;
                buffer.push_str(error);
                save_operand(&mut buffer, &mut tokens);
                index += error.len();
            }
            b'@' => return Err(()),
            b' ' | b'\n' => {
                tokens.push(Token {
                    value: (byte as char).to_string(),
                    token_type: TokenType::WhiteSpace,
                    subtype: Subtype::None,
                });
                index += 1;
                while index < bytes.len() && matches!(bytes[index], b' ' | b'\n') {
                    index += 1;
                }
            }
            b'+' | b'-' | b'*' | b'/' | b'^' | b'&' | b'=' | b'>' | b'<' | b'%' => {
                let (value, consumed) = if index + 1 < bytes.len()
                    && matches!(&formula[index..index + 2], ">=" | "<=" | "<>")
                {
                    (formula[index..index + 2].to_string(), 2)
                } else {
                    ((byte as char).to_string(), 1)
                };
                let token_type = if byte == b'%' {
                    TokenType::OperatorPostfix
                } else if !matches!(byte, b'+' | b'-') {
                    TokenType::OperatorInfix
                } else {
                    let previous = tokens
                        .iter()
                        .rev()
                        .find(|token| token.token_type != TokenType::WhiteSpace);
                    if previous.is_some_and(|token| {
                        token.subtype == Subtype::Close
                            || token.token_type == TokenType::OperatorPostfix
                            || token.token_type == TokenType::Operand
                    }) {
                        TokenType::OperatorInfix
                    } else {
                        TokenType::OperatorPrefix
                    }
                };
                tokens.push(Token {
                    value,
                    token_type,
                    subtype: Subtype::None,
                });
                index += consumed;
            }
            b'{' => {
                if !buffer.is_empty() || stack.len() >= MAX_NESTING {
                    return Err(());
                }
                tokens.push(Token::opener("{".to_string(), TokenType::Array));
                stack.push(TokenType::Array);
                index += 1;
            }
            b'(' => {
                if stack.len() >= MAX_NESTING {
                    return Err(());
                }
                let (value, token_type) = if buffer.is_empty() {
                    ("(".to_string(), TokenType::Paren)
                } else {
                    (format!("{}(", std::mem::take(&mut buffer)), TokenType::Func)
                };
                tokens.push(Token::opener(value, token_type));
                stack.push(token_type);
                index += 1;
            }
            b')' | b'}' => {
                let token_type = stack.pop().ok_or(())?;
                let closer = Token::closer(token_type);
                if closer.value.as_bytes()[0] != byte {
                    return Err(());
                }
                tokens.push(closer);
                index += 1;
            }
            b';' | b',' => {
                if byte == b';' {
                    tokens.push(Token {
                        value: ";".to_string(),
                        token_type: TokenType::Sep,
                        subtype: Subtype::Row,
                    });
                } else if stack.last().is_none_or(|kind| *kind == TokenType::Paren) {
                    tokens.push(Token {
                        value: ",".to_string(),
                        token_type: TokenType::OperatorInfix,
                        subtype: Subtype::None,
                    });
                } else {
                    tokens.push(Token {
                        value: ",".to_string(),
                        token_type: TokenType::Sep,
                        subtype: Subtype::Arg,
                    });
                }
                index += 1;
            }
            _ => {
                buffer.push(byte as char);
                index += 1;
            }
        }
    }
    if !stack.is_empty() {
        return Err(());
    }
    save_operand(&mut buffer, &mut tokens);
    Ok(tokens)
}

fn split_range(value: &str) -> Option<(&str, &str, &str)> {
    let (sheet, reference) = value
        .rsplit_once('!')
        .map_or(("", value), |(sheet, reference)| (sheet, reference));
    let (start, end) = reference.split_once(':')?;
    Some((sheet, start, end))
}

fn endpoint(value: &str) -> Option<(&str, u32)> {
    let value = value.strip_prefix('$').unwrap_or(value);
    let column_end = value
        .bytes()
        .position(|byte| byte == b'$' || byte.is_ascii_digit())?;
    let column = &value[..column_end];
    if column.is_empty()
        || column.len() > 3
        || !column.bytes().all(|byte| byte.is_ascii_uppercase())
    {
        return None;
    }
    let row = value[column_end..]
        .strip_prefix('$')
        .unwrap_or(&value[column_end..]);
    if row.is_empty() || !row.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    Some((column, row.parse().ok()?))
}

fn is_range_extension(base: &str, current: &str) -> bool {
    let Some((base_sheet, base_start, base_end)) = split_range(base) else {
        return false;
    };
    let Some((current_sheet, current_start, current_end)) = split_range(current) else {
        return false;
    };
    if base_sheet != current_sheet || base_start != current_start {
        return false;
    }
    let Some((base_column, base_row)) = endpoint(base_end) else {
        return false;
    };
    let Some((current_column, current_row)) = endpoint(current_end) else {
        return false;
    };
    if base_column == current_column && current_row >= base_row {
        return true;
    }
    base_row == current_row
        && (current_column.len() > base_column.len()
            || (current_column.len() == base_column.len() && current_column >= base_column))
}

fn differs_only_by_extension(base: &[Token], current: &[Token]) -> bool {
    if base.len() != current.len() {
        return false;
    }
    let mut extension_seen = false;
    for (base_token, current_token) in base.iter().zip(current) {
        if base_token.token_type != current_token.token_type
            || base_token.subtype != current_token.subtype
        {
            return false;
        }
        if base_token.value == current_token.value {
            continue;
        }
        if base_token.token_type != TokenType::Operand
            || base_token.subtype != Subtype::Range
            || !is_range_extension(&base_token.value, &current_token.value)
        {
            return false;
        }
        extension_seen = true;
    }
    extension_seen
}

fn shape_label(token: &Token) -> String {
    if token.token_type == TokenType::Operand && token.subtype == Subtype::Range {
        "REF".to_string()
    } else if token.token_type == TokenType::Operand && token.subtype == Subtype::Number {
        "NUM".to_string()
    } else if token.token_type == TokenType::Operand && token.subtype == Subtype::Text {
        "TEXT".to_string()
    } else {
        token.value.to_lowercase()
    }
}

fn boundary_open(token: Option<&Token>) -> bool {
    token.is_none_or(|token| {
        (matches!(token.token_type, TokenType::Func | TokenType::Paren)
            && token.subtype == Subtype::Open)
            || (token.token_type == TokenType::Sep && token.subtype == Subtype::Arg)
    })
}

fn boundary_close(token: Option<&Token>) -> bool {
    token.is_none_or(|token| {
        (matches!(token.token_type, TokenType::Func | TokenType::Paren)
            && token.subtype == Subtype::Close)
            || (token.token_type == TokenType::Sep && token.subtype == Subtype::Arg)
    })
}

fn find_occurrence(outer: &[Token], inner: &[String], shape: bool) -> Option<usize> {
    if inner.is_empty() || inner.len() >= outer.len() {
        return None;
    }
    let outer_keys: Vec<String> = outer
        .iter()
        .map(|token| {
            if shape {
                shape_label(token)
            } else {
                token.value.clone()
            }
        })
        .collect();
    for start in 0..=outer.len() - inner.len() {
        if outer_keys[start..start + inner.len()] != *inner {
            continue;
        }
        if boundary_open(start.checked_sub(1).map(|index| &outer[index]))
            && boundary_close(outer.get(start + inner.len()))
        {
            return Some(start);
        }
    }
    None
}

fn meaningful_core(tokens: &[Token]) -> bool {
    tokens.iter().any(|token| {
        (token.token_type == TokenType::Operand && token.subtype == Subtype::Range)
            || (token.token_type == TokenType::Func && token.subtype == Subtype::Open)
    })
}

fn structural_core(tokens: &[Token]) -> bool {
    tokens
        .iter()
        .any(|token| token.token_type == TokenType::Func && token.subtype == Subtype::Open)
        || tokens
            .iter()
            .filter(|token| token.token_type == TokenType::Operand)
            .count()
            >= 2
}

fn skeleton_key(outer: &[Token], start: usize, span: usize) -> String {
    let mut labels: Vec<String> = outer[..start].iter().map(shape_label).collect();
    labels.push("<CORE>".to_string());
    labels.extend(outer[start + span..].iter().map(shape_label));
    let digest = Sha256::digest(labels.join("|").as_bytes());
    format!("{digest:x}")[..12].to_string()
}

#[derive(Clone, Debug)]
struct WrapperMatch {
    kind: &'static str,
    exact: bool,
    skeleton: String,
}

fn detect_wrapper(base: &[Token], current: &[Token]) -> Option<WrapperMatch> {
    let candidates = [("wrapped", current, base), ("unwrapped", base, current)];
    for (kind, outer, inner) in candidates {
        if !meaningful_core(inner) {
            continue;
        }
        let keys: Vec<String> = inner.iter().map(|token| token.value.clone()).collect();
        if let Some(start) = find_occurrence(outer, &keys, false) {
            return Some(WrapperMatch {
                kind,
                exact: true,
                skeleton: skeleton_key(outer, start, inner.len()),
            });
        }
    }
    for (kind, outer, inner) in candidates {
        if !meaningful_core(inner) || !structural_core(inner) {
            continue;
        }
        let keys: Vec<String> = inner.iter().map(shape_label).collect();
        if let Some(start) = find_occurrence(outer, &keys, true) {
            return Some(WrapperMatch {
                kind,
                exact: false,
                skeleton: skeleton_key(outer, start, inner.len()),
            });
        }
    }
    None
}

fn function_closes(tokens: &[Token]) -> Result<HashMap<usize, usize>, ()> {
    let mut stack = Vec::new();
    let mut closes = HashMap::new();
    for (position, token) in tokens.iter().enumerate() {
        if token.token_type == TokenType::Func && token.subtype == Subtype::Open {
            let name = token
                .value
                .strip_suffix('(')
                .unwrap_or(&token.value)
                .rsplit('.')
                .next()
                .unwrap_or("")
                .to_ascii_lowercase();
            if matches!(name.as_str(), "let" | "lambda") {
                return Err(());
            }
            stack.push(position);
        } else if token.token_type == TokenType::Func && token.subtype == Subtype::Close {
            closes.insert(stack.pop().ok_or(())?, position);
        }
    }
    if stack.is_empty() {
        Ok(closes)
    } else {
        Err(())
    }
}

fn collect_references(tokens: &[Token]) -> Result<HashSet<String>, ()> {
    let closes = function_closes(tokens)?;
    fn walk(
        tokens: &[Token],
        closes: &HashMap<usize, usize>,
        start: usize,
        end: usize,
        dynamic: Option<&str>,
        output: &mut HashSet<String>,
    ) -> Result<(), ()> {
        let mut position = start;
        while position < end {
            let token = &tokens[position];
            if token.token_type == TokenType::Func && token.subtype == Subtype::Open {
                let close = *closes.get(&position).ok_or(())?;
                if close >= end {
                    return Err(());
                }
                let name = token
                    .value
                    .strip_suffix('(')
                    .unwrap_or(&token.value)
                    .rsplit('.')
                    .next()
                    .unwrap_or("")
                    .to_ascii_lowercase();
                let function_dynamic = match name.as_str() {
                    "anchorarray" => Some("spill"),
                    "single" => Some("implicit"),
                    _ => None,
                };
                walk(
                    tokens,
                    closes,
                    position + 1,
                    close,
                    function_dynamic,
                    output,
                )?;
                position = close + 1;
                continue;
            }
            if token.token_type == TokenType::Operand && token.subtype == Subtype::Range {
                let value = match dynamic {
                    Some("spill") => format!("{}#", token.value),
                    Some("implicit") => format!("@{}", token.value),
                    _ => token.value.clone(),
                };
                output.insert(value.to_ascii_lowercase());
            }
            position += 1;
        }
        Ok(())
    }
    let mut output = HashSet::new();
    walk(tokens, &closes, 0, tokens.len(), None, &mut output)?;
    Ok(output)
}

type BatchInput = (String, String, String, String);
type BatchOutput = (bool, bool, Option<String>, Option<bool>, String, bool);

fn classify_batch_inner(pairs: Vec<BatchInput>) -> Vec<BatchOutput> {
    let mut wrapper_cache: HashMap<(String, String), Option<WrapperMatch>> = HashMap::new();
    pairs
        .into_iter()
        .map(|(base_raw, current_raw, base_norm, current_norm)| {
            let Ok(base_raw_tokens) = tokenize(&base_raw) else {
                return (false, false, None, None, String::new(), false);
            };
            let Ok(current_raw_tokens) = tokenize(&current_raw) else {
                return (false, false, None, None, String::new(), false);
            };
            let Ok(base_norm_tokens) = tokenize(&base_norm) else {
                return (false, false, None, None, String::new(), false);
            };
            let Ok(current_norm_tokens) = tokenize(&current_norm) else {
                return (false, false, None, None, String::new(), false);
            };
            let Ok(base_references) = collect_references(&base_raw_tokens) else {
                return (false, false, None, None, String::new(), false);
            };
            let Ok(current_references) = collect_references(&current_raw_tokens) else {
                return (false, false, None, None, String::new(), false);
            };
            let expected = differs_only_by_extension(&base_raw_tokens, &current_raw_tokens);
            let wrapper = if expected {
                None
            } else {
                wrapper_cache
                    .entry((base_norm.clone(), current_norm.clone()))
                    .or_insert_with(|| detect_wrapper(&base_norm_tokens, &current_norm_tokens))
                    .clone()
            };
            let added_reference = current_references
                .iter()
                .any(|reference| !base_references.contains(reference));
            match wrapper {
                Some(wrapper) => {
                    let event = format!("formula-wrapper:{}:{}", wrapper.kind, wrapper.skeleton);
                    (
                        true,
                        expected,
                        Some(wrapper.kind.to_string()),
                        Some(wrapper.exact),
                        event,
                        added_reference,
                    )
                }
                None => (true, expected, None, None, String::new(), added_reference),
            }
        })
        .collect()
}

/// One GIL-detached call over a bounded batch. Each input tuple is
/// `(baseline_raw, current_raw, baseline_r1c1, current_r1c1)`; each output is
/// `(supported, expected_extension, wrapper_kind, wrapper_exact, event_key,
/// added_reference)`.
#[pyfunction]
pub fn formula_delta_batch(py: Python<'_>, pairs: Vec<BatchInput>) -> Vec<BatchOutput> {
    py.detach(|| classify_batch_inner(pairs))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tokenizes_core_formula_shapes() {
        let tokens = tokenize("=SUM(C2:C10)").unwrap();
        assert_eq!(tokens.len(), 3);
        assert_eq!(tokens[0].token_type, TokenType::Func);
        assert_eq!(tokens[1].subtype, Subtype::Range);
        assert_eq!(tokens[1].value, "C2:C10");
    }

    #[test]
    fn classifies_extension_wrapper_and_added_reference() {
        let rows = classify_batch_inner(vec![
            (
                "=SUM(C2:C10)".into(),
                "=SUM(C2:C15)".into(),
                "=SUM(RC:R[8]C)".into(),
                "=SUM(RC:R[13]C)".into(),
            ),
            (
                "=IF(R[1]C[-2],RC[3],NA())".into(),
                "=IF(R1C8,C[3],IF(R[1]C[-2],RC[3],NA()),NA())".into(),
                "=IF(R[1]C[-2],RC[3],NA())".into(),
                "=IF(R1C8,C[3],IF(R[1]C[-2],RC[3],NA()),NA())".into(),
            ),
            (
                "=B5".into(),
                "=B5+C5".into(),
                "=RC".into(),
                "=RC+RC[1]".into(),
            ),
        ]);
        assert!(rows[0].0 && rows[0].1);
        assert_eq!(rows[1].2.as_deref(), Some("wrapped"));
        assert_eq!(rows[1].3, Some(true));
        assert!(rows[2].5);
    }

    #[test]
    fn unsupported_syntax_falls_back() {
        let rows = classify_batch_inner(vec![(
            "=LET(x,A1,x+1)".into(),
            "=LET(x,A1,x+2)".into(),
            "=LET(x,RC[-1],x+1)".into(),
            "=LET(x,RC[-1],x+2)".into(),
        )]);
        assert!(!rows[0].0);
    }

    #[test]
    fn oversized_or_excessively_nested_input_falls_back() {
        let oversized = format!("={}", "A".repeat(MAX_FORMULA_BYTES));
        let nested = format!(
            "={}A1{}",
            "F(".repeat(MAX_NESTING + 1),
            ")".repeat(MAX_NESTING + 1)
        );
        let rows = classify_batch_inner(vec![
            (oversized.clone(), "=A1".into(), oversized, "=RC".into()),
            (nested.clone(), "=A1".into(), nested, "=RC".into()),
        ]);
        assert!(rows.iter().all(|row| !row.0));
    }
}
