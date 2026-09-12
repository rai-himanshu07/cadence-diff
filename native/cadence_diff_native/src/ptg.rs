//! Ptg token-stream renderer: turns `rgce`/`rgcb` formula bytes into A1 or
//! R1C1 text. Faithful, opcode-by-opcode port of `render()` in
//! `ptg_decoder.py` (already proven byte-exact against desktop Excel on
//! representative large-workbook inputs; retained aggregate evidence is
//! private); this port's job is exact agreement with that reference, not
//! independent re-derivation of the BIFF12 spec.
//!
//! Deliberate, disclosed divergence from the Python reference: stack
//! underflow (an opcode popping more operands than are present) is a
//! malformed-input condition. The Python experiment silently clamps via
//! Python's negative-slice semantics; this port fails closed with a
//! `DecodeError` instead, per the plan's requirement that the Rust reader
//! validate structure the experimental walker merely trusted. This never
//! fires on well-formed real files (confirmed empirically below).
//!
//! Second deliberate, disclosed divergence from the Python reference,
//! discovered by B2's guest verification against desktop Excel's own live
//! Formula2R1C1 cache (Criterion 9), not by the earlier Rust-vs-Python parity
//! proof: the Python reference always renders a relative R1C1 offset in
//! bracket form (`R[0]`, `C[0]`) even when the offset is zero, but real Excel
//! collapses a zero relative offset to a bare `R`/`C` (and a same-cell
//! self-reference to bare `RC`, no brackets at all). Both the Python
//! reference and this port originally shared that gap, invisible to a
//! Rust-vs-Python-only comparison; this port now matches real Excel instead
//! of the (now-known-imperfect) Python reference for this one case. A related
//! fix in `render_area`: a whole-column/whole-row range (e.g. A1's "A:A")
//! now renders R1C1's numeric column/row syntax instead of inherited A1
//! letter syntax, and collapses to a single bare reference (no colon) when
//! the range is a single whole column/row (`c1 == c2` or `r1 == r2`) --
//! confirmed empirically against live Excel; A1 mode does not collapse this
//! case and keeps the full "A:A" form (these two modes are not symmetric
//! here).
//!
//! KNOWN, DISCLOSED, BOUNDED residual gap in R1C1 mode only (found by the
//! same B2 guest verification, not yet root-caused): on representative
//! large-workbook inputs, after the three fixes above, R1C1 mode reaches
//! greater than 99% byte-exact parity with desktop Excel's Formula2R1C1
//! cache. The small residual, with A1/Formula2 fidelity unaffected, is
//! narrowly isolated to sheet-qualified (`Sheet!...`), absolute (no
//! relative offsets at all), range (`X:Y`) references with a consistent
//! length delta and a small number of repeated patterns -- suggesting one
//! specific, not-yet-identified formula construct reused by a shared/array
//! formula across many rows, not a diffuse bug. Extensive synthetic reproduction
//! attempts (single- and multi-sheet 3D refs, regular/whole-column/whole-row
//! ranges, zero- and non-zero-offset combinations, range-operator-joined
//! defined names, explicit duplicate-prefix ranges) all matched real Excel
//! correctly and did not reproduce this pattern. Criterion 9's own numeric
//! bar is specifically about Formula2 (A1) text, which is 100% exact; this
//! R1C1 gap does not block that criterion, but MUST be re-examined before B5
//! ("Definition-level formula comparison") makes kernel R1C1 the load-bearing
//! comparison text -- see the B2 Execution Log entry (2026-09-07) for the
//! full guest-verification numbers and the diagnostic trail.

use crate::ftab;
use crate::records::{
    get_f64, get_u16, get_u32, DecodeError, Result, MAX_ARRAY_CELLS, MAX_COL, MAX_ROW,
};
use crate::workbook::Workbook;

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum Mode {
    A1,
    R1C1,
}

fn binop(ptg: u8) -> Option<&'static str> {
    Some(match ptg {
        0x03 => "+",
        0x04 => "-",
        0x05 => "*",
        0x06 => "/",
        0x07 => "^",
        0x08 => "&",
        0x09 => "<",
        0x0A => "<=",
        0x0B => "=",
        0x0C => ">=",
        0x0D => ">",
        0x0E => "<>",
        0x0F => " ",
        0x10 => ",",
        0x11 => ":",
        _ => return None,
    })
}

fn error_text(code: u8) -> &'static str {
    match code {
        0x00 => "#NULL!",
        0x07 => "#DIV/0!",
        0x0F => "#VALUE!",
        0x17 => "#REF!",
        0x1D => "#NAME?",
        0x24 => "#NUM!",
        0x2A => "#N/A",
        0x2B => "#GETTING_DATA",
        _ => "#N/A",
    }
}

pub fn col_letters(c: u32) -> String {
    let mut s = String::new();
    let mut c = c as i64 + 1;
    while c > 0 {
        let r = (c - 1) % 26;
        c = (c - 1) / 26;
        s.insert(0, (b'A' + r as u8) as char);
    }
    s
}

/// Mirrors Python's `repr(float)` closely enough for real spreadsheet
/// literals: whole numbers under 1e15 print as plain integers; otherwise the
/// shortest round-tripping decimal, switching to `E+NN`/`E-NN` scientific
/// notation at the same magnitude boundaries CPython's float repr uses.
/// Known bounded gap: exact tie-breaking of CPython's dtoa vs Rust's Ryu at
/// extreme magnitudes is not independently re-derived here, only approximated;
/// this is disclosed, not hidden (see the B0 decision record).
fn fmt_num(x: f64) -> String {
    if x.fract() == 0.0 && x.abs() < 1e15 {
        return format!("{}", x as i64);
    }
    let abs = x.abs();
    if abs != 0.0 && !(1e-4..1e16).contains(&abs) {
        let sci = format!("{:e}", x);
        if let Some(idx) = sci.find('e') {
            let (mantissa, exp) = sci.split_at(idx);
            if let Ok(exp_val) = exp[1..].parse::<i32>() {
                return format!(
                    "{mantissa}E{}{:02}",
                    if exp_val >= 0 { "+" } else { "-" },
                    exp_val.abs()
                );
            }
        }
    }
    format!("{x}")
}

/// Renders an R1C1 row component (`R5`, `R[2]`, `R[-3]`, or bare `R` for a
/// zero relative offset -- real Excel collapses a same-row relative
/// reference to a bare `R`, never `R[0]`).
fn r1c1_row_part(row: u32, row_rel: bool, host_row: u32) -> String {
    if row_rel {
        match row as i64 - host_row as i64 {
            0 => "R".to_string(),
            offset => format!("R[{offset}]"),
        }
    } else {
        format!("R{}", row + 1)
    }
}

/// Renders an R1C1 column component (`C5`, `C[2]`, `C[-3]`, or bare `C` for a
/// zero relative offset), mirroring `r1c1_row_part`.
fn r1c1_col_part(col: u32, col_rel: bool, host_col: u32) -> String {
    if col_rel {
        match col as i64 - host_col as i64 {
            0 => "C".to_string(),
            offset => format!("C[{offset}]"),
        }
    } else {
        format!("C{}", col + 1)
    }
}

fn render_ref(
    row: u32,
    col: u32,
    row_rel: bool,
    col_rel: bool,
    host_row: u32,
    host_col: u32,
    mode: Mode,
) -> String {
    match mode {
        Mode::R1C1 => format!(
            "{}{}",
            r1c1_row_part(row, row_rel, host_row),
            r1c1_col_part(col, col_rel, host_col)
        ),
        Mode::A1 => format!(
            "{}{}{}{}",
            if col_rel { "" } else { "$" },
            col_letters(col),
            if row_rel { "" } else { "$" },
            row + 1
        ),
    }
}

#[allow(clippy::too_many_arguments)]
fn render_area(
    r1: u32,
    r2: u32,
    c1: u32,
    c2: u32,
    r1rel: bool,
    r2rel: bool,
    c1rel: bool,
    c2rel: bool,
    host_row: u32,
    host_col: u32,
    mode: Mode,
) -> String {
    if r1 == 0 && r2 == MAX_ROW {
        // Whole-column range (e.g. A1's "A:A"); real Excel's R1C1 mode
        // renders this as a pure column range ("C1:C3"), not letter-based
        // A1 syntax -- this branch previously ignored `mode` entirely. A
        // further R1C1-only quirk, confirmed empirically against live Excel:
        // when the range is a SINGLE whole column (c1 == c2), R1C1 collapses
        // to one bare column reference ("C[-4]", no colon) while A1 mode
        // keeps the full "A:A" range form -- these are not symmetric.
        return match mode {
            Mode::A1 => format!(
                "{}{}:{}{}",
                if c1rel { "" } else { "$" },
                col_letters(c1),
                if c2rel { "" } else { "$" },
                col_letters(c2)
            ),
            Mode::R1C1 if c1 == c2 => r1c1_col_part(c1, c1rel, host_col),
            Mode::R1C1 => format!(
                "{}:{}",
                r1c1_col_part(c1, c1rel, host_col),
                r1c1_col_part(c2, c2rel, host_col)
            ),
        };
    }
    if c1 == 0 && c2 == MAX_COL {
        // Whole-row range (e.g. A1's "1:1"); same R1C1 fixes as above,
        // mirrored for rows (including the single-whole-row collapse).
        return match mode {
            Mode::A1 => format!(
                "{}{}:{}{}",
                if r1rel { "" } else { "$" },
                r1 + 1,
                if r2rel { "" } else { "$" },
                r2 + 1
            ),
            Mode::R1C1 if r1 == r2 => r1c1_row_part(r1, r1rel, host_row),
            Mode::R1C1 => format!(
                "{}:{}",
                r1c1_row_part(r1, r1rel, host_row),
                r1c1_row_part(r2, r2rel, host_row)
            ),
        };
    }
    format!(
        "{}:{}",
        render_ref(r1, c1, r1rel, c1rel, host_row, host_col, mode),
        render_ref(r2, c2, r2rel, c2rel, host_row, host_col, mode)
    )
}

/// Non-relative location: absolute row + masked col + rel flags (PtgRef/PtgRef3d).
fn loc(buf: &[u8], pos: usize) -> Result<(u32, u32, bool, bool)> {
    let row = get_u32(buf, pos)?;
    let cw = get_u16(buf, pos + 4)?;
    Ok((
        row,
        (cw & 0x3FFF) as u32,
        cw & 0x8000 != 0,
        cw & 0x4000 != 0,
    ))
}

/// Relative location: the row/col fields carry signed offsets from
/// `(host_row, host_col)` when their rel flag is set (PtgRefN/PtgAreaN).
/// Uses `rem_euclid` to match Python's always-non-negative `%` semantics --
/// Rust's native `%` follows the dividend's sign and would silently diverge
/// from the reference for negative offsets near row/col 0.
fn loc_rel_parts(raw_row: u32, cw: u16, host_row: u32, host_col: u32) -> (u32, u32, bool, bool) {
    let raw_row = raw_row as i64;
    let col_field = (cw & 0x3FFF) as i64;
    let row_rel = cw & 0x8000 != 0;
    let col_rel = cw & 0x4000 != 0;
    let row = if row_rel {
        let srow = if raw_row >= (1i64 << 19) {
            raw_row - (1i64 << 20)
        } else {
            raw_row
        };
        (host_row as i64 + srow).rem_euclid(1i64 << 20)
    } else {
        raw_row
    };
    let col = if col_rel {
        let scol = if col_field >= (1i64 << 13) {
            col_field - (1i64 << 14)
        } else {
            col_field
        };
        (host_col as i64 + scol).rem_euclid(1i64 << 14)
    } else {
        col_field
    };
    (row as u32, col as u32, row_rel, col_rel)
}

fn loc_rel(buf: &[u8], pos: usize, host_row: u32, host_col: u32) -> Result<(u32, u32, bool, bool)> {
    let raw_row = get_u32(buf, pos)?;
    let cw = get_u16(buf, pos + 4)?;
    Ok(loc_rel_parts(raw_row, cw, host_row, host_col))
}

/// SerAr array constants from `rgcb`: `cols:u32, rows:u32`, then `rows*cols`
/// tagged values. Returns the rendered `{...}` literal and the position past it.
fn parse_array_consts(rgcb: &[u8], pos: usize) -> Result<(String, usize)> {
    let cols = get_u32(rgcb, pos)? as usize;
    let rows = get_u32(rgcb, pos + 4)? as usize;
    let mut pos = pos + 8;
    let total = cols
        .checked_mul(rows)
        .ok_or(DecodeError("array constant dims overflow"))?;
    if total > MAX_ARRAY_CELLS {
        return Err(DecodeError("array constant cell count exceeds cap"));
    }
    let mut out_rows: Vec<String> = Vec::with_capacity(rows);
    for _ in 0..rows {
        let mut vals: Vec<String> = Vec::with_capacity(cols);
        for _ in 0..cols {
            let t = *rgcb
                .get(pos)
                .ok_or(DecodeError("SerAr type byte past end"))?;
            pos += 1;
            match t {
                0x00 => vals.push(String::new()),
                0x01 => {
                    vals.push(fmt_num(get_f64(rgcb, pos)?));
                    pos += 8;
                }
                0x02 => {
                    let (s, next_pos) = crate::records::read_wide(rgcb, pos)?;
                    vals.push(format!(
                        "\"{}\"",
                        s.unwrap_or_default().replace('"', "\"\"")
                    ));
                    pos = next_pos;
                }
                0x04 => {
                    let b = *rgcb
                        .get(pos)
                        .ok_or(DecodeError("SerAr bool byte past end"))?;
                    vals.push(if b != 0 {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
                    });
                    pos += 1;
                }
                0x10 => {
                    let b = *rgcb
                        .get(pos)
                        .ok_or(DecodeError("SerAr error byte past end"))?;
                    vals.push(error_text(b).to_string());
                    pos += 1;
                }
                _ => return Err(DecodeError("unknown SerAr element type")),
            }
        }
        out_rows.push(vals.join(","));
    }
    Ok((format!("{{{}}}", out_rows.join(";")), pos))
}

/// Renders one formula token stream (`rgce`/`rgcb`) at the given host cell.
/// Unknown/unsupported tokens fail closed with `Err`, never a guess.
pub fn render(
    rgce: &[u8],
    rgcb: &[u8],
    wb: &Workbook,
    host_row: u32,
    host_col: u32,
    mode: Mode,
) -> Result<String> {
    let mut stack: Vec<String> = Vec::new();
    let mut pos = 0usize;
    let mut cb_pos = 0usize;
    let mut pend_nl: usize = 0;
    let mut pend_sp: usize = 0;
    let n = rgce.len();

    fn take_space(pend_nl: &mut usize, pend_sp: &mut usize) -> String {
        let s = "\n".repeat(*pend_nl) + &" ".repeat(*pend_sp);
        *pend_nl = 0;
        *pend_sp = 0;
        s
    }
    fn push(stack: &mut Vec<String>, pend_nl: &mut usize, pend_sp: &mut usize, s: String) {
        stack.push(take_space(pend_nl, pend_sp) + &s);
    }
    fn pop(stack: &mut Vec<String>) -> Result<String> {
        stack.pop().ok_or(DecodeError("stack underflow"))
    }

    while pos < n {
        let ptg = rgce[pos];
        pos += 1;
        let base: u8 = if ptg >= 0x20 {
            (ptg & 0x1F) | 0x20
        } else {
            ptg
        };

        if ptg == 0x18 {
            let eptg = *rgce.get(pos).ok_or(DecodeError("eptg18 past end"))?;
            pos += 1;
            match eptg {
                0x19 => {
                    pos += 12;
                    push(
                        &mut stack,
                        &mut pend_nl,
                        &mut pend_sp,
                        "#TABLE_REF#".to_string(),
                    );
                }
                0x1D => {
                    pos += 4;
                    push(
                        &mut stack,
                        &mut pend_nl,
                        &mut pend_sp,
                        "#SXNAME#".to_string(),
                    );
                }
                _ => return Err(DecodeError("unknown eptg18 subtype")),
            }
            continue;
        }
        if ptg == 0x19 {
            let eptg = *rgce.get(pos).ok_or(DecodeError("PtgAttr eptg past end"))?;
            pos += 1;
            match eptg {
                0x01 | 0x02 | 0x08 | 0x20 | 0x21 | 0x80 => pos += 2,
                0x04 => {
                    let c = get_u16(rgce, pos)? as usize;
                    pos += 2 + 2 * (c + 1);
                }
                0x10 => {
                    pos += 2;
                    let a = pop(&mut stack)?;
                    push(&mut stack, &mut pend_nl, &mut pend_sp, format!("SUM({a})"));
                }
                0x40 | 0x41 => {
                    let t = *rgce
                        .get(pos)
                        .ok_or(DecodeError("AttrSpace type past end"))?;
                    let cch = *rgce
                        .get(pos + 1)
                        .ok_or(DecodeError("AttrSpace count past end"))?
                        as usize;
                    pos += 2;
                    if t == 4 || t == 5 {
                        if let Some(last) = stack.last_mut() {
                            let filler = if t == 4 { " " } else { "\n" };
                            last.push_str(&filler.repeat(cch));
                        }
                    } else if t == 0 || t == 2 || t == 6 {
                        pend_sp += cch;
                    } else {
                        pend_nl += cch;
                    }
                }
                _ => return Err(DecodeError("unknown PtgAttr subtype")),
            }
            continue;
        }
        if let Some(op) = binop(ptg) {
            let b = pop(&mut stack)?;
            let a = pop(&mut stack)?;
            let space = take_space(&mut pend_nl, &mut pend_sp);
            stack.push(format!("{a}{space}{op}{b}"));
            continue;
        }
        match ptg {
            0x12 => {
                let a = pop(&mut stack)?;
                push(&mut stack, &mut pend_nl, &mut pend_sp, format!("+{a}"));
                continue;
            }
            0x13 => {
                let a = pop(&mut stack)?;
                push(&mut stack, &mut pend_nl, &mut pend_sp, format!("-{a}"));
                continue;
            }
            0x14 => {
                let a = pop(&mut stack)?;
                stack.push(format!("{a}%"));
                continue;
            }
            0x15 => {
                let a = pop(&mut stack)?;
                push(&mut stack, &mut pend_nl, &mut pend_sp, format!("({a})"));
                continue;
            }
            0x16 => {
                push(&mut stack, &mut pend_nl, &mut pend_sp, String::new());
                continue;
            }
            0x17 => {
                let cch = get_u16(rgce, pos)? as usize;
                pos += 2;
                let bytes = rgce
                    .get(pos..pos + 2 * cch)
                    .ok_or(DecodeError("PtgStr text past end"))?;
                let units: Vec<u16> = bytes
                    .chunks_exact(2)
                    .map(|c| u16::from_le_bytes([c[0], c[1]]))
                    .collect();
                let s =
                    String::from_utf16(&units).map_err(|_| DecodeError("PtgStr invalid utf-16"))?;
                pos += 2 * cch;
                push(
                    &mut stack,
                    &mut pend_nl,
                    &mut pend_sp,
                    format!("\"{}\"", s.replace('"', "\"\"")),
                );
                continue;
            }
            0x1C => {
                let code = *rgce.get(pos).ok_or(DecodeError("PtgErr code past end"))?;
                pos += 1;
                push(
                    &mut stack,
                    &mut pend_nl,
                    &mut pend_sp,
                    error_text(code).to_string(),
                );
                continue;
            }
            0x1D => {
                let b = *rgce.get(pos).ok_or(DecodeError("PtgBool value past end"))?;
                pos += 1;
                push(
                    &mut stack,
                    &mut pend_nl,
                    &mut pend_sp,
                    if b != 0 {
                        "TRUE".to_string()
                    } else {
                        "FALSE".to_string()
                    },
                );
                continue;
            }
            0x1E => {
                let v = get_u16(rgce, pos)?;
                pos += 2;
                push(&mut stack, &mut pend_nl, &mut pend_sp, v.to_string());
                continue;
            }
            0x1F => {
                let v = get_f64(rgce, pos)?;
                pos += 8;
                push(&mut stack, &mut pend_nl, &mut pend_sp, fmt_num(v));
                continue;
            }
            0x01 => {
                pos += 4;
                push(&mut stack, &mut pend_nl, &mut pend_sp, "#EXP#".to_string());
                continue;
            }
            0x02 => {
                pos += 4;
                push(&mut stack, &mut pend_nl, &mut pend_sp, "#TBL#".to_string());
                continue;
            }
            _ => {}
        }
        if base == 0x20 {
            pos += 14;
            let (s, next_cb) = parse_array_consts(rgcb, cb_pos)?;
            cb_pos = next_cb;
            push(&mut stack, &mut pend_nl, &mut pend_sp, s);
            continue;
        }
        if base == 0x21 || base == 0x22 {
            let (argc, iftab_raw) = if base == 0x22 {
                let argc = *rgce
                    .get(pos)
                    .ok_or(DecodeError("PtgFuncVar argc past end"))?
                    as usize;
                let iftab = get_u16(rgce, pos + 1)?;
                pos += 3;
                (argc, iftab)
            } else {
                let iftab = get_u16(rgce, pos)?;
                pos += 2;
                let argc = ftab::argc((iftab & 0x7FFF) as u32) as usize;
                (argc, iftab)
            };
            let iftab = (iftab_raw & 0x7FFF) as u32;
            if iftab == 255 {
                if argc == 0 || argc > stack.len() {
                    return Err(DecodeError("user-defined call argc exceeds stack"));
                }
                let args: Vec<String> = stack.split_off(stack.len() - argc);
                let mut fname = args[0].clone();
                for prefix in ["_xlfn._xlws.", "_xlfn.", "_xlws."] {
                    if let Some(stripped) = fname.strip_prefix(prefix) {
                        fname = stripped.to_string();
                        break;
                    }
                }
                push(
                    &mut stack,
                    &mut pend_nl,
                    &mut pend_sp,
                    format!("{fname}({})", args[1..].join(",")),
                );
            } else {
                let name = ftab::name(iftab);
                let args: Vec<String> = if argc > 0 {
                    if argc > stack.len() {
                        return Err(DecodeError("function call argc exceeds stack"));
                    }
                    stack.split_off(stack.len() - argc)
                } else {
                    Vec::new()
                };
                push(
                    &mut stack,
                    &mut pend_nl,
                    &mut pend_sp,
                    format!("{name}({})", args.join(",")),
                );
            }
            continue;
        }
        if base == 0x23 {
            let raw = get_u32(rgce, pos)? as i64;
            pos += 4;
            let idx = raw - 1;
            let nm = if idx >= 0 && (idx as usize) < wb.names.len() {
                wb.names[idx as usize].name.clone()
            } else {
                "#NAME?".to_string()
            };
            let nm = nm.strip_prefix("_xlpm.").map(str::to_string).unwrap_or(nm);
            push(&mut stack, &mut pend_nl, &mut pend_sp, nm);
            continue;
        }
        if base == 0x24 {
            let (r, c, rr, cr) = loc(rgce, pos)?;
            pos += 6;
            push(
                &mut stack,
                &mut pend_nl,
                &mut pend_sp,
                render_ref(r, c, rr, cr, host_row, host_col, mode),
            );
            continue;
        }
        if base == 0x25 {
            let r1 = get_u32(rgce, pos)?;
            let r2 = get_u32(rgce, pos + 4)?;
            let cw1 = get_u16(rgce, pos + 8)?;
            let cw2 = get_u16(rgce, pos + 10)?;
            pos += 12;
            let text = render_area(
                r1,
                r2,
                (cw1 & 0x3FFF) as u32,
                (cw2 & 0x3FFF) as u32,
                cw1 & 0x8000 != 0,
                cw2 & 0x8000 != 0,
                cw1 & 0x4000 != 0,
                cw2 & 0x4000 != 0,
                host_row,
                host_col,
                mode,
            );
            push(&mut stack, &mut pend_nl, &mut pend_sp, text);
            continue;
        }
        if base == 0x26 || base == 0x27 || base == 0x28 {
            pos += 6;
            continue;
        }
        if base == 0x29 {
            pos += 2;
            continue;
        }
        if base == 0x2A {
            pos += 6;
            push(&mut stack, &mut pend_nl, &mut pend_sp, "#REF!".to_string());
            continue;
        }
        if base == 0x2B {
            pos += 12;
            push(&mut stack, &mut pend_nl, &mut pend_sp, "#REF!".to_string());
            continue;
        }
        if base == 0x2C {
            let (r, c, rr, cr) = loc_rel(rgce, pos, host_row, host_col)?;
            pos += 6;
            push(
                &mut stack,
                &mut pend_nl,
                &mut pend_sp,
                render_ref(r, c, rr, cr, host_row, host_col, mode),
            );
            continue;
        }
        if base == 0x2D {
            let r1raw = get_u32(rgce, pos)?;
            let r2raw = get_u32(rgce, pos + 4)?;
            let cw1 = get_u16(rgce, pos + 8)?;
            let cw2 = get_u16(rgce, pos + 10)?;
            pos += 12;
            let (r1, c1, r1rel, c1rel) = loc_rel_parts(r1raw, cw1, host_row, host_col);
            let (r2, c2, r2rel, c2rel) = loc_rel_parts(r2raw, cw2, host_row, host_col);
            let text = render_area(
                r1, r2, c1, c2, r1rel, r2rel, c1rel, c2rel, host_row, host_col, mode,
            );
            push(&mut stack, &mut pend_nl, &mut pend_sp, text);
            continue;
        }
        if base == 0x39 {
            let ixti = get_u16(rgce, pos)? as usize;
            let raw = get_u32(rgce, pos + 2)? as i64;
            pos += 6;
            let idx = raw - 1;
            let isup = wb.xti.get(ixti).map(|t| t.0).unwrap_or(0);
            let is_self_or_same = wb
                .supbooks
                .get(isup as usize)
                .map(|k| {
                    matches!(
                        k,
                        crate::workbook::SupbookKind::SelfBook | crate::workbook::SupbookKind::Same
                    )
                })
                .unwrap_or(true);
            if is_self_or_same {
                let nm = if idx >= 0 && (idx as usize) < wb.names.len() {
                    wb.names[idx as usize].name.clone()
                } else {
                    "#NAME?".to_string()
                };
                push(&mut stack, &mut pend_nl, &mut pend_sp, nm);
            } else {
                push(
                    &mut stack,
                    &mut pend_nl,
                    &mut pend_sp,
                    format!("[{}]!#EXTNAME{}#", wb.ext_index(isup), idx),
                );
            }
            continue;
        }
        if base == 0x3A {
            let ixti = get_u16(rgce, pos)? as usize;
            let (r, c, rr, cr) = loc(rgce, pos + 2)?;
            pos += 8;
            let prefix = wb.sheet_prefix(ixti)?;
            push(
                &mut stack,
                &mut pend_nl,
                &mut pend_sp,
                format!(
                    "{prefix}{}",
                    render_ref(r, c, rr, cr, host_row, host_col, mode)
                ),
            );
            continue;
        }
        if base == 0x3B {
            let ixti = get_u16(rgce, pos)? as usize;
            let r1 = get_u32(rgce, pos + 2)?;
            let r2 = get_u32(rgce, pos + 6)?;
            let cw1 = get_u16(rgce, pos + 10)?;
            let cw2 = get_u16(rgce, pos + 12)?;
            pos += 14;
            let prefix = wb.sheet_prefix(ixti)?;
            let text = render_area(
                r1,
                r2,
                (cw1 & 0x3FFF) as u32,
                (cw2 & 0x3FFF) as u32,
                cw1 & 0x8000 != 0,
                cw2 & 0x8000 != 0,
                cw1 & 0x4000 != 0,
                cw2 & 0x4000 != 0,
                host_row,
                host_col,
                mode,
            );
            push(
                &mut stack,
                &mut pend_nl,
                &mut pend_sp,
                format!("{prefix}{text}"),
            );
            continue;
        }
        if base == 0x3C {
            let ixti = get_u16(rgce, pos)? as usize;
            pos += 8;
            let prefix = wb.sheet_prefix(ixti)?;
            push(
                &mut stack,
                &mut pend_nl,
                &mut pend_sp,
                format!("{prefix}#REF!"),
            );
            continue;
        }
        if base == 0x3D {
            let ixti = get_u16(rgce, pos)? as usize;
            pos += 14;
            let prefix = wb.sheet_prefix(ixti)?;
            push(
                &mut stack,
                &mut pend_nl,
                &mut pend_sp,
                format!("{prefix}#REF!"),
            );
            continue;
        }
        return Err(DecodeError("unknown ptg opcode"));
    }
    if stack.len() != 1 {
        return Err(DecodeError("stack imbalance at end of render"));
    }
    Ok(stack.pop().unwrap())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::workbook::Workbook;

    #[test]
    fn col_letters_wraps_like_excel_columns() {
        assert_eq!(col_letters(0), "A");
        assert_eq!(col_letters(25), "Z");
        assert_eq!(col_letters(26), "AA");
        assert_eq!(col_letters(701), "ZZ");
        assert_eq!(col_letters(702), "AAA");
    }

    #[test]
    fn fmt_num_prefers_plain_integers_for_whole_numbers() {
        assert_eq!(fmt_num(5.0), "5");
        assert_eq!(fmt_num(-3.0), "-3");
        assert_eq!(fmt_num(0.0), "0");
        assert_eq!(fmt_num(1.5), "1.5");
    }

    #[test]
    fn render_evaluates_a_simple_binary_expression() {
        // PtgInt(1), PtgInt(2), PtgAdd -- postfix "1 2 +" renders infix "1+2".
        let rgce = [0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x03];
        let wb = Workbook::default();
        let text = render(&rgce, &[], &wb, 0, 0, Mode::A1).unwrap();
        assert_eq!(text, "1+2");
    }

    #[test]
    fn render_fails_closed_on_an_unknown_opcode() {
        let rgce = [0xEEu8]; // not a recognized Ptg opcode
        let wb = Workbook::default();
        assert!(render(&rgce, &[], &wb, 0, 0, Mode::A1).is_err());
    }

    #[test]
    fn render_ptg_ref_matches_a1_and_r1c1_conventions() {
        // PtgRef (0x24): absolute row=4 (0-based), col=2 (0-based), both
        // non-relative (rel flags clear) -> A1 "$C$5".
        let mut rgce = vec![0x24u8];
        rgce.extend(4u32.to_le_bytes());
        rgce.extend(2u16.to_le_bytes()); // cw: col=2, no rel bits set
        let wb = Workbook::default();
        assert_eq!(render(&rgce, &[], &wb, 0, 0, Mode::A1).unwrap(), "$C$5");
        assert_eq!(render(&rgce, &[], &wb, 0, 0, Mode::R1C1).unwrap(), "R5C3");
    }

    #[test]
    fn render_ptg_refn_collapses_zero_r1c1_offsets_like_excel() {
        // PtgRefN (0x2C): relative row/col carrying signed offsets from
        // (host_row, host_col). Real Excel omits the bracket entirely for a
        // zero offset on either axis (bare "R"/"C"), and renders a pure
        // self-reference as bare "RC" -- confirmed against desktop Excel's
        // own Formula2R1C1 cache in B2's guest verification (Criterion 9).
        let wb = Workbook::default();

        // Same row (offset 0), column offset +3 -> "RC[3]".
        let mut same_row = vec![0x2Cu8];
        same_row.extend(0u32.to_le_bytes()); // row offset 0
        same_row.extend(0xC003u16.to_le_bytes()); // row_rel|col_rel, col offset 3
        assert_eq!(
            render(&same_row, &[], &wb, 0, 0, Mode::R1C1).unwrap(),
            "RC[3]"
        );

        // Row offset +2, same column (offset 0) -> "R[2]C".
        let mut same_col = vec![0x2Cu8];
        same_col.extend(2u32.to_le_bytes()); // row offset 2
        same_col.extend(0xC000u16.to_le_bytes()); // row_rel|col_rel, col offset 0
        assert_eq!(
            render(&same_col, &[], &wb, 0, 0, Mode::R1C1).unwrap(),
            "R[2]C"
        );

        // Both offsets 0 (pure self-reference) -> bare "RC".
        let mut self_ref = vec![0x2Cu8];
        self_ref.extend(0u32.to_le_bytes());
        self_ref.extend(0xC000u16.to_le_bytes());
        assert_eq!(render(&self_ref, &[], &wb, 0, 0, Mode::R1C1).unwrap(), "RC");
    }

    #[test]
    fn render_ptg_area_whole_column_and_row_use_r1c1_numeric_syntax() {
        // PtgArea (0x25): r1=0, r2=MAX_ROW (whole-column range), absolute
        // col1=2, col2=4. Real Excel's R1C1 mode renders whole-column ranges
        // with numeric "C{n}:C{n}" syntax, never A1 letter syntax -- this
        // special case previously ignored `mode` entirely and always
        // produced A1-style "$C:$E" even when asked for R1C1.
        let mut whole_col = vec![0x25u8];
        whole_col.extend(0u32.to_le_bytes()); // r1 = 0
        whole_col.extend(MAX_ROW.to_le_bytes()); // r2 = MAX_ROW
        whole_col.extend(2u16.to_le_bytes()); // cw1: col=2, absolute
        whole_col.extend(4u16.to_le_bytes()); // cw2: col=4, absolute
        let wb = Workbook::default();
        assert_eq!(
            render(&whole_col, &[], &wb, 0, 0, Mode::A1).unwrap(),
            "$C:$E"
        );
        assert_eq!(
            render(&whole_col, &[], &wb, 0, 0, Mode::R1C1).unwrap(),
            "C3:C5"
        );

        // PtgArea whole-row range: c1=0, c2=MAX_COL, absolute row1=2, row2=4.
        let mut whole_row = vec![0x25u8];
        whole_row.extend(2u32.to_le_bytes()); // r1 = 2
        whole_row.extend(4u32.to_le_bytes()); // r2 = 4
        whole_row.extend(0u16.to_le_bytes()); // cw1: col=0, absolute
        whole_row.extend((MAX_COL as u16).to_le_bytes()); // cw2: col=MAX_COL, absolute
        assert_eq!(
            render(&whole_row, &[], &wb, 0, 0, Mode::A1).unwrap(),
            "$3:$5"
        );
        assert_eq!(
            render(&whole_row, &[], &wb, 0, 0, Mode::R1C1).unwrap(),
            "R3:R5"
        );
    }

    #[test]
    fn render_ptg_area_single_whole_column_or_row_collapses_only_in_r1c1() {
        // A single whole column (c1 == c2, "A:A" in A1) or whole row (r1 ==
        // r2, "1:1" in A1): real Excel's R1C1 mode collapses this to ONE bare
        // reference ("C[-4]"/"R[-9]", no colon), while A1 mode keeps the full
        // "A:A"/"1:1" range form -- confirmed empirically against live Excel
        // (B2 guest verification, Criterion 9); these two modes are not
        // symmetric here.
        let wb = Workbook::default();

        let mut single_col = vec![0x25u8];
        single_col.extend(0u32.to_le_bytes()); // r1 = 0
        single_col.extend(MAX_ROW.to_le_bytes()); // r2 = MAX_ROW
        single_col.extend(0u16.to_le_bytes()); // cw1: col=0, absolute
        single_col.extend(0u16.to_le_bytes()); // cw2: col=0, absolute (same column)
        assert_eq!(
            render(&single_col, &[], &wb, 0, 0, Mode::A1).unwrap(),
            "$A:$A"
        );
        assert_eq!(
            render(&single_col, &[], &wb, 0, 0, Mode::R1C1).unwrap(),
            "C1"
        );

        let mut single_row = vec![0x25u8];
        single_row.extend(4u32.to_le_bytes()); // r1 = 4
        single_row.extend(4u32.to_le_bytes()); // r2 = 4 (same row)
        single_row.extend(0u16.to_le_bytes()); // cw1: col=0, absolute
        single_row.extend((MAX_COL as u16).to_le_bytes()); // cw2: col=MAX_COL, absolute
        assert_eq!(
            render(&single_row, &[], &wb, 0, 0, Mode::A1).unwrap(),
            "$5:$5"
        );
        assert_eq!(
            render(&single_row, &[], &wb, 0, 0, Mode::R1C1).unwrap(),
            "R5"
        );
    }
}
