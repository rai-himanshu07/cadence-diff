//! Raw BIFF12 cell-value reader, ported byte-for-byte from pyxlsb's own
//! `CellHandler`/`StringTable`/`reader.py` behavior (verified by direct
//! inspection of the installed pyxlsb 1.x source -- see the B1 decision
//! record), NOT calamine. calamine's `worksheet_range()` auto-classifies
//! date-formatted numeric cells into a separate `Data::DateTime` variant
//! that pyxlsb never produces (confirmed empirically: 2,014/1,943,753 real
//! LARGE_WORKBOOK cells reclassified on one file alone), which would break exact
//! parity with the pyxlsb path this project's number-format/date
//! conversion already depends on running AFTER, not before, this layer.
//! This revises B0's tentative "keep calamine for values" note with direct
//! evidence -- see the B1 Execution Log entry.
//!
//! Deliberately mirrors pyxlsb's own "always float, even for whole numbers"
//! RK decode and its `errors='replace'` lenient UTF-16 string decode -- both
//! are correctness requirements for parity, not stylistic choices.

use crate::records::{get_f64, get_u32, naive, DecodeError, Records, Result, MAX_WIDE_CHARS};

const R_ROWHDR: u32 = 0x00;
const R_BLANK: u32 = 0x01;
const R_NUM: u32 = 0x02; // BrtCellRk: compact 4-byte number
const R_BOOLERR: u32 = 0x03; // BrtCellError: 1-byte error code
const R_BOOL: u32 = 0x04; // BrtCellBool: 1-byte 0/1
const R_FLOAT: u32 = 0x05; // BrtCellReal: full 8-byte double
const R_STRING: u32 = 0x07; // BrtCellIsst: 4-byte SST index
const R_FMLA_STRING: u32 = 0x08; // cached formula result: inline string
const R_FMLA_FLOAT: u32 = 0x09; // cached formula result: full 8-byte double
const R_FMLA_BOOL: u32 = 0x0A; // cached formula result: 1-byte 0/1
const R_FMLA_BOOLERR: u32 = 0x0B; // cached formula result: 1-byte error code
const R_SI: u32 = 0x13; // BrtSSTItem
const R_SST_END: u32 = naive(0x01A0);

/// Reads pyxlsb's RK-style compact 4-byte number. Bit 0x02 = integer-encoded
/// (still returned as `f64`, matching pyxlsb's `float(intval >> 2)` exactly
/// -- pyxlsb never distinguishes int- vs float-valued numeric cells, so
/// neither does this reader). Bit 0x01 = divide by 100 (2-decimal scaling).
fn read_rk(buf: &[u8], pos: usize) -> Result<f64> {
    let raw = get_u32(buf, pos)? as i32;
    let v = if raw & 0x02 != 0 {
        (raw >> 2) as f64
    } else {
        // Clear the low 2 flag bits, then reinterpret as the high 32 bits of
        // an 8-byte double (matching pyxlsb's `\x00\x00\x00\x00 + pack(intval & 0xFFFFFFFC)`).
        let high = (raw as u32) & 0xFFFF_FFFC;
        f64::from_bits((high as u64) << 32)
    };
    Ok(if raw & 0x01 != 0 { v / 100.0 } else { v })
}

/// Lenient UTF-16LE decode matching Python's `bytes.decode('utf-16',
/// errors='replace')`: an unpaired/invalid code unit becomes U+FFFD instead
/// of failing the whole string. `pos` is the byte offset of a u32
/// code-unit count followed by that many UTF-16LE units (pyxlsb's
/// `read_string`/`read_int` pair, distinct from this crate's stricter
/// `records::read_wide` used for the already-parity-proven formula path).
fn read_string_lenient(buf: &[u8], pos: usize) -> Result<(String, usize)> {
    let cch = get_u32(buf, pos)? as usize;
    let pos = pos + 4;
    if cch > MAX_WIDE_CHARS {
        return Err(DecodeError("string length exceeds cap"));
    }
    let byte_len = cch * 2;
    let bytes = buf
        .get(pos..pos + byte_len)
        .ok_or(DecodeError("string bytes past end"))?;
    let units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    let s: String = char::decode_utf16(units)
        .map(|r| r.unwrap_or(char::REPLACEMENT_CHARACTER))
        .collect();
    Ok((s, pos + byte_len))
}

/// Parses `xl/sharedStrings.bin` into an ordinal-indexed string table,
/// mirroring pyxlsb's `StringTable`. Absent part -> empty table (a workbook
/// with zero string cells may omit it entirely).
pub fn parse_shared_strings(data: &[u8]) -> Result<Vec<String>> {
    let mut strings = Vec::new();
    for rec in Records::new(data) {
        let rec = rec?;
        if rec.id == R_SI {
            // 1-byte grbit flag precedes the string, mirroring
            // `StringInstanceHandler` (`reader.skip(1)` then `read_string()`).
            let (s, _) = read_string_lenient(rec.payload, 1)?;
            strings.push(s);
        } else if rec.id == R_SST_END {
            break;
        }
    }
    Ok(strings)
}

/// One decoded, non-blank cell in pyxlsb's own raw shape: exactly one of
/// `num`/`boolean`/`text` is populated (mirrors `.v` being a plain float,
/// bool, str [including the `"0xNN"` error-hex convention `_XLSB_ERRORS`
/// keys on], or absent for blanks -- blanks are never emitted as rows here,
/// matching pyxlsb's own sparse iteration).
pub struct RawCell {
    pub row: u32,
    pub col: u32,
    pub num: Option<f64>,
    pub boolean: Option<bool>,
    pub text: Option<String>,
}

/// Decodes every non-blank cell value on one worksheet part, resolving
/// shared-string indices via `sst`. Fails closed per cell (a cell that
/// cannot be decoded is silently absent, matching this crate's formula
/// reader's disclosed-gap convention) rather than guessed; a genuinely
/// corrupt record stream stops the walk at the point of corruption.
pub fn decode_sheet_values(part: &[u8], sst: &[String]) -> Vec<RawCell> {
    let mut out = Vec::new();
    let mut row: u32 = 0;
    for rec in Records::new(part) {
        let Ok(rec) = rec else { break };
        let pl = rec.payload;
        if rec.id == R_ROWHDR {
            if let Ok(v) = get_u32(pl, 0) {
                row = v;
            }
            continue;
        }
        let Ok(col) = get_u32(pl, 0) else { continue };
        let cell = match rec.id {
            R_BLANK => None,
            R_NUM => read_rk(pl, 8).ok().map(|v| RawCell {
                row,
                col,
                num: Some(v),
                boolean: None,
                text: None,
            }),
            R_FLOAT | R_FMLA_FLOAT => get_f64(pl, 8).ok().map(|v| RawCell {
                row,
                col,
                num: Some(v),
                boolean: None,
                text: None,
            }),
            R_BOOL | R_FMLA_BOOL => pl.get(8).map(|&b| RawCell {
                row,
                col,
                num: None,
                boolean: Some(b != 0),
                text: None,
            }),
            R_BOOLERR | R_FMLA_BOOLERR => pl.get(8).map(|&b| RawCell {
                row,
                col,
                num: None,
                boolean: None,
                text: Some(format!("0x{b:x}")),
            }),
            R_STRING => get_u32(pl, 8)
                .ok()
                .and_then(|idx| sst.get(idx as usize))
                .map(|s| RawCell {
                    row,
                    col,
                    num: None,
                    boolean: None,
                    text: Some(s.clone()),
                }),
            R_FMLA_STRING => read_string_lenient(pl, 8).ok().map(|(s, _)| RawCell {
                row,
                col,
                num: None,
                boolean: None,
                text: Some(s),
            }),
            _ => None,
        };
        if let Some(cell) = cell {
            out.push(cell);
        }
    }
    out
}

/// Opens an XLSB from in-memory bytes (never a file path -- see
/// `sheet::decode_workbook`'s doc comment for why), parses the workbook
/// context (sheets/rels) and the shared string table, then decodes every
/// worksheet's cell values. Mirrors `sheet::decode_workbook`'s zip/rels
/// plumbing exactly (reused, not duplicated) but reads value records
/// instead of formula records.
pub fn decode_workbook_values(data: &[u8]) -> Result<Vec<(String, Vec<RawCell>)>> {
    let cursor = std::io::Cursor::new(data);
    let mut archive =
        zip::ZipArchive::new(cursor).map_err(|_| DecodeError("not a valid zip/xlsb container"))?;
    let workbook_bin = crate::sheet::read_zip_entry(&mut archive, "xl/workbook.bin")?;
    let wb = crate::workbook::parse_workbook(&workbook_bin)?;
    let rels_xml = crate::sheet::read_zip_entry(&mut archive, "xl/_rels/workbook.bin.rels")?;
    let rel_targets = crate::workbook::parse_workbook_rels(&rels_xml)?;

    let sst = match crate::sheet::read_zip_entry(&mut archive, "xl/sharedStrings.bin") {
        Ok(data) => parse_shared_strings(&data)?,
        Err(_) => Vec::new(),
    };

    let mut reports = Vec::new();
    for (name, relid) in wb.sheets.iter().zip(wb.relids.iter()) {
        let Some(relid) = relid else { continue };
        let Some(target) = rel_targets.get(relid) else {
            continue;
        };
        if !target.contains("worksheets") {
            continue;
        }
        let part = crate::sheet::part_path(target);
        let data = match crate::sheet::read_zip_entry(&mut archive, &part) {
            Ok(d) => d,
            Err(_) => continue,
        };
        reports.push((name.clone(), decode_sheet_values(&data, &sst)));
    }
    Ok(reports)
}
