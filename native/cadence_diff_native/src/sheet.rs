//! Worksheet formula walk: tracks the current row via `BrtRowHdr`, collects
//! `BrtShrFmla`/`BrtArrFmla` shared/array definitions, defers follower cells
//! (whose `rgce` is a single `PtgExp` pointing at a definition) until the
//! whole sheet has been scanned, then resolves and renders every cell.
//! Ports `decode_sheet`/`_cell_formula`/`decode_workbook` in `ptg_decoder.py`.

use crate::ptg::render;
pub use crate::ptg::Mode;
use crate::records::{get_slice, get_u32, naive, DecodeError, Records, Result};
use crate::workbook::{parse_workbook, parse_workbook_rels, Workbook};
use std::collections::HashMap;
use std::io::Read;

pub(crate) const R_ROWHDR: u32 = 0x00;
pub(crate) const R_SHRFMLA: u32 = naive(0x01AB);
pub(crate) const R_ARRFMLA: u32 = naive(0x01AA);

pub(crate) fn fmla_kind(rid: u32) -> Option<&'static str> {
    match rid {
        0x08 => Some("str"),
        0x09 => Some("num"),
        0x0A => Some("bool"),
        0x0B => Some("err"),
        _ => None,
    }
}

pub(crate) fn cell_formula<'a>(kind: &str, pl: &'a [u8]) -> Result<(u32, &'a [u8], &'a [u8])> {
    let col = get_u32(pl, 0)? & 0x3FFF;
    let mut pos = 8usize;
    match kind {
        "str" => {
            let cch = get_u32(pl, pos)? as usize;
            pos += 4 + 2 * cch;
        }
        "num" => pos += 8,
        _ => pos += 1,
    }
    pos += 2; // grbitFlags
    let cce = get_u32(pl, pos)? as usize;
    pos += 4;
    let rgce = get_slice(pl, pos, cce)?;
    pos += cce;
    let cb = get_u32(pl, pos)? as usize;
    pos += 4;
    let rgcb = get_slice(pl, pos, cb)?;
    Ok((col, rgce, rgcb))
}

#[derive(Default)]
pub struct SheetReport {
    pub name: String,
    pub cells: HashMap<(u32, u32), String>,
    /// Fixed diagnostic message -> occurrence count. Messages are always
    /// static strings (see `DecodeError`), never formula text or coordinates.
    pub errors: HashMap<String, u32>,
    pub unresolved_followers: u32,
}

fn record_error(errors: &mut HashMap<String, u32>, key: &str) {
    *errors.entry(key.to_string()).or_insert(0) += 1;
}

/// Decodes every formula cell on one worksheet part to A1 or R1C1 text
/// (fail-closed per cell: a cell that cannot be decoded is absent from
/// `cells` and counted in `errors`/`unresolved_followers` instead of guessed).
pub fn decode_sheet(part: &[u8], wb: &Workbook, sheet_name: String, mode: Mode) -> SheetReport {
    let mut report = SheetReport {
        name: sheet_name,
        ..Default::default()
    };
    let mut row: u32 = 0;
    let mut pending: Vec<(u32, u32, &[u8])> = Vec::new();
    let mut shared: HashMap<(u32, u32), Vec<(u32, u32, &[u8], &[u8])>> = HashMap::new();
    let mut wide: Vec<(u32, u32, u32, u32, &[u8], &[u8])> = Vec::new();

    for rec in Records::new(part) {
        let rec = match rec {
            Ok(r) => r,
            Err(e) => {
                // A corrupt record stream cannot be safely resynchronized;
                // stop with whatever was already decoded, disclosed via the
                // error count rather than guessing at a resync point.
                record_error(&mut report.errors, e.0);
                break;
            }
        };
        let pl = rec.payload;
        if rec.id == R_ROWHDR {
            match get_u32(pl, 0) {
                Ok(v) => row = v,
                Err(_) => record_error(&mut report.errors, "row header malformed"),
            }
        } else if let Some(kind) = fmla_kind(rec.id) {
            match cell_formula(kind, pl) {
                Ok((col, rgce, rgcb)) => {
                    if !rgce.is_empty() && rgce[0] == 0x01 {
                        pending.push((row, col, rgce));
                    } else {
                        match render(rgce, rgcb, wb, row, col, mode) {
                            Ok(text) => {
                                report.cells.insert((row, col), text);
                            }
                            Err(e) => record_error(&mut report.errors, e.0),
                        }
                    }
                }
                Err(e) => record_error(&mut report.errors, e.0),
            }
        } else if rec.id == R_SHRFMLA || rec.id == R_ARRFMLA {
            let parsed: Result<()> = (|| {
                let r1 = get_u32(pl, 0)?;
                let r2 = get_u32(pl, 4)?;
                let c1 = get_u32(pl, 8)?;
                let c2 = get_u32(pl, 12)?;
                let mut pos = 16 + usize::from(rec.id == R_ARRFMLA);
                let cce = get_u32(pl, pos)? as usize;
                pos += 4;
                let rgce = get_slice(pl, pos, cce)?;
                pos += cce;
                let cb = get_u32(pl, pos)? as usize;
                pos += 4;
                let rgcb = get_slice(pl, pos, cb)?;
                if c2.saturating_sub(c1) > 64 {
                    wide.push((r1, r2, c1, c2, rgce, rgcb));
                } else {
                    for c in c1..=c2 {
                        shared
                            .entry((r1, c))
                            .or_default()
                            .push((r1, r2, rgce, rgcb));
                    }
                }
                Ok(())
            })();
            if let Err(e) = parsed {
                record_error(&mut report.errors, e.0);
            }
        }
    }

    for (row, col, rgce) in pending {
        let exp_row = match get_u32(rgce, 1) {
            Ok(v) => v,
            Err(_) => {
                record_error(&mut report.errors, "PtgExp anchor malformed");
                continue;
            }
        };
        let mut hit: Option<(&[u8], &[u8])> = None;
        if let Some(candidates) = shared.get(&(exp_row, col)) {
            for &(r1, r2, srgce, srgcb) in candidates {
                if r1 <= row && row <= r2 {
                    hit = Some((srgce, srgcb));
                    break;
                }
            }
        }
        if hit.is_none() {
            for &(r1, r2, c1, c2, srgce, srgcb) in &wide {
                if r1 == exp_row && r1 <= row && row <= r2 && c1 <= col && col <= c2 {
                    hit = Some((srgce, srgcb));
                    break;
                }
            }
        }
        let Some((srgce, srgcb)) = hit else {
            report.unresolved_followers += 1;
            continue;
        };
        match render(srgce, srgcb, wb, row, col, mode) {
            Ok(text) => {
                report.cells.insert((row, col), text);
            }
            Err(e) => record_error(&mut report.errors, &format!("follower:{}", e.0)),
        }
    }

    report
}

pub(crate) fn read_zip_entry<R: Read + std::io::Seek>(
    archive: &mut zip::ZipArchive<R>,
    name: &str,
) -> Result<Vec<u8>> {
    let mut entry = archive
        .by_name(name)
        .map_err(|_| DecodeError("zip entry not found"))?;
    let mut buf = Vec::with_capacity(entry.size() as usize);
    entry
        .read_to_end(&mut buf)
        .map_err(|_| DecodeError("zip entry read failed"))?;
    Ok(buf)
}

pub(crate) fn part_path(target: &str) -> String {
    let stripped = target.trim_start_matches('/');
    let stripped = stripped.strip_prefix("xl/").unwrap_or(stripped);
    format!("xl/{stripped}")
}

/// Opens an XLSB (a ZIP/OPC package) from in-memory bytes -- never a file
/// path, so this matches whatever bytes the caller actually has in hand
/// (e.g. already password-decrypted, which may differ from the on-disk
/// file) -- parses the workbook context, and decodes every worksheet's
/// formula cells to A1 or R1C1 text.
pub fn decode_workbook(data: &[u8], mode: Mode) -> Result<(Workbook, Vec<SheetReport>)> {
    let cursor = std::io::Cursor::new(data);
    let mut archive =
        zip::ZipArchive::new(cursor).map_err(|_| DecodeError("not a valid zip/xlsb container"))?;
    let workbook_bin = read_zip_entry(&mut archive, "xl/workbook.bin")?;
    let wb = parse_workbook(&workbook_bin)?;
    let rels_xml = read_zip_entry(&mut archive, "xl/_rels/workbook.bin.rels")?;
    let rel_targets = parse_workbook_rels(&rels_xml)?;

    let mut reports = Vec::new();
    for (name, relid) in wb.sheets.iter().zip(wb.relids.iter()) {
        let Some(relid) = relid else { continue };
        let Some(target) = rel_targets.get(relid) else {
            continue;
        };
        if !target.contains("worksheets") {
            continue;
        }
        let part = part_path(target);
        let data = match read_zip_entry(&mut archive, &part) {
            Ok(d) => d,
            Err(_) => continue,
        };
        reports.push(decode_sheet(&data, &wb, name.clone(), mode));
    }
    Ok((wb, reports))
}
