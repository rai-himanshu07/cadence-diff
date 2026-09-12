//! Production-shaped formula surface: per-definition R1C1 text (deduplicated,
//! rendered once per shared/array group -- a definition_id links each cell to
//! its group) plus eager per-cell A1 text (needed since every existing
//! `CellRecord.formula` consumer expects per-cell A1, and A1 genuinely varies
//! per follower position, unlike R1C1) plus defined names with rendered A1
//! targets. This is B2's "definitions + per-cell ids + names surface"
//! deliverable, distinct from B0/B1's flat per-cell probe-oriented output
//! (`sheet::decode_workbook`/`values::decode_workbook_values`), which stay
//! unchanged since their exact-parity proof already stands and nothing here
//! should risk regressing it.

use crate::ptg::{render, Mode};
use crate::records::{get_u32, DecodeError, Records, Result};
use crate::sheet::{
    cell_formula, fmla_kind, part_path, read_zip_entry, R_ARRFMLA, R_ROWHDR, R_SHRFMLA,
};
use crate::workbook::Workbook;
use std::collections::HashMap;

pub struct DefinedNameSurface {
    pub name: String,
    /// Rendered A1 target text, or `None` if it failed to decode (fail
    /// closed, never guessed).
    pub target_a1: Option<String>,
}

#[derive(Default)]
pub struct SheetSurface {
    pub name: String,
    pub cell_rows: Vec<u32>,
    pub cell_cols: Vec<u32>,
    pub cell_a1: Vec<String>,
    pub cell_definition_id: Vec<u32>,
    /// Indexed by definition id (0-based, contiguous within this sheet).
    pub definition_r1c1: Vec<String>,
}

/// Registers (or looks up) a definition's R1C1 text by the identity of its
/// `rgce` slice (pointer + length -- stable for the lifetime of the one
/// `part` buffer this function processes, mirroring `ptg_decoder.py`'s own
/// `id(hit[0])` memoization trick).
fn definition_id_for(
    rgce: &[u8],
    rgcb: &[u8],
    host_row: u32,
    host_col: u32,
    wb: &Workbook,
    ids: &mut HashMap<(usize, usize), u32>,
    r1c1_table: &mut Vec<String>,
) -> u32 {
    let key = (rgce.as_ptr() as usize, rgce.len());
    if let Some(&id) = ids.get(&key) {
        return id;
    }
    let text =
        render(rgce, rgcb, wb, host_row, host_col, Mode::R1C1).unwrap_or_else(|_| String::new());
    // An empty string here is never trusted blindly downstream: the Python
    // boundary (`_validated_r1c1_cells`) cross-checks every definition's
    // R1C1 text against an independent, already-Excel-proven `to_r1c1()`
    // computation and replaces any mismatch (including this empty-string
    // case) before it reaches a consumer -- this is the established B5
    // validate-and-fallback design, not an unguarded fail-open path. A1
    // above has no such downstream check, which is why it must fail closed
    // by omission instead.
    let id = r1c1_table.len() as u32;
    r1c1_table.push(text);
    ids.insert(key, id);
    id
}

pub fn decode_sheet_surface(part: &[u8], wb: &Workbook, sheet_name: String) -> SheetSurface {
    let mut surface = SheetSurface {
        name: sheet_name,
        ..Default::default()
    };
    let mut row: u32 = 0;
    let mut pending: Vec<(u32, u32, &[u8])> = Vec::new();
    let mut shared: HashMap<(u32, u32), Vec<(u32, u32, &[u8], &[u8])>> = HashMap::new();
    let mut wide: Vec<(u32, u32, u32, u32, &[u8], &[u8])> = Vec::new();
    let mut definition_ids: HashMap<(usize, usize), u32> = HashMap::new();

    for rec in Records::new(part) {
        let Ok(rec) = rec else { break };
        let pl = rec.payload;
        if rec.id == R_ROWHDR {
            if let Ok(v) = get_u32(pl, 0) {
                row = v;
            }
        } else if let Some(kind) = fmla_kind(rec.id) {
            let Ok((col, rgce, rgcb)) = cell_formula(kind, pl) else {
                continue;
            };
            if !rgce.is_empty() && rgce[0] == 0x01 {
                pending.push((row, col, rgce));
                continue;
            }
            // Direct (non-shared) formula: its own singleton definition.
            // Fail closed: a render failure (unknown/malformed token) leaves
            // this cell entirely absent from the surface -- never a blank or
            // synthesized "=" formula (Criterion 1).
            let Ok(a1) = render(rgce, rgcb, wb, row, col, Mode::A1) else {
                continue;
            };
            let def_id = definition_id_for(
                rgce,
                rgcb,
                row,
                col,
                wb,
                &mut definition_ids,
                &mut surface.definition_r1c1,
            );
            surface.cell_rows.push(row);
            surface.cell_cols.push(col);
            surface.cell_a1.push(a1);
            surface.cell_definition_id.push(def_id);
        } else if rec.id == R_SHRFMLA || rec.id == R_ARRFMLA {
            let parsed: Result<()> = (|| {
                let r1 = get_u32(pl, 0)?;
                let r2 = get_u32(pl, 4)?;
                let c1 = get_u32(pl, 8)?;
                let c2 = get_u32(pl, 12)?;
                let mut pos = 16 + usize::from(rec.id == R_ARRFMLA);
                let cce = get_u32(pl, pos)? as usize;
                pos += 4;
                let rgce = crate::records::get_slice(pl, pos, cce)?;
                pos += cce;
                let cb = get_u32(pl, pos)? as usize;
                pos += 4;
                let rgcb = crate::records::get_slice(pl, pos, cb)?;
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
            let _ = parsed; // malformed shared/array headers simply contribute no definitions
        }
    }

    for (row, col, rgce) in pending {
        let Ok(exp_row) = get_u32(rgce, 1) else {
            continue;
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
        let Some((srgce, srgcb)) = hit else { continue }; // unresolved: no text, no definition (disclosed by absence)
                                                          // Fail closed: a render failure here is likewise absent, never a
                                                          // blank or synthesized "=" formula (Criterion 1) -- other followers
                                                          // of the same shared/array definition still register it normally.
        let Ok(a1) = render(srgce, srgcb, wb, row, col, Mode::A1) else {
            continue;
        };
        let def_id = definition_id_for(
            srgce,
            srgcb,
            row,
            col,
            wb,
            &mut definition_ids,
            &mut surface.definition_r1c1,
        );
        surface.cell_rows.push(row);
        surface.cell_cols.push(col);
        surface.cell_a1.push(a1);
        surface.cell_definition_id.push(def_id);
    }

    surface
}

fn decode_defined_names(wb: &Workbook) -> Vec<DefinedNameSurface> {
    wb.names
        .iter()
        .map(|n| DefinedNameSurface {
            name: n.name.clone(),
            target_a1: render(&n.rgce, &n.rgcb, wb, 0, 0, Mode::A1).ok(),
        })
        .collect()
}

pub struct WorkbookSurface {
    pub sheets: Vec<SheetSurface>,
    pub defined_names: Vec<DefinedNameSurface>,
}

/// Opens an XLSB from in-memory bytes and decodes the full production
/// formula surface: per-sheet cells (A1 text + definition id), per-sheet
/// deduplicated definition R1C1 table, and defined names with A1 targets.
pub fn decode_workbook_surface(data: &[u8]) -> Result<WorkbookSurface> {
    let cursor = std::io::Cursor::new(data);
    let mut archive =
        zip::ZipArchive::new(cursor).map_err(|_| DecodeError("not a valid zip/xlsb container"))?;
    let workbook_bin = read_zip_entry(&mut archive, "xl/workbook.bin")?;
    let wb = crate::workbook::parse_workbook(&workbook_bin)?;
    let rels_xml = read_zip_entry(&mut archive, "xl/_rels/workbook.bin.rels")?;
    let rel_targets = crate::workbook::parse_workbook_rels(&rels_xml)?;

    let mut sheets = Vec::new();
    for (name, relid) in wb.sheets.iter().zip(wb.relids.iter()) {
        let Some(relid) = relid else { continue };
        let Some(target) = rel_targets.get(relid) else {
            continue;
        };
        if !target.contains("worksheets") {
            continue;
        }
        let part = part_path(target);
        let Ok(data) = read_zip_entry(&mut archive, &part) else {
            continue;
        };
        sheets.push(decode_sheet_surface(&data, &wb, name.clone()));
    }
    let defined_names = decode_defined_names(&wb);
    Ok(WorkbookSurface {
        sheets,
        defined_names,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::workbook::Workbook;

    fn record(id: u8, payload: &[u8]) -> Vec<u8> {
        let mut out = vec![id, payload.len() as u8];
        out.extend_from_slice(payload);
        out
    }

    /// One minimal BIFF12 worksheet part: a `BrtRowHdr` setting `row`,
    /// followed by one direct (non-shared) `BrtCellReal`-with-formula
    /// ("num" kind, record id 0x09) cell at `col` whose formula token
    /// stream is exactly `rgce` -- with no cached double, no rgcb, and
    /// zeroed filler bytes the formula walk never reads.
    fn direct_formula_cell_part(row: u32, col: u32, rgce: &[u8]) -> Vec<u8> {
        let mut out = record(0x00, &row.to_le_bytes()); // BrtRowHdr
        let mut payload = Vec::new();
        payload.extend(col.to_le_bytes()); // col (masked with 0x3FFF on read)
        payload.extend([0u8; 4]); // ixfe/reserved filler, never read
        payload.extend([0u8; 8]); // cached double result, never read
        payload.extend([0u8; 2]); // grbitFlags, never read
        payload.extend((rgce.len() as u32).to_le_bytes()); // cce
        payload.extend_from_slice(rgce);
        payload.extend(0u32.to_le_bytes()); // cb = 0 (no rgcb)
        out.extend(record(0x09, &payload));
        out
    }

    #[test]
    fn decode_sheet_surface_omits_a_cell_whose_a1_render_fails() {
        // Criterion 1: a render failure (here, an unrecognized Ptg opcode)
        // must leave the cell entirely absent -- never a blank/"=" entry.
        let wb = Workbook::default();
        let part = direct_formula_cell_part(0, 0, &[0xEE]);

        let surface = decode_sheet_surface(&part, &wb, "Data".to_string());

        assert!(surface.cell_rows.is_empty());
        assert!(surface.cell_cols.is_empty());
        assert!(surface.cell_a1.is_empty());
        assert!(surface.cell_definition_id.is_empty());
        assert!(surface.definition_r1c1.is_empty());
    }

    #[test]
    fn decode_sheet_surface_keeps_a_cell_whose_a1_render_succeeds() {
        // PtgInt(1), PtgInt(2), PtgAdd -- postfix "1 2 +" renders "1+2".
        let rgce = [0x1E, 0x01, 0x00, 0x1E, 0x02, 0x00, 0x03];
        let wb = Workbook::default();
        let part = direct_formula_cell_part(0, 0, &rgce);

        let surface = decode_sheet_surface(&part, &wb, "Data".to_string());

        assert_eq!(surface.cell_rows, vec![0]);
        assert_eq!(surface.cell_cols, vec![0]);
        assert_eq!(surface.cell_a1, vec!["1+2".to_string()]);
        assert_eq!(surface.cell_definition_id, vec![0]);
        assert_eq!(surface.definition_r1c1, vec!["1+2".to_string()]);
    }
}
