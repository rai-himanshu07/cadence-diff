//! Workbook context: sheet list, external-reference (XTI/supbook) table and
//! defined names, parsed from `xl/workbook.bin`. Needed to render 3-D
//! references (`PtgRef3d`/`PtgArea3d`), name references (`PtgName`/
//! `PtgNameX`) and to map `BrtBundleSh` sheets to their worksheet parts via
//! `xl/_rels/workbook.bin.rels`. Ports `Workbook`/`parse_workbook` in
//! `ptg_decoder.py`.

use crate::records::{get_i32, get_u32, naive, read_wide, DecodeError, Records, Result};
use once_cell::sync::Lazy;
use regex::Regex;

const R_BUNDLESH: u32 = naive(0x009C);
const R_EXTERNSHEET: u32 = naive(0x016A);
const R_NAME: u32 = 0x27;
const R_SUPSELF: u32 = naive(0x0165);
const R_SUPSAME: u32 = naive(0x0166);
const R_SUPBOOKSRC: u32 = naive(0x0163);
const R_SUPTABS: u32 = naive(0x0167);
const R_SUPADDIN: u32 = naive(0x0169);

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum SupbookKind {
    SelfBook,
    Same,
    Ext,
    Addin,
}

/// One defined name: its own name, home sheet (`-1` workbook-level), and the
/// raw formula token bytes (rendered lazily, only if a caller needs the
/// target text).
pub struct DefinedName {
    pub name: String,
    #[allow(dead_code)] // parsed for completeness; not yet consumed by any caller
    pub itab: i32,
    pub rgce: Vec<u8>,
    pub rgcb: Vec<u8>,
}

#[derive(Default)]
pub struct Workbook {
    pub sheets: Vec<String>,
    /// (supbook index, first sheet index, last sheet index) per BrtExternSheet entry.
    pub xti: Vec<(i32, i32, i32)>,
    pub supbooks: Vec<SupbookKind>,
    pub suptabs: std::collections::HashMap<i32, Vec<String>>,
    pub names: Vec<DefinedName>,
    /// BrtBundleSh relationship ids, in the same order as `sheets`.
    pub relids: Vec<Option<String>>,
}

impl Workbook {
    /// 1-based index among external supbooks, as Excel shows `[n]`.
    pub fn ext_index(&self, isup: i32) -> i32 {
        let mut k = 0;
        for (i, kind) in self.supbooks.iter().enumerate() {
            if *kind == SupbookKind::Ext {
                k += 1;
            }
            if i as i32 == isup {
                return k;
            }
        }
        0
    }

    pub fn sheet_prefix(&self, ixti: usize) -> Result<String> {
        let (isup, first, last) = *self
            .xti
            .get(ixti)
            .ok_or(DecodeError("xti index out of range"))?;
        let kind = self
            .supbooks
            .get(isup as usize)
            .copied()
            .unwrap_or(SupbookKind::SelfBook);
        if kind == SupbookKind::SelfBook || kind == SupbookKind::Same {
            if first == -2 {
                return Ok(String::new()); // workbook-level (names)
            }
            if first < 0 || first as usize >= self.sheets.len() {
                return Ok("#REF!!".to_string());
            }
            let a = &self.sheets[first as usize];
            if last != first && last >= 0 && (last as usize) < self.sheets.len() {
                return Ok(quote_sheet(&format!("{a}:{}", self.sheets[last as usize])) + "!");
            }
            return Ok(quote_sheet(a) + "!");
        }
        // external workbook
        let n = self.ext_index(isup);
        let tabs = self.suptabs.get(&isup).cloned().unwrap_or_default();
        let a = if first >= 0 && (first as usize) < tabs.len() {
            tabs[first as usize].clone()
        } else {
            String::new()
        };
        if last != first && last >= 0 && (last as usize) < tabs.len() {
            return Ok(quote_sheet(&format!("[{n}]{a}:{}", tabs[last as usize])) + "!");
        }
        if !a.is_empty() {
            return Ok(quote_sheet(&format!("[{n}]{a}")) + "!");
        }
        Ok(format!("[{n}]!"))
    }
}

static PLAIN_SHEET: Lazy<Regex> = Lazy::new(|| Regex::new(r"^[A-Za-z_\\][A-Za-z0-9_.]*$").unwrap());
static LOOKS_LIKE_REF: Lazy<Regex> =
    Lazy::new(|| Regex::new(r"(?i)^([A-Za-z]{1,3}[0-9]+|R[0-9]*C[0-9]*|[Rr]|[Cc])$").unwrap());

pub fn quote_sheet(name: &str) -> String {
    let upper = name.to_ascii_uppercase();
    if PLAIN_SHEET.is_match(name)
        && !LOOKS_LIKE_REF.is_match(name)
        && !upper.starts_with("TRUE")
        && !upper.starts_with("FALSE")
    {
        return name.to_string();
    }
    format!("'{}'", name.replace('\'', "''"))
}

/// Parses `xl/workbook.bin`: sheets/relids (BrtBundleSh), supbook table
/// (BrtSupSelf/Same/BookSrc/AddIn + BrtExternSheetTable... via BrtExternSheet),
/// per-supbook external sheet names (BrtSupTabs) and defined names (BrtName).
pub fn parse_workbook(data: &[u8]) -> Result<Workbook> {
    let mut wb = Workbook::default();
    let mut cur_sup: i32 = -1;
    for rec in Records::new(data) {
        let rec = rec?;
        let pl = rec.payload;
        if rec.id == R_BUNDLESH {
            let pos = 8;
            let (relid, pos) = read_wide(pl, pos)?;
            let (name, _pos) = read_wide(pl, pos)?;
            wb.sheets.push(name.unwrap_or_default());
            wb.relids.push(relid);
        } else if rec.id == R_SUPSELF {
            wb.supbooks.push(SupbookKind::SelfBook);
            cur_sup = wb.supbooks.len() as i32 - 1;
        } else if rec.id == R_SUPSAME {
            wb.supbooks.push(SupbookKind::Same);
            cur_sup = wb.supbooks.len() as i32 - 1;
        } else if rec.id == R_SUPBOOKSRC {
            wb.supbooks.push(SupbookKind::Ext);
            cur_sup = wb.supbooks.len() as i32 - 1;
        } else if rec.id == R_SUPADDIN {
            wb.supbooks.push(SupbookKind::Addin);
            cur_sup = wb.supbooks.len() as i32 - 1;
        } else if rec.id == R_SUPTABS {
            let cnt = get_u32(pl, 0)?;
            let mut pos = 4usize;
            let mut tabs = Vec::with_capacity(cnt.min(1 << 20) as usize);
            for _ in 0..cnt {
                let (s, next_pos) = read_wide(pl, pos)?;
                tabs.push(s.unwrap_or_default());
                pos = next_pos;
            }
            wb.suptabs.insert(cur_sup, tabs);
        } else if rec.id == R_EXTERNSHEET {
            let cnt = get_u32(pl, 0)? as usize;
            for i in 0..cnt {
                let o = 4 + 12 * i;
                let isup = get_u32(pl, o)? as i32;
                let first = get_i32(pl, o + 4)?;
                let last = get_i32(pl, o + 8)?;
                wb.xti.push((isup, first, last));
            }
        } else if rec.id == R_NAME {
            let itab = get_i32(pl, 5)?;
            let (name, pos) = read_wide(pl, 9)?;
            let cce = get_u32(pl, pos)? as usize;
            let pos = pos + 4;
            let rgce = pl
                .get(pos..pos + cce)
                .ok_or(DecodeError("BrtName rgce overruns payload"))?
                .to_vec();
            let pos = pos + cce;
            let cb = get_u32(pl, pos)? as usize;
            let pos = pos + 4;
            let rgcb = pl
                .get(pos..pos + cb)
                .ok_or(DecodeError("BrtName rgcb overruns payload"))?
                .to_vec();
            wb.names.push(DefinedName {
                name: name.unwrap_or_default(),
                itab,
                rgce,
                rgcb,
            });
        }
    }
    Ok(wb)
}

/// Parses `xl/_rels/workbook.bin.rels` into an `Id -> Target` map, matching
/// on `<Relationship .../>` elements regardless of attribute order.
pub fn parse_workbook_rels(xml: &[u8]) -> Result<std::collections::HashMap<String, String>> {
    use quick_xml::events::Event;
    use quick_xml::reader::Reader;

    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut map = std::collections::HashMap::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(Event::Empty(e)) | Ok(Event::Start(e)) => {
                if e.local_name().as_ref() == "Relationship" {
                    let mut id = None;
                    let mut target = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            "Id" => id = Some(attr.value.into_owned()),
                            "Target" => target = Some(attr.value.into_owned()),
                            _ => {}
                        }
                    }
                    if let (Some(id), Some(target)) = (id, target) {
                        map.insert(id, target);
                    }
                }
            }
            Ok(_) => {}
            Err(_) => return Err(DecodeError("malformed workbook rels xml")),
        }
        buf.clear();
    }
    Ok(map)
}
