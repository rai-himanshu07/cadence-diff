//! Function table (FTAB): built-in function names + declared argument counts
//! for `PtgFunc`/`PtgFuncVar`. Same source data as `ptg_decoder.py`'s
//! `ftab.json`, embedded at compile time so the wheel needs no data file on
//! disk.

use serde::Deserialize;
use std::sync::OnceLock;

#[derive(Deserialize)]
struct RawFtab {
    ftab: Vec<String>,
    argc: Vec<u32>,
}

static FTAB: OnceLock<RawFtab> = OnceLock::new();

fn table() -> &'static RawFtab {
    FTAB.get_or_init(|| {
        let raw = include_str!("../ftab.json");
        serde_json::from_str(raw)
            .expect("ftab.json must parse: embedded at build time, not hostile input")
    })
}

/// Function name for a given FTAB index, or a placeholder if out of range
/// (mirrors `ptg_decoder.py`'s `f"#FUNC{iftab}#"` fallback).
pub fn name(iftab: u32) -> String {
    table()
        .ftab
        .get(iftab as usize)
        .cloned()
        .unwrap_or_else(|| format!("#FUNC{iftab}#"))
}

/// Declared argument count for `PtgFunc` (fixed-arity) calls; 0 if unknown.
pub fn argc(iftab: u32) -> u32 {
    table().argc.get(iftab as usize).copied().unwrap_or(0)
}
