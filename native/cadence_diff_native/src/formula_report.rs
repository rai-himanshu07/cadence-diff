//! PyO3-facing entry points for the B0/B1/B2 parity proofs. Returns per-sheet
//! decoded data as plain tuples/columnar arrays (never a custom PyO3 class,
//! to keep the FFI surface trivial for these feasibility experiments). The
//! caller (a private, aggregate-only probe script) compares this against
//! the Python reference (`ptg_decoder.py` for formulas, `pyxlsb` for values)
//! and never persists or prints the underlying text/values themselves.

use crate::ptg::Mode;
use pyo3::exceptions::PyRuntimeError;
use pyo3::prelude::*;

/// `data` is the raw XLSB file bytes (never a path -- see
/// `sheet::decode_workbook`'s doc comment). Per sheet: `(name, rows, cols,
/// texts, unresolved_follower_count, [(error_message, count), ...])`.
/// `rows[i]/cols[i]/texts[i]` line up positionally for every successfully
/// decoded formula cell; a cell missing from these arrays failed closed
/// (counted in `unresolved_follower_count` or the error list instead of guessed).
/// `mode` is `"r1c1"` or `"a1"`.
#[pyfunction]
pub fn formula_r1c1_report(
    data: &[u8],
    mode: &str,
) -> PyResult<
    Vec<(
        String,
        Vec<u32>,
        Vec<u32>,
        Vec<String>,
        u32,
        Vec<(String, u32)>,
    )>,
> {
    let mode = match mode {
        "r1c1" => Mode::R1C1,
        "a1" => Mode::A1,
        _ => return Err(PyRuntimeError::new_err("mode must be 'r1c1' or 'a1'")),
    };
    let (_wb, reports) =
        crate::sheet::decode_workbook(data, mode).map_err(|e| PyRuntimeError::new_err(e.0))?;
    let mut out = Vec::with_capacity(reports.len());
    for report in reports {
        let mut rows = Vec::with_capacity(report.cells.len());
        let mut cols = Vec::with_capacity(report.cells.len());
        let mut texts = Vec::with_capacity(report.cells.len());
        for ((row, col), text) in report.cells {
            rows.push(row);
            cols.push(col);
            texts.push(text);
        }
        let errors: Vec<(String, u32)> = report.errors.into_iter().collect();
        out.push((
            report.name,
            rows,
            cols,
            texts,
            report.unresolved_followers,
            errors,
        ));
    }
    Ok(out)
}

/// `data` is the raw XLSB file bytes (never a path). Per sheet: `(name,
/// [(row, col, num, boolean, text), ...])` -- exactly one of
/// `num`/`boolean`/`text` is populated per cell, mirroring pyxlsb's own raw
/// `.v` shape (plain float / bool / str, blanks never emitted) so the Python
/// side can feed it through the existing, unchanged `_xlsb_value`.
#[pyfunction]
pub fn raw_values_report(
    data: &[u8],
) -> PyResult<
    Vec<(
        String,
        Vec<(u32, u32, Option<f64>, Option<bool>, Option<String>)>,
    )>,
> {
    let reports =
        crate::values::decode_workbook_values(data).map_err(|e| PyRuntimeError::new_err(e.0))?;
    Ok(reports
        .into_iter()
        .map(|(name, cells)| {
            let rows = cells
                .into_iter()
                .map(|c| (c.row, c.col, c.num, c.boolean, c.text))
                .collect();
            (name, rows)
        })
        .collect())
}

/// B2's production-shaped surface. Per sheet:
/// `(name, cell_rows, cell_cols, cell_a1_text, cell_definition_id,
/// definition_r1c1_table)` -- `cell_definition_id[i]` indexes into
/// `definition_r1c1_table` (deduplicated: every follower of one shared/array
/// group shares one entry). Plus workbook-level defined names:
/// `(name, target_a1_or_none)`.
#[pyfunction]
pub fn formula_surface_report(
    data: &[u8],
) -> PyResult<(
    Vec<(
        String,
        Vec<u32>,
        Vec<u32>,
        Vec<String>,
        Vec<u32>,
        Vec<String>,
    )>,
    Vec<(String, Option<String>)>,
)> {
    let surface =
        crate::surface::decode_workbook_surface(data).map_err(|e| PyRuntimeError::new_err(e.0))?;
    let sheets = surface
        .sheets
        .into_iter()
        .map(|s| {
            (
                s.name,
                s.cell_rows,
                s.cell_cols,
                s.cell_a1,
                s.cell_definition_id,
                s.definition_r1c1,
            )
        })
        .collect();
    let names = surface
        .defined_names
        .into_iter()
        .map(|n| (n.name, n.target_a1))
        .collect();
    Ok((sheets, names))
}
