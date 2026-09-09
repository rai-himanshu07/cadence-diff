//! PyO3 extension module entry point. Registers the B0/B1/B2 kernel
//! functions (`native/xlsbkernel/`'s owned, bounds-checked BIFF12 reader --
//! records/workbook/ptg/sheet/values/surface -- see each module's own doc
//! comment). No longer depends on `calamine`: the initial exploratory
//! `formula_stats`/`values_columnar` functions this crate started from
//! (calamine-backed) were superseded by B0's owned formula reader and B1's
//! owned values reader (calamine auto-converts date-formatted cells in a
//! way pyxlsb never does -- see the B1 Execution Log entry -- so values
//! needed an owned reader too); those two functions and the `calamine`
//! dependency were removed once nothing called them anymore.
use pyo3::prelude::*;

mod formula_report;
mod ftab;
mod ptg;
mod records;
mod sheet;
mod surface;
mod values;
mod workbook;

#[pymodule]
fn xlsbkernel(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(formula_report::formula_r1c1_report, m)?)?;
    m.add_function(wrap_pyfunction!(formula_report::raw_values_report, m)?)?;
    m.add_function(wrap_pyfunction!(formula_report::formula_surface_report, m)?)?;
    Ok(())
}
