//! sheetrs-answer: the answer of a compiled program, written into the workbook it was made from.
//!
//! A workbook compiled in a workbook mode (`sheetc build --embed-template`, or `sheetc calculate -o`)
//! is its own template: [`fill_template`] runs the program and writes what it computed into a copy of
//! the workbook, keeping its formulas, styles and everything else, with the inputs that were given in
//! their cells. It sits between the engine (the program and its run) and `sheetrs-io` (the writers of
//! XLSX and ODS), and is what the compiler, the embedded runtimes and the WebAssembly modules use to
//! answer with a workbook; the plan of an editor is not part of it.

mod error;
mod template;

pub use error::{Error, Result};
pub use template::{
    fill_template, fill_template_to, fill_template_to_also, fill_template_to_with_answer,
    fill_template_with_answer,
};

/// What a template was filled with: the number of cells written.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Written {
    /// The cells written (results of formulas and inputs).
    pub cells: usize,
}
