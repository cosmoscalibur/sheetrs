//! Writes every cell of an ODS file to the standard output, one per line, for
//! comparing a reader with the oracle on a workbook that must not be committed. Nothing but the
//! tab-separated lines is printed.

use std::env;
use std::error::Error;
use std::fs::File;
use std::io::{BufWriter, Write};

use sheetrs_core::{FormulaRef, ValueRef, Workbook};

fn escape(text: &str) -> String {
    text.replace('\\', "\\\\")
        .replace('\n', "\\n")
        .replace('\t', "\\t")
        .replace('\r', "\\r")
}

fn main() -> Result<(), Box<dyn Error>> {
    let path = env::args().nth(1).ok_or("usage: dump_cells <file.ods>")?;
    let workbook: Workbook = sheetrs_ods::read(File::open(&path)?)?;
    let mut out = BufWriter::new(std::io::stdout().lock());
    for (index, sheet) in workbook.sheets().enumerate() {
        for (pos, cell) in sheet.cells() {
            let (kind, value) = match &cell.value {
                ValueRef::Number(number) => ("n", format!("{number:?}")),
                ValueRef::Bool(flag) => ("n", format!("{:?}", f64::from(u8::from(*flag)))),
                ValueRef::Text("") if cell.formula.is_none() => continue,
                ValueRef::Text(text) => ("s", escape(text)),
                ValueRef::Error(kind) => ("e", kind.as_text().to_owned()),
                ValueRef::Empty => ("-", String::new()),
            };
            // 1 is a formula with its own text, 2 a member of a shared group.
            let formula = match cell.formula {
                Some(FormulaRef::Text(_)) => 1,
                Some(FormulaRef::Shared { .. }) => 2,
                None => 0,
            };
            let stale = u8::from(cell.stale);
            writeln!(
                out,
                "{index}\t{}\t{}\t{kind}\t{value}\t{formula}\t{stale}",
                pos.row, pos.col
            )?;
        }
    }
    Ok(())
}
