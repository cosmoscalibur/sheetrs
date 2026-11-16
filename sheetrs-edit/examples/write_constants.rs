//! Writes constants into a workbook through `BookEdit::write_cells`, for the oracle of the writers
//! (`tests/assets/writer_results/make_oracle.py` opens the file it makes in `LibreOffice`).
//!
//! ```text
//! cargo run -p sheetrs-edit --example write_constants -- IN OUT SHEET < cells.json
//! ```
//!
//! `cells.json` is a list of `{"row": 0, "col": 0, "kind": "number|text|boolean|empty", "value": ...}`
//! (0-based). A number may be given as a text (`"NaN"`, `"inf"`, `"-inf"`) for the values JSON cannot hold.

use serde_json::Value;
use sheetrs_edit::{BookEdit, CellWrites, Constant};
use sheetrs_io::Book;
use std::io::Read;

fn number(value: &Value) -> f64 {
    match value {
        Value::String(text) => match text.as_str() {
            "NaN" => f64::NAN,
            "inf" => f64::INFINITY,
            "-inf" => f64::NEG_INFINITY,
            other => other.parse().expect("a number"),
        },
        other => other.as_f64().expect("a number"),
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = std::env::args().collect();
    let [_, input, output, sheet] = args.as_slice() else {
        return Err("usage: write_constants IN OUT SHEET < cells.json".into());
    };
    let mut text = String::new();
    std::io::stdin().read_to_string(&mut text)?;
    let cells: Vec<Value> = serde_json::from_str(&text)?;
    let mut writes = CellWrites::new();
    for cell in &cells {
        let row = u32::try_from(cell["row"].as_u64().ok_or("row")?)?;
        let col = u32::try_from(cell["col"].as_u64().ok_or("col")?)?;
        let value = &cell["value"];
        let constant = match cell["kind"].as_str().ok_or("kind")? {
            "number" => Constant::Number(number(value)),
            "text" => Constant::Text(value.as_str().ok_or("text")?.to_string()),
            "boolean" => Constant::Boolean(value.as_bool().ok_or("boolean")?),
            _ => Constant::Empty,
        };
        writes.constant(sheet.as_str(), row, col, constant);
    }
    let mut book = Book::from_file(input)?;
    book.write_cells(writes);
    book.write(output)?;
    Ok(())
}
