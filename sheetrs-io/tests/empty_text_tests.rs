//! An empty text written in every way a file can write it, read as `LibreOffice` 26.8 reads it:
//! `tests/assets/empty_text.{xlsx,ods}` (see `oracle/make_empty_text.py`). Sheet `T` holds the cells, and
//! row `i` of sheet `Calc` what `LibreOffice` answered about `T!A<i>` when it opened that very file
//! (`ISBLANK` in column A, `ISTEXT` in column E).
//!
//! The two formats differ, and the readers follow each: an ODS text cell with a paragraph, even an empty
//! one, is an empty text; with none (`office:string-value=""` alone) it is blank. In an XLSX every empty
//! text of a shared string, an inline string or a `t="str"` cell is a blank cell.

use std::collections::HashMap;
use std::path::Path;

use sheetrs_io::Book;
use sheetrs_io::reader::{Cell, CellSink, CellValue};

#[derive(Default)]
struct Collect {
    cells: HashMap<(u16, u32, u32), Cell>,
}

impl CellSink for Collect {
    fn cell(&mut self, sheet: u16, at: (u32, u32), cell: Cell) {
        self.cells.insert((sheet, at.0, at.1), cell);
    }
}

fn cells_of(file: &str) -> Collect {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets")
        .join(file);
    let mut sink = Collect::default();
    Book::from_file(&path)
        .unwrap()
        .workbook_into(&mut sink)
        .unwrap();
    sink
}

fn is_true(cells: &Collect, row: u32, col: u32) -> bool {
    matches!(
        cells.cells.get(&(1, row, col)).map(|c| &c.value),
        Some(CellValue::Boolean(true))
    )
}

fn each_cell(file: &str, rows: u32) {
    let cells = cells_of(file);
    for row in 0..rows {
        let blank = is_true(&cells, row, 0);
        let text = is_true(&cells, row, 4);
        let at = cells.cells.get(&(0, row, 0));
        let value = at.map_or(&CellValue::Empty, |c| &c.value);
        let formula = at.is_some_and(|c| c.formula.is_some());
        let label = format!("{file} T!A{}", row + 1);
        if blank {
            assert!(
                matches!(value, CellValue::Empty),
                "{label}: LibreOffice says blank, the reader gives {value:?}"
            );
        } else if text {
            // A formula with no cached result has no value to give, whatever it calculates.
            if formula && matches!(value, CellValue::Empty) {
                continue;
            }
            assert!(
                matches!(value, CellValue::Text(_)),
                "{label}: LibreOffice says text, the reader gives {value:?}"
            );
        }
    }
}

#[test]
fn an_xlsx_empty_text_is_blank_as_libreoffice_reads_it() {
    each_cell("empty_text.xlsx", 11);
}

#[test]
fn an_ods_empty_text_is_blank_or_empty_text_as_libreoffice_reads_it() {
    each_cell("empty_text.ods", 11);
}

#[test]
fn the_cells_that_hold_an_empty_text_in_each_format_are_the_ones_libreoffice_says() {
    // The expectation spelled out, row by row (the cases of `make_empty_text.py`).
    let ods = cells_of("empty_text.ods");
    let kind = |c: &Collect, row: u32| match c.cells.get(&(0, row, 0)).map(|x| &x.value) {
        None | Some(CellValue::Empty) => "blank",
        Some(CellValue::Text(t)) if t.is_empty() => "empty text",
        Some(CellValue::Text(_)) => "text",
        Some(_) => "other",
    };
    let ods_rows: Vec<&str> = (0..8).map(|r| kind(&ods, r)).collect();
    assert_eq!(
        ods_rows,
        [
            "empty text", // an empty paragraph
            "empty text", // string-value empty and an empty paragraph
            "blank",      // string-value empty, no paragraph
            "blank",      // no value, no paragraph
            "empty text", // start and end tags
            "empty text", // with the calcext type
            "blank",      // nothing
            "empty text", // formula ="" with its empty result
        ]
    );
    let xlsx = cells_of("empty_text.xlsx");
    let xlsx_rows: Vec<&str> = (0..8).map(|r| kind(&xlsx, r)).collect();
    assert_eq!(
        xlsx_rows,
        [
            "blank",
            "blank",
            "blank",
            "blank",
            "blank",
            "blank",
            "blank",
            "empty text", // formula ="" with its empty result
        ]
    );
}
