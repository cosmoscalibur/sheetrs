//! Which formulas of an ODS `LibreOffice` calculates when it opens the file, whatever result the file caches, as the reader tells
//! them to its sink (`CellSink::recalculated_on_load`). The rule is the one observed with `LibreOffice` 26.8
//! (`sheetrs-edit/tests/assets/write_to/probe_ods_recalculation.py`): the type of the cached result and the number format of the cell.

use std::fmt::Write as _;
use std::io::{Cursor, Write};

use sheetrs_io::Book;
use sheetrs_io::reader::{Cell, CellSink, Properties, Selection, Window};
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

#[derive(Default)]
struct Flags {
    next: bool,
    first_only: bool,
    flagged: Vec<u32>,
}

impl CellSink for Flags {
    fn recalculated_on_load(&mut self, _sheet: u16, _at: (u32, u32)) {
        self.next = true;
    }
    fn wants_more(&self) -> bool {
        !(self.first_only && self.next)
    }
    fn cell(&mut self, _sheet: u16, at: (u32, u32), _cell: Cell) {
        if std::mem::take(&mut self.next) {
            self.flagged.push(at.1);
        }
    }
}

const FLOAT: &str = r#"office:value-type="float" office:value="999""#;
const PERCENT: &str = r#"office:value-type="percentage" office:value="9.99""#;
const BOOLEAN: &str = r#"office:value-type="boolean" office:boolean-value="false""#;
const TIME: &str = r#"office:value-type="time" office:time-value="PT05H00M00S""#;
const TEXT: &str = r#"office:value-type="string" office:string-value="S""#;
const ERROR: &str =
    r#"office:value-type="string" office:string-value="" calcext:value-type="error""#;
const COLUMNS: &str = r#"<table:table-column table:number-columns-repeated="8"/>"#;

/// The columns of a table with the columns `columns` and one row (with the attributes `row`) of formula cells (the attributes of each
/// in `cells`) that the reader says `LibreOffice` calculates on load.
fn flagged(columns: &str, row: &str, cells: &[&str]) -> Vec<u32> {
    flagged_until(columns, row, cells, false)
}

fn flagged_until(columns: &str, row: &str, cells: &[&str], first_only: bool) -> Vec<u32> {
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let mut part = |name: &str, text: String| {
        zip.start_file(name, SimpleFileOptions::default()).unwrap();
        zip.write_all(text.as_bytes()).unwrap();
    };
    part(
        "mimetype",
        "application/vnd.oasis.opendocument.spreadsheet".to_string(),
    );
    part(
        "META-INF/manifest.xml",
        r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#.to_string(),
    );
    let mut cell_xml = String::new();
    for c in cells {
        write!(
            cell_xml,
            r#"<table:table-cell table:formula="of:=1+1" {c}><text:p>x</text:p></table:table-cell>"#
        )
        .unwrap();
    }
    let cells = cell_xml;
    part(
        "content.xml",
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:version="1.3"><office:automatic-styles><number:number-style style:name="N2"><number:number number:decimal-places="2" number:min-integer-digits="1"/></number:number-style><style:style style:name="two" style:family="table-cell" style:data-style-name="N2"/><style:style style:name="plain" style:family="table-cell"/></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="T">{columns}<table:table-row{row}>{cells}</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
        ),
    );
    let book = Book::from_bytes(zip.finish().unwrap().into_inner()).unwrap();
    let mut flags = Flags {
        first_only,
        ..Flags::default()
    };
    book.selection_into(
        &Selection::new(Properties::VALUE).window_everywhere(Window::ALL),
        &mut flags,
    )
    .unwrap();
    flags.flagged
}

#[test]
fn a_float_without_format_is_calculated_on_load_other_types_are_kept() {
    assert_eq!(
        flagged(COLUMNS, "", &[FLOAT, PERCENT, BOOLEAN, TIME, TEXT]),
        [0]
    );
}

#[test]
fn a_currency_a_date_and_an_untyped_result_are_calculated_on_load() {
    let currency = r#"office:value-type="currency" office:currency="USD" office:value="999""#;
    let date = r#"office:value-type="date" office:date-value="1999-01-01""#;
    assert_eq!(flagged(COLUMNS, "", &[currency, date, ""]), [0, 1, 2]);
}

#[test]
fn a_cached_error_is_calculated_on_load_even_in_a_cell_with_a_style() {
    let styled = format!(r#"table:style-name="two" {ERROR}"#);
    assert_eq!(flagged(COLUMNS, "", &[ERROR, &styled]), [0, 1]);
}

#[test]
fn a_cell_with_a_style_of_its_own_keeps_what_it_caches() {
    let two = format!(r#"table:style-name="two" {FLOAT}"#);
    let plain = format!(r#"table:style-name="plain" {FLOAT}"#);
    assert_eq!(flagged(COLUMNS, "", &[&two, &plain]), Vec::<u32>::new());
}

#[test]
fn a_cell_takes_the_number_format_of_its_row_before_the_one_of_its_column() {
    let formatted_column = r#"<table:table-column table:default-cell-style-name="two"/>"#;
    let general_column = r#"<table:table-column table:default-cell-style-name="plain"/>"#;
    assert_eq!(flagged(formatted_column, "", &[FLOAT]), Vec::<u32>::new());
    assert_eq!(flagged(general_column, "", &[FLOAT]), [0]);
    let formatted_row = r#" table:default-cell-style-name="two""#;
    let general_row = r#" table:default-cell-style-name="plain""#;
    assert_eq!(
        flagged(general_column, formatted_row, &[FLOAT]),
        Vec::<u32>::new()
    );
    assert_eq!(flagged(formatted_column, general_row, &[FLOAT]), [0]);
}

#[test]
fn a_cell_in_an_undeclared_column_is_calculated_unless_it_caches_text() {
    // The first cell is in the one column the table declares; the others are past it.
    assert_eq!(
        flagged("<table:table-column/>", "", &[PERCENT, BOOLEAN, TEXT]),
        [1]
    );
    assert_eq!(flagged("", "", &[PERCENT, TEXT]), [0]);
}

#[test]
fn an_array_formula_is_calculated_on_load_whatever_it_caches() {
    let array = format!(
        r#"table:number-matrix-columns-spanned="1" table:number-matrix-rows-spanned="1" {TEXT}"#
    );
    assert_eq!(flagged(COLUMNS, "", &[&array]), [0]);
}

#[test]
fn of_a_run_of_repeated_cells_the_first_is_calculated_on_load() {
    let repeated = format!(r#"table:number-columns-repeated="3" {FLOAT}"#);
    assert_eq!(flagged(COLUMNS, "", &[&repeated]), [0]);
}

#[test]
fn a_sink_that_wants_no_more_stops_at_the_first_calculated_cell() {
    let cells = [FLOAT, FLOAT, FLOAT];
    assert_eq!(flagged_until(COLUMNS, "", &cells, true), [0]);
    assert_eq!(flagged_until(COLUMNS, "", &cells, false), [0, 1, 2]);
}
