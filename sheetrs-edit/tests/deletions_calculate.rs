//! After a deletion the engine calculates every formula again, so what a cell holds is what its formula
//! gives now (`#REF!` for a deleted reference, the new sum of a range that shrank), whatever the
//! workbook cached. The workbooks are made here, with the results they cache, and what is asserted is
//! the value each cell has in the output.

use sheetrs_edit::BookEdit;
use sheetrs_io::Book;
use sheetrs_io::reader::CellValue;
use std::io::Write;
use std::path::{Path, PathBuf};

const EXTENSIONS: [&str; 2] = ["xlsx", "ods"];

/// A cell of the test workbooks: a constant, or a formula with the result it caches.
enum Cell {
    Number(f64),
    /// The formula in the notation of the file (Excel's for an XLSX, `OpenFormula` for an ODS), and
    /// the result it caches.
    Formula(String, f64),
}

use Cell::Number as N;

/// A formula, written in the notation of the format of `extension`: `excel` for an XLSX and `ods`
/// (the text after `of:=`) for an ODS.
fn f(extension: &str, excel: &str, ods: &str, cached: f64) -> Cell {
    let text = if extension == "xlsx" { excel } else { ods };
    Cell::Formula(text.to_string(), cached)
}

/// The sheets: a name and the cells of its first column.
type Sheets<'a> = &'a [(&'a str, Vec<Cell>)];

fn write_xlsx(path: &Path, sheets: Sheets<'_>) {
    let mut zip = zip::ZipWriter::new(std::fs::File::create(path).unwrap());
    let options = zip::write::SimpleFileOptions::default();
    let mut part = |name: &str, text: &str| {
        zip.start_file(name, options).unwrap();
        zip.write_all(text.as_bytes()).unwrap();
    };
    let overrides: String = (1..=sheets.len())
        .map(|i| format!(r#"<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#))
        .collect::<Vec<String>>().concat();
    part(
        "[Content_Types].xml",
        &format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>{overrides}</Types>"#
        ),
    );
    part(
        "_rels/.rels",
        r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
    );
    let names: String = sheets
        .iter()
        .enumerate()
        .map(|(i, (name, _))| {
            format!(
                r#"<sheet name="{name}" sheetId="{id}" r:id="rId{id}"/>"#,
                id = i + 1
            )
        })
        .collect::<Vec<String>>()
        .concat();
    part(
        "xl/workbook.xml",
        &format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>{names}</sheets></workbook>"#
        ),
    );
    let rels: String = (1..=sheets.len())
        .map(|i| format!(r#"<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>"#))
        .collect::<Vec<String>>().concat();
    part(
        "xl/_rels/workbook.xml.rels",
        &format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
        ),
    );
    for (i, (_, cells)) in sheets.iter().enumerate() {
        let rows: String = cells
            .iter()
            .enumerate()
            .map(|(r, cell)| {
                let at = format!("A{}", r + 1);
                let inner = match cell {
                    N(n) => format!("<v>{n}</v>"),
                    Cell::Formula(text, cached) => format!("<f>{text}</f><v>{cached}</v>"),
                };
                format!(r#"<row r="{}"><c r="{at}">{inner}</c></row>"#, r + 1)
            })
            .collect::<Vec<String>>()
            .concat();
        part(
            &format!("xl/worksheets/sheet{}.xml", i + 1),
            &format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{rows}</sheetData></worksheet>"#
            ),
        );
    }
    zip.finish().unwrap();
}

fn write_ods(path: &Path, sheets: Sheets<'_>) {
    let mut zip = zip::ZipWriter::new(std::fs::File::create(path).unwrap());
    zip.start_file(
        "mimetype",
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored),
    )
    .unwrap();
    zip.write_all(b"application/vnd.oasis.opendocument.spreadsheet")
        .unwrap();
    let options = zip::write::SimpleFileOptions::default();
    zip.start_file("META-INF/manifest.xml", options).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#).unwrap();
    let tables: String = sheets
        .iter()
        .map(|(name, cells)| {
            let rows: String = cells
                .iter()
                .map(|cell| {
                    let attributes = match cell {
                        N(n) => format!(r#"office:value-type="float" office:value="{n}""#),
                        Cell::Formula(text, cached) => format!(
                            r#"table:formula="of:={text}" office:value-type="float" office:value="{cached}""#
                        ),
                    };
                    let shown = match cell {
                        N(n) | Cell::Formula(_, n) => n,
                    };
                    format!("<table:table-row><table:table-cell {attributes}><text:p>{shown}</text:p></table:table-cell></table:table-row>")
                })
                .collect::<Vec<String>>().concat();
            format!(r#"<table:table table:name="{name}">{rows}</table:table>"#)
        })
        .collect::<Vec<String>>()
        .concat();
    zip.start_file("content.xml", options).unwrap();
    zip.write_all(format!(r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:version="1.3"><office:body><office:spreadsheet>{tables}</office:spreadsheet></office:body></office:document-content>"#).as_bytes()).unwrap();
    zip.finish().unwrap();
}

/// The workbook in the format of `extension`, in `dir`.
fn workbook(dir: &Path, extension: &str, sheets: Sheets<'_>) -> PathBuf {
    let path = dir.join(format!("book.{extension}"));
    if extension == "xlsx" {
        write_xlsx(&path, sheets);
    } else {
        write_ods(&path, sheets);
    }
    path
}

fn value(path: &Path, sheet: &str, row: u32) -> CellValue {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let sheet = workbook.sheets.iter().find(|s| s.name == sheet).unwrap();
    sheet
        .cells
        .iter()
        .find(|(at, _)| (at.0, at.1) == (row, 0))
        .map_or(CellValue::Empty, |(_, cell)| cell.value.clone())
}

fn number(path: &Path, sheet: &str, row: u32) -> f64 {
    match value(path, sheet, row) {
        CellValue::Number(n) => n,
        other => panic!("{sheet} row {row} is {other:?}, not a number"),
    }
}

fn is_ref_error(path: &Path, sheet: &str, row: u32) -> bool {
    matches!(
        value(path, sheet, row),
        CellValue::Error(sheetrs_io::reader::ExcelError::Ref)
    )
}

/// Data in `D`, and `F` with formulas that read it: one that reads a cell, one that reads that
/// formula, one that reads a cell below, and one that reads the cell above the one that is deleted.
fn reading_a_cell(e: &str) -> Vec<(&'static str, Vec<Cell>)> {
    vec![
        ("D", vec![N(10.0), N(20.0), N(30.0)]),
        (
            "F",
            vec![
                f(e, "D!A2", "[$D.A2]", 20.0),
                f(e, "A1*2", "[.A1]*2", 40.0),
                f(e, "D!A3+1", "[$D.A3]+1", 31.0),
                f(e, "D!A1", "[$D.A1]", 10.0),
            ],
        ),
        ("U", vec![N(7.0)]),
    ]
}

fn editor(path: &Path) -> Book {
    Book::from_file(path).unwrap()
}

#[test]
fn deleting_a_sheet_that_nothing_reads_calculates_every_formula_and_gives_the_same_results() {
    for extension in EXTENSIONS {
        let dir = tempfile::tempdir().unwrap();
        let input = workbook(dir.path(), extension, &reading_a_cell(extension));
        let output = dir.path().join(format!("out.{extension}"));
        let mut book = editor(&input);
        book.delete_sheets(["U"]);
        let plan = book.explain();
        assert!(plan.contains("pass 2 of 2"), "{extension}: {plan}");
        assert!(plan.contains("recalculate"), "{extension}: {plan}");
        book.write(&output).unwrap();
        for (row, cached) in [(0, 20.0), (1, 40.0), (2, 31.0), (3, 10.0)] {
            assert!(
                (number(&output, "F", row) - cached).abs() < 1e-9,
                "{extension} F row {row}"
            );
        }
    }
}

#[test]
fn converting_every_formula_and_deleting_rows_keeps_the_values_the_workbook_had() {
    for extension in EXTENSIONS {
        let dir = tempfile::tempdir().unwrap();
        let input = workbook(dir.path(), extension, &reading_a_cell(extension));
        let output = dir.path().join(format!("out.{extension}"));
        let mut book = editor(&input);
        book.delete_rows(["D!2".parse().unwrap()])
            .formula_to_value(["F".parse().unwrap()]);
        book.write(&output).unwrap();
        // The values the workbook had, as the user saw them when asking.
        assert!((number(&output, "F", 0) - 20.0).abs() < 1e-9, "{extension}");
        assert!((number(&output, "F", 1) - 40.0).abs() < 1e-9, "{extension}");
    }
}

#[test]
fn a_formula_that_reads_a_deleted_cell_is_given_the_reference_error() {
    for extension in EXTENSIONS {
        let dir = tempfile::tempdir().unwrap();
        let input = workbook(dir.path(), extension, &reading_a_cell(extension));
        let output = dir.path().join(format!("out.{extension}"));
        let mut book = editor(&input);
        book.delete_rows(["D!2".parse().unwrap()]);
        book.write(&output).unwrap();
        // The cell that is read and the formula that reads that formula.
        assert!(is_ref_error(&output, "F", 0), "{extension}");
        assert!(is_ref_error(&output, "F", 1), "{extension}");
        // A formula that reads a cell below only moved with it; the one above reads what it read.
        assert!((number(&output, "F", 2) - 31.0).abs() < 1e-9, "{extension}");
        assert!((number(&output, "F", 3) - 10.0).abs() < 1e-9, "{extension}");
    }
}

/// A formula that reads a cell, a range that shrinks and a formula that traps the error.
fn reading_a_range(e: &str) -> Vec<(&'static str, Vec<Cell>)> {
    vec![
        ("D", vec![N(10.0), N(20.0), N(30.0)]),
        (
            "F",
            vec![
                f(e, "D!A2", "[$D.A2]", 20.0),
                f(e, "SUM(D!A1:A3)", "SUM([$D.A1:.A3])", 60.0),
                f(e, "IFERROR(D!A2,0)", "IFERROR([$D.A2];0)", 20.0),
            ],
        ),
    ]
}

#[test]
fn a_range_that_shrinks_gives_its_new_result_and_iferror_does_not_trap_a_deleted_reference() {
    for extension in EXTENSIONS {
        let dir = tempfile::tempdir().unwrap();
        let input = workbook(dir.path(), extension, &reading_a_range(extension));
        let output = dir.path().join(format!("out.{extension}"));
        let mut book = editor(&input);
        book.delete_rows(["D!2".parse().unwrap()]);
        book.write(&output).unwrap();
        assert!(is_ref_error(&output, "F", 0), "{extension}");
        // The range is now the cells 10 and 30. `IFERROR` does not trap the reference to deleted cells: it
        // is no error value but a reference with no cell, and LibreOffice answers `#REF!` (the oracle of
        // `write_to_libreoffice.rs`, `refs`: `IFERROR(Data!A3,7)` after deleting the row).
        assert!((number(&output, "F", 1) - 40.0).abs() < 1e-9, "{extension}");
        assert!(is_ref_error(&output, "F", 2), "{extension}");
    }
}

/// A formula that mixes a deleted reference with a cell the plan leaves alone, one that reads it,
/// and one that reads positions.
fn mixing_with_a_cell_that_is_left_alone(e: &str) -> Vec<(&'static str, Vec<Cell>)> {
    vec![
        ("D", vec![N(10.0), N(20.0), N(30.0)]),
        (
            "F",
            vec![
                f(e, "D!A1*D!A2", "[$D.A1]*[$D.A2]", 200.0),
                f(e, "A1+1", "[.A1]+1", 201.0),
                f(e, "ROW(D!A1)", "ROW([$D.A1])", 1.0),
            ],
        ),
    ]
}

#[test]
fn a_formula_that_mixes_a_deleted_reference_with_a_number_that_is_left_alone_is_given_the_error() {
    for extension in EXTENSIONS {
        let dir = tempfile::tempdir().unwrap();
        let input = workbook(
            dir.path(),
            extension,
            &mixing_with_a_cell_that_is_left_alone(extension),
        );
        let output = dir.path().join(format!("out.{extension}"));
        let mut book = editor(&input);
        book.delete_rows(["D!2".parse().unwrap()]);
        book.write(&output).unwrap();
        assert!(is_ref_error(&output, "F", 0), "{extension}");
        assert!(is_ref_error(&output, "F", 1), "{extension}");
        assert!((number(&output, "F", 2) - 1.0).abs() < 1e-9, "{extension}");
    }
}
