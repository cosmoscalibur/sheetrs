//! The mark to calculate every formula when the workbook is opened (`fullCalcOnLoad` of an XLSX), in the build, the
//! calculation and the edit.
//!
//! A workbook that carries it caches nothing to trust, so a build (`sheetc build`) and a calculation (`sheetc calculate`,
//! `BookEdit::calculate`) take it as a run with `calculate_all`, and so does the edit, which is a calculation and a writer.
//! What leaves with a result the engine calculated has no mark. An ODS has no such mark (`LibreOffice` writes none), so
//! it keeps what it caches, as an unmarked XLSX does.
//!
//! `cached/book.*` caches `Calc!A6` = `rate*2` as 140 where `rate` (`Data!C1`) is 7 and the formula is 14: the value says
//! whether the cache was read or the formula calculated.

use serde_json::json;
use sheetrs_compiler::memory::{BookCompile, Options};
use sheetrs_edit::{BookEdit, Calculation};
use sheetrs_engine::ir::Mode;
use sheetrs_io::Book;
use sheetrs_query::{BookQuery, RangeOptions};
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};

fn book_file(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/cached")
        .join(format!("book.{extension}"))
}

/// The XLSX with `fullCalcOnLoad` on, as an application that leaves stale results writes it.
fn marked(bytes: &[u8]) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..archive.len() {
        let mut entry = archive.by_index(i).unwrap();
        let name = entry.name().unwrap().into_owned();
        let mut data = Vec::new();
        entry.read_to_end(&mut data).unwrap();
        if name == "xl/workbook.xml" {
            let text = String::from_utf8(data).unwrap();
            assert!(text.contains(r#"<calcPr calcId="124519"/>"#));
            data = text
                .replace(
                    r#"<calcPr calcId="124519"/>"#,
                    r#"<calcPr calcId="124519" fullCalcOnLoad="1"/>"#,
                )
                .into_bytes();
        }
        out.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        out.write_all(&data).unwrap();
    }
    out.finish().unwrap().into_inner()
}

fn part(bytes: &[u8], name: &str) -> String {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut text = String::new();
    archive
        .by_name(name)
        .unwrap()
        .read_to_string(&mut text)
        .unwrap();
    text
}

fn xlsx(is_marked: bool) -> Book {
    let bytes = std::fs::read(book_file("xlsx")).unwrap();
    Book::from_bytes(if is_marked { marked(&bytes) } else { bytes }).unwrap()
}

/// The value of `Calc!A6` in the workbook `bytes`, read as the file holds it.
fn a6(bytes: Vec<u8>) -> f64 {
    let blocks = Book::from_bytes(bytes)
        .unwrap()
        .range(["Calc!A6"], &RangeOptions::default())
        .unwrap();
    blocks[0].values.as_ref().unwrap()[0][0]
        .as_ref()
        .and_then(serde_json::Value::as_f64)
        .unwrap()
}

#[test]
fn a_marked_xlsx_is_built_as_a_run_with_calculate_all() {
    let build = |book: &Book| {
        book.compile(&Options::default(), Mode::Values)
            .unwrap()
            .0
            .build
            .options
            .calculate_all
    };
    assert!(xlsx(true).recalculation_marked().unwrap());
    assert!(build(&xlsx(true)));
    assert!(!xlsx(false).recalculation_marked().unwrap());
    assert!(!build(&xlsx(false)));
}

#[test]
fn an_ods_has_no_mark_and_keeps_what_it_caches() {
    let ods = Book::from_file(book_file("ods")).unwrap();
    assert!(!ods.recalculation_marked().unwrap());
    let program = ods.compile(&Options::default(), Mode::Values).unwrap().0;
    assert!(!program.build.options.calculate_all);
}

#[test]
fn a_calculation_of_a_marked_xlsx_calculates_every_formula_and_leaves_no_mark() {
    for (is_marked, expected) in [(true, 14.0), (false, 140.0)] {
        let mut book = xlsx(is_marked);
        BookEdit::calculate(
            &mut book,
            Calculation {
                inputs: json!({"Data!A1": 1}).as_object().cloned().unwrap(),
                ..Calculation::default()
            },
        );
        let mut bytes = Vec::new();
        book.write_to(&mut bytes).unwrap();
        // The marked workbook caches nothing to trust: `rate*2` is calculated. The unmarked one keeps the cache of the
        // formula the input does not reach.
        assert!(
            (a6(bytes.clone()) - expected).abs() < 1e-9,
            "marked: {is_marked}"
        );
        assert!(
            !part(&bytes, "xl/workbook.xml").contains("fullCalcOnLoad"),
            "marked: {is_marked}: the workbook is calculated, so it leaves with no mark"
        );
    }
}

#[test]
fn an_edit_of_a_marked_xlsx_inherits_the_calculation_and_leaves_no_mark() {
    let mut book = xlsx(true);
    book.hide_rows(["Data!2".parse().unwrap()]);
    let mut bytes = Vec::new();
    book.write_to(&mut bytes).unwrap();
    assert!((a6(bytes.clone()) - 14.0).abs() < 1e-9);
    assert!(!part(&bytes, "xl/workbook.xml").contains("fullCalcOnLoad"));
}
