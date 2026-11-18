use std::error::Error as StdError;
use std::io::{Cursor, Write};

use zip::CompressionMethod;
use zip::write::{SimpleFileOptions, ZipWriter};

use sheetrs_core::{CellPos, DateSystem, Value, Visibility};
use sheetrs_ods::{Error, read};

mod common;
use common::{content_xml, document, package_with};

type TestResult = Result<(), Box<dyn StdError>>;

/// A package that has a `mimetype` and nothing else.
fn package_without_content() -> Result<Cursor<Vec<u8>>, Box<dyn StdError>> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    writer.start_file("mimetype", options)?;
    writer.write_all(b"application/vnd.oasis.opendocument.spreadsheet")?;
    Ok(writer.finish()?)
}

fn at(row: u32, col: u32) -> CellPos {
    CellPos { row, col }
}

#[test]
fn tables_keep_order_names_and_visibility() -> TestResult {
    let content = document(
        r#"<style:style style:name="ta1" style:family="table"><style:table-properties table:display="true"/></style:style><style:style style:name="ta2" style:family="table"><style:table-properties table:display="false"/></style:style>"#,
        r#"<table:table table:name="First" table:style-name="ta1"/><table:table table:name="Second" table:style-name="ta2"/><table:table table:name="Third"/>"#,
    );
    let workbook = read(package_with(&content, &[])?)?;
    let sheets: Vec<_> = workbook
        .sheets()
        .map(|sheet| (sheet.name().to_owned(), sheet.visibility()))
        .collect();
    assert_eq!(
        sheets,
        [
            ("First".to_owned(), Visibility::Visible),
            ("Second".to_owned(), Visibility::Hidden),
            ("Third".to_owned(), Visibility::Visible),
        ]
    );
    Ok(())
}

#[test]
fn the_null_date_picks_the_date_system() -> TestResult {
    for (null, expected) in [
        ("1899-12-30", DateSystem::Date1900),
        ("1904-01-01", DateSystem::Date1904),
        ("1970-01-01", DateSystem::Custom { zero: 0 }),
        ("not a date", DateSystem::Date1900),
    ] {
        let settings = format!(
            r#"<table:calculation-settings><table:null-date table:date-value="{null}"/></table:calculation-settings>"#
        );
        let workbook = read(package_with(&content_xml("", &settings), &[])?)?;
        assert_eq!(workbook.date_system(), expected, "{null}");
    }
    Ok(())
}

#[test]
fn a_cell_with_a_formula_keeps_the_text_of_its_attribute() -> TestResult {
    let rows = r#"<table:table-row><table:table-cell table:formula="of:=SUM([.B1];2)" office:value-type="float" office:value="5"/></table:table-row>"#;
    let workbook = read(package_with(&content_xml(rows, ""), &[])?)?;
    let cell = workbook
        .sheet("S1")
        .and_then(|sheet| sheet.get(at(0, 0)))
        .map(|view| view.to_cell())
        .ok_or("cell")?;
    assert_eq!(cell.value, Value::Number(5.0));
    assert_eq!(
        cell.formula,
        Some(sheetrs_core::Formula::Text("of:=SUM([.B1];2)".to_owned()))
    );
    assert!(cell.stale);
    Ok(())
}

#[test]
fn rows_repeated_a_million_times_stop_at_the_last_row_of_the_sheet() -> TestResult {
    let rows = r#"<table:table-row table:number-rows-repeated="2000000"><table:table-cell office:value-type="float" office:value="1"/></table:table-row>"#;
    let workbook = read(package_with(&content_xml(rows, ""), &[])?)?;
    assert_eq!(
        workbook.sheet("S1").map(|sheet| sheet.len()),
        Some(usize::try_from(sheetrs_core::MAX_ROWS)?)
    );
    Ok(())
}

#[test]
fn a_package_without_content_is_reported() {
    let package = package_without_content().expect("a package");
    assert!(matches!(read(package), Err(Error::MissingPart(part)) if part == "content.xml"));
}

#[test]
fn input_that_is_not_a_zip_package_is_rejected() {
    assert!(matches!(
        read(Cursor::new(b"not a package".to_vec())),
        Err(Error::Zip(_))
    ));
}

fn paragraph_cell(inner: &str) -> String {
    format!(
        r#"<table:table-row><table:table-cell office:value-type="string"><text:p>{inner}</text:p></table:table-cell></table:table-row>"#
    )
}

#[test]
fn elements_nested_without_limit_in_a_paragraph_are_refused() -> TestResult {
    let nested = "<text:x>".repeat(100_000);
    let rows = paragraph_cell(&nested);
    let outcome = read(package_with(&content_xml(&rows, ""), &[])?);
    // The document is not well-formed either, but the nesting is what stops the reader first.
    assert!(matches!(outcome, Err(Error::Malformed { .. })));
    Ok(())
}

#[test]
fn a_paragraph_that_grows_through_spaces_without_limit_is_refused() -> TestResult {
    let spaces = r#"<text:s text:c="16000000"/>"#.repeat(3);
    let outcome = read(package_with(
        &content_xml(&paragraph_cell(&spaces), ""),
        &[],
    )?);
    assert!(matches!(outcome, Err(Error::Malformed { .. })));
    Ok(())
}

#[test]
fn a_table_that_repeats_to_more_cells_than_the_reader_writes_is_refused() -> TestResult {
    let rows = r#"<table:table-row table:number-rows-repeated="1048576"><table:table-cell office:value-type="float" office:value="1" table:number-columns-repeated="100"/></table:table-row>"#;
    let outcome = read(package_with(&content_xml(rows, ""), &[])?);
    assert!(matches!(outcome, Err(Error::TooManyCells(_))));
    Ok(())
}

#[test]
fn a_document_that_is_cut_short_is_refused() -> TestResult {
    let content = content_xml(&paragraph_cell("a"), "");
    let cut = content.replace("</table:table>", "");
    assert!(matches!(
        read(package_with(&cut, &[])?),
        Err(Error::Malformed { .. } | Error::Xml { .. })
    ));
    Ok(())
}
