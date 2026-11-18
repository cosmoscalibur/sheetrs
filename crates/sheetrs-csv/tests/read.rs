use std::error::Error as StdError;

use sheetrs_core::{CellPos, MAX_COLS, Value};
use sheetrs_csv::{FormulaFields, Options, read};

type TestResult = Result<(), Box<dyn StdError>>;

fn cell(
    text: &str,
    options: &Options,
    row: u32,
    col: u32,
) -> Result<Option<Value>, Box<dyn StdError>> {
    let workbook = read(text.as_bytes(), "data", options)?;
    Ok(workbook
        .sheet("data")
        .and_then(|sheet| sheet.get(CellPos { row, col }))
        .map(|view| view.to_cell().value))
}

#[test]
fn the_sheet_has_the_name_the_caller_gives() -> TestResult {
    let workbook = read("a\n".as_bytes(), "Prices", &Options::comma_separated())?;
    assert_eq!(
        workbook
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["Prices"]
    );
    Ok(())
}

#[test]
fn numbers_and_text_are_told_apart() -> TestResult {
    let options = Options::comma_separated();
    assert_eq!(cell("1.5,abc\n", &options, 0, 0)?, Some(Value::Number(1.5)));
    assert_eq!(
        cell("1.5,abc\n", &options, 0, 1)?,
        Some(Value::Text("abc".into()))
    );
    assert_eq!(cell("1.5,,x\n", &options, 0, 1)?, None);
    Ok(())
}

#[test]
fn a_formula_field_is_a_formula_only_when_the_option_says_so() -> TestResult {
    let text = "=1+1\n";
    assert_eq!(
        cell(text, &Options::comma_separated(), 0, 0)?,
        Some(Value::Text("=1+1".into()))
    );
    let options = Options {
        formula_fields: FormulaFields::Evaluated,
        ..Options::comma_separated()
    };
    let workbook = read(text.as_bytes(), "data", &options)?;
    let cell = workbook
        .sheet("data")
        .and_then(|sheet| sheet.get(CellPos { row: 0, col: 0 }))
        .map(|view| view.to_cell())
        .ok_or("cell")?;
    assert_eq!(
        cell.formula,
        Some(sheetrs_core::Formula::Text("=1+1".to_owned()))
    );
    assert!(cell.stale);
    Ok(())
}

#[test]
fn fields_past_the_last_column_are_dropped() -> TestResult {
    let row = (0..=MAX_COLS)
        .map(|number| number.to_string())
        .collect::<Vec<_>>()
        .join(",");
    let workbook = read(row.as_bytes(), "data", &Options::comma_separated())?;
    assert_eq!(
        workbook.sheet("data").map(|sheet| sheet.len()),
        Some(usize::try_from(MAX_COLS)?)
    );
    Ok(())
}

#[test]
fn input_that_cannot_be_read_is_reported() {
    struct Broken;
    impl std::io::Read for Broken {
        fn read(&mut self, _: &mut [u8]) -> std::io::Result<usize> {
            Err(std::io::Error::other("broken"))
        }
    }
    assert!(matches!(
        read(Broken, "data", &Options::comma_separated()),
        Err(sheetrs_csv::Error::Io(_))
    ));
}

#[test]
fn many_quotes_that_never_close_are_read_in_linear_time() -> TestResult {
    // Each line opens a quote that no later quote closes; re-reading the rest of the text for
    // each of them would take minutes.
    let text = "\"x\n".repeat(300_000);
    let workbook = read(text.as_bytes(), "data", &Options::comma_separated())?;
    assert_eq!(
        workbook.sheet("data").map(|sheet| sheet.len()),
        Some(300_000)
    );
    Ok(())
}
