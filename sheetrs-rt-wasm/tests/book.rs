//! The book of a module, natively: the program of a workbook as the payload a module carries, decoded
//! and answered as an executable would. (That the module does the same inside WebAssembly is the
//! test of `sheetrs-wasm`, which instantiates the modules `sheetc` writes.)

use serde_json::json;
use sheetrs_compiler::memory::{BookCompile, Options};
use sheetrs_engine::bundle::{self, Embedded};
use sheetrs_engine::calculate::Overrides;
use sheetrs_engine::ir::Mode;
use sheetrs_rt_wasm::Book;

fn loan() -> sheetrs_io::Book {
    sheetrs_io::Book::from_file(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../sheetrs-compiler/tests/assets/loan.xlsx"
    ))
    .expect("the loan workbook")
}

fn payload(mode: Mode, with_template: bool) -> &'static [u8] {
    let source = loan();
    let (mut program, _) = source.compile(&Options::default(), mode).expect("compiles");
    let bytes = std::fs::read(source.path().unwrap()).unwrap();
    let template = with_template.then_some(("xlsx", bytes.as_slice()));
    // A module keeps its payload for as long as it lives.
    Vec::leak(bundle::payload(&mut program, template).expect("payload"))
}

#[test]
fn the_book_answers_what_the_program_answers() {
    let source = loan();
    let (program, _) = source.compile(&Options::default(), Mode::Values).unwrap();
    let request =
        json!({"inputs": {"principal": 250_000, "rate": 0.045, "years": 15, "month": 24}});
    let expected = program
        .handle(&request, &Overrides::default())
        .unwrap()
        .to_json();
    let book = Book::from_payload(payload(Mode::Values, false)).unwrap();
    let answer = book
        .calculate(request.to_string().as_bytes(), &Overrides::default())
        .unwrap();
    // The text, not a parsed value: parsing a float may differ in the last digit.
    assert_eq!(answer, expected.to_string());
    // An empty request is the declared outputs, as for an executable.
    assert!(
        book.calculate(b"", &Overrides::default())
            .unwrap()
            .contains("payment")
    );
    assert!(book.meta().contains("\"mode\":\"values\""));
    assert!(book.info().contains("\"cells\""));
}

#[test]
fn a_request_the_program_does_not_take_is_an_error_not_a_panic() {
    let book = Book::from_payload(payload(Mode::Values, false)).unwrap();
    let error = book
        .calculate(br#"{"inputs": {"bogus": 1}}"#, &Overrides::default())
        .unwrap_err();
    assert!(error.contains("bogus"), "{error}");
    let error = book
        .calculate(b"{not json", &Overrides::default())
        .unwrap_err();
    assert!(error.contains("not JSON"), "{error}");
}

#[test]
fn a_payload_that_is_not_one_is_refused() {
    assert!(Book::from_payload(b"not a payload at all, not even close").is_err());
    let good = payload(Mode::Values, false);
    // The trailer says how long the program is: cut the payload short and it no longer fits.
    assert!(Book::from_payload(&good[good.len() / 2..]).is_err());
}

#[test]
fn the_payload_is_the_one_an_executable_carries() {
    let embedded: Embedded = bundle::read_payload(payload(Mode::Embedded, true)).unwrap();
    let template = embedded.template.expect("the workbook travels");
    assert_eq!(template.extension, "xlsx");
    assert!(template.bytes.starts_with(b"PK"));
    assert_eq!(embedded.program.mode, Mode::Embedded);
}

#[cfg(feature = "template")]
#[test]
fn a_book_with_its_workbook_writes_the_answer_into_it() {
    let book = Book::from_payload(payload(Mode::Embedded, true)).unwrap();
    assert!(book.has_template());
    let request = json!({"inputs": {"principal": 250_000}}).to_string();
    let (workbook, json) = book
        .write(request.as_bytes(), true, &Overrides::default())
        .unwrap();
    assert!(workbook.starts_with(b"PK"));
    assert!(json.unwrap().contains("payment"));
    let (_, json) = book
        .write(request.as_bytes(), false, &Overrides::default())
        .unwrap();
    assert!(json.is_none());
    // Without its workbook a book cannot write one.
    let plain = Book::from_payload(payload(Mode::Values, false)).unwrap();
    assert!(!plain.has_template());
    assert!(plain.write(b"", false, &Overrides::default()).is_err());
}
