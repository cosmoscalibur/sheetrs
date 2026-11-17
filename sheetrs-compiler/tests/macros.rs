//! A workbook that carries macros is refused by the build itself, in every entry point that builds, with no policy of
//! the layer above: `Book::compile`, `Book::calculate` (the entry of `sheetc calculate`, `sheetrs edit`, Python and
//! WebAssembly) and the command line. Secure mode opens it. `inspect` and `verify` only read the workbook and never refuse
//! it. The workbooks are those of `sheetrs-edit` (`tests/assets/from_cache/macro.*`, `Calc!A1` calls `MYFUNC`).

use serde_json::json;
use sheetrs_compiler::memory::{BookCompile, Options};
use sheetrs_engine::ir::Mode;
use sheetrs_io::Book;
use std::path::{Path, PathBuf};
use std::process::Command;

const SHEETC: &str = env!("CARGO_BIN_EXE_sheetc");

fn fixture(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-edit/tests/assets/from_cache")
        .join(format!("macro.{extension}"))
}

fn book(extension: &str, secure: bool) -> Book {
    Book::from_file(fixture(extension))
        .unwrap()
        .with_secure_mode(secure)
}

fn refusal<T>(result: sheetrs_compiler::Result<T>) -> String {
    match result {
        Err(sheetrs_compiler::Error::Macros(message)) => message,
        Err(other) => panic!("another error: {other}"),
        Ok(_) => panic!("the workbook was built"),
    }
}

#[test]
fn compile_refuses_a_workbook_with_macros_unless_secure_mode_is_on() {
    for extension in ["xlsx", "ods"] {
        let message = refusal(book(extension, false).compile(&Options::default(), Mode::Values));
        assert!(
            message.contains("macros") && message.contains("secure mode"),
            "{extension}: {message}"
        );
        let (_, report) = book(extension, true)
            .compile(&Options::default(), Mode::Values)
            .unwrap_or_else(|e| panic!("{extension}: {e}"));
        assert!(report.formulas_total > 0, "{extension}");
    }
}

#[test]
fn calculate_refuses_a_workbook_with_macros_unless_secure_mode_is_on() {
    for extension in ["xlsx", "ods"] {
        let request = json!({});
        let message = refusal(book(extension, false).calculate(
            &request,
            &Options::default(),
            Mode::Recalculate,
            false,
        ));
        assert!(message.contains("secure mode"), "{extension}: {message}");
        book(extension, true)
            .calculate(&request, &Options::default(), Mode::Recalculate, false)
            .unwrap_or_else(|e| panic!("{extension}: {e}"));
    }
}

#[test]
fn the_secure_mode_of_the_options_opens_a_workbook_with_macros_too() {
    let mut options = Options::default();
    options.config.secure_mode = Some(true);
    book("xlsx", false).compile(&options, Mode::Values).unwrap();
}

#[test]
fn inspect_and_verify_read_a_workbook_with_macros() {
    for extension in ["xlsx", "ods"] {
        let book = book(extension, false);
        let inspection = book.inspect(&Options::default()).unwrap();
        assert!(inspection.formulas > 0, "{extension}");
        let verified = book
            .verify(&Options::default(), &[], None, None)
            .unwrap_or_else(|e| panic!("{extension}: {e}"));
        assert!(verified["sheets"].is_array(), "{extension}");
    }
}

fn sheetc(args: &[&str], extension: &str) -> (bool, String) {
    let out = Command::new(SHEETC)
        .args(args)
        .arg("-i")
        .arg(fixture(extension))
        .output()
        .unwrap();
    (
        out.status.success(),
        String::from_utf8_lossy(&out.stderr).into_owned(),
    )
}

#[test]
fn sheetc_commands_that_build_refuse_and_the_ones_that_read_do_not() {
    for extension in ["xlsx", "ods"] {
        let (ok, err) = sheetc(&["calculate"], extension);
        assert!(!ok && err.contains("--secure-mode"), "{extension}: {err}");
        let (ok, err) = sheetc(&["calculate", "--secure-mode"], extension);
        assert!(ok, "{extension}: {err}");
        // `verify` exits 1 when the engine differs from the cache (a macro does), and says nothing of a refusal.
        let (ok, err) = sheetc(&["inspect"], extension);
        assert!(ok, "inspect {extension}: {err}");
        let (_, err) = sheetc(&["verify"], extension);
        assert!(!err.contains("secure mode"), "verify {extension}: {err}");
    }
}
