//! `sheetrs lint --format`: the executable writes what the library renders.

use sheetrs_io::Book;
use sheetrs_linter::output::render_report;
use sheetrs_linter::{Linter, LinterConfig, OutputFormat};
use std::process::Command;

const XLSX: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/lint_book.xlsx");

fn run(format: OutputFormat) -> (String, usize) {
    let workbook = Book::from_file(XLSX).unwrap().workbook().unwrap();
    let config = LinterConfig::default();
    let violations = Linter::with_config(config.clone()).lint_workbook(&workbook);
    let text = render_report(format, XLSX, &violations, &workbook, &config, false).unwrap();
    (text, violations.len())
}

#[test]
fn the_executable_writes_what_the_library_renders() {
    for (flag, format) in [("json", OutputFormat::Json), ("sarif", OutputFormat::Sarif)] {
        let output = Command::new(env!("CARGO_BIN_EXE_sheetrs"))
            .arg("lint")
            .args(["-i", XLSX, "--format", flag])
            .output()
            .unwrap();
        // The exit code is 1 when there are errors; the document is written either way.
        let (expected, _) = run(format);
        assert_eq!(
            String::from_utf8(output.stdout).unwrap(),
            expected,
            "{flag}"
        );
    }
}
