//! ODS/XLSX parity: the same workbook saved in both formats must compute the
//! same results through the same engine.

use serde_json::Value as Json;
use sheetrs_compiler::{Overrides, Spec, compile, loader};
use std::path::Path;

fn results(file: &str) -> Json {
    // minimal_test.* are the suite-wide fixtures at the workspace root.
    let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
    let path = if file.starts_with("minimal_test") {
        manifest.join("../tests").join(file)
    } else {
        manifest.join("tests/assets").join(file)
    };
    let mut model = loader::load(&path).unwrap();
    // LibreOffice reads a bare `#REF!` of an XLSX formula as an unknown name; that import
    // quirk is not what this test compares (the readers), so both are read alike.
    model.from_xlsx = false;
    let spec = Spec {
        fold: false,
        now: Some(46318.5),
        // The test workbook holds an external reference, which a build refuses: compare the rest.
        inspect: true,
        ..Spec::default()
    };
    let (program, _) = compile(&model, &spec).unwrap();
    program.execute(&Json::Null, &Overrides::default()).unwrap()
}

#[test]
fn minimal_workbook_computes_identically() {
    let (x, o) = (results("minimal_test.xlsx"), results("minimal_test.ods"));
    assert!(
        x.as_object().unwrap().len() >= 3,
        "expected several formula sheets"
    );
    assert_eq!(x, o);
}

#[test]
fn loan_workbook_computes_identically() {
    let (x, o) = (results("loan.xlsx"), results("loan.ods"));
    assert_eq!(x, o);
}

/// The compiler reads a workbook straight into its model (`loader::load`); this must give exactly
/// what reading the workbook first and converting it gives (`read_workbook` + `from_workbook_owned`,
/// the path the linter's reader still takes), for every workbook at hand.
#[test]
fn reading_straight_into_the_model_matches_reading_the_workbook() {
    let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
    let mut files: Vec<std::path::PathBuf> = Vec::new();
    let mut dirs = vec![manifest.join("tests/assets"), manifest.join("../tests")];
    // The large workbooks of `tests/production` (not in the repository) are read too when asked:
    // a debug build takes minutes on them (`SHEETC_PARITY_PRODUCTION=1 cargo test --release`).
    if std::env::var_os("SHEETC_PARITY_PRODUCTION").is_some() {
        dirs.push(manifest.join("../tests/production"));
    }
    for dir in dirs {
        let Ok(entries) = std::fs::read_dir(dir) else {
            continue;
        };
        files.extend(entries.flatten().map(|e| e.path()).filter(|p| {
            p.extension()
                .and_then(|e| e.to_str())
                .is_some_and(|e| matches!(e, "xlsx" | "ods"))
        }));
    }
    files.sort();
    assert!(files.len() >= 8, "fixtures missing: {}", files.len());
    for path in files {
        // A workbook that cannot be read (a macro file, a damaged one) fails the same way twice.
        let wb = sheetrs_core::reader::read_workbook(&path);
        let direct = loader::load(&path);
        let (Ok(wb), Ok(direct)) = (wb, direct) else {
            continue;
        };
        let ods = path.extension().is_some_and(|e| e == "ods");
        let mut via = loader::from_workbook_owned(wb, ods);
        via.from_xlsx = !ods;
        let name = path.display();
        assert_eq!(direct.sheets, via.sheets, "{name}: sheets");
        assert_eq!(direct.names, via.names, "{name}: names");
        assert_eq!(direct.date_system, via.date_system, "{name}: date system");
        assert_eq!(direct.settings, via.settings, "{name}: settings");
        assert_eq!(direct.hidden_rows, via.hidden_rows, "{name}: hidden rows");
        assert_eq!(direct.macros, via.macros, "{name}: macros");
        assert_eq!(direct.diagnostics, via.diagnostics, "{name}: diagnostics");
        assert_eq!(direct.from_xlsx, via.from_xlsx, "{name}: format");
        assert_eq!(direct.cells.len(), via.cells.len(), "{name}: cell count");
        for (at, cell) in via.cells.iter() {
            let got = direct.cells.get(&at);
            assert!(
                got == Some(cell),
                "{name}: cell {at:?}\n via workbook: {cell:?}\n straight:     {got:?}"
            );
        }
    }
}
