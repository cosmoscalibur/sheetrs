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
    let model = loader::load(&path).unwrap();
    let spec = Spec {
        fold: false,
        now: Some(46318.5),
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
