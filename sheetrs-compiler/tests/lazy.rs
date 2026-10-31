//! An executable of a big program keeps its constants apart and decodes the blocks a request
//! reads: the answers must be those of the program held whole.

use serde_json::{Value as Json, json};
use sheetrs_compiler::compile::QueryLevel;
use sheetrs_compiler::{BindingSpec, Model, Overrides, Spec, Value, compile};
use std::io::Write;
use std::process::{Command, Stdio};

const SHEETC: &str = env!("CARGO_BIN_EXE_sheetc");

fn big_model(rows: u32) -> Model {
    let mut m = Model::new();
    let s = m.add_sheet("Data");
    for i in 1..=rows {
        m.set(s, &format!("A{i}"), Value::Number(f64::from(i)));
        m.set(s, &format!("B{i}"), Value::text(format!("row {i}")));
        m.set(s, &format!("C{i}"), Value::Bool(i % 2 == 0));
    }
    // Input-dependent formulas stay in the program; the rest is folded into constants.
    m.set(s, "E1", Value::Number(2.0));
    m.set_formula(s, "F1", "=E1*A1");
    m.set_formula(s, "F2", &format!("=SUM(A1:A{rows})+E1"));
    m.set_formula(s, "G1", &format!("=SUM(A1:A{rows})"));
    m.set_formula(s, "G2", "=VLOOKUP(\"row 20000\",B1:B20000,1,FALSE)");
    m.set_formula(s, "G3", "=INDEX(B:B,12345)");
    m
}

fn ask(exe: &std::path::Path, request: &Json) -> Json {
    let mut child = Command::new(exe)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    child
        .stdin
        .take()
        .unwrap()
        .write_all(request.to_string().as_bytes())
        .unwrap();
    let out = child.wait_with_output().unwrap();
    assert!(
        out.status.success(),
        "{}",
        String::from_utf8_lossy(&out.stderr)
    );
    serde_json::from_slice(&out.stdout).unwrap()
}

#[test]
fn constants_kept_apart_answer_as_the_whole_program() {
    let m = big_model(20_000);
    let spec = Spec {
        inputs: vec![BindingSpec::named("factor", "Data!E1")],
        query: QueryLevel::Any,
        fold: true,
        ..Spec::default()
    };
    let (mut program, _) = compile(&m, &spec).unwrap();
    let whole = program.clone();
    let cells_before = program.cells.len();

    let dir = tempfile::tempdir().unwrap();
    let exe = dir.path().join("big");
    sheetrs_engine::bundle::write_executable(&mut program, std::path::Path::new(SHEETC), &exe)
        .unwrap();
    assert!(
        program.cells.len() < cells_before / 10,
        "the constants left the program: {} of {cells_before} cells remain",
        program.cells.len()
    );

    let requests = [
        json!({"ranges": ["Data!A1:C3", "Data!F1:G3"]}),
        json!({"ranges": ["Data!A19998:C20000"], "majorDimension": "COLUMNS"}),
        json!({"inputs": {"factor": 5}, "ranges": ["Data!F1:F2", "Data!A4096:B4097"]}),
        // Open-ended: the extent comes from the constants.
        json!({"ranges": ["Data!A:A"]}),
        json!({"ranges": ["Data!E1"], "inputs": {"factor": {"error": "#N/A"}}}),
    ];
    for r in &requests {
        let expected = whole.handle(r, &Overrides::default()).unwrap();
        assert_eq!(ask(&exe, r), expected, "{r}");
    }
    let (code, info) = {
        let out = Command::new(&exe).arg("--info").output().unwrap();
        (out.status.code(), String::from_utf8(out.stdout).unwrap())
    };
    assert_eq!(code, Some(0));
    let info: Json = serde_json::from_str(&info).unwrap();
    assert_eq!(info["cells"], whole.cells.len());
}
