//! Cost of text comparison on a large workbook: `cargo run --release -p sheetrs-compiler
//! --example bench_text -- FILE` (the file of the `timing` example: `Data!I` holds
//! 600 000 labels, `Data!K` formulas).
use sheetrs_compiler::{BindingSpec, Model, Overrides, Spec, compile, loader};
use std::time::Instant;

fn main() {
    let file = std::env::args().nth(1).expect("file");
    let wb = sheetrs_core::reader::read_workbook(&file).unwrap();
    let mut model: Model = loader::from_workbook_owned(wb, false);
    let s = model.sheet_id("Summary").unwrap();
    let cases = [
        (
            "MATCH, text not found (600k scans)",
            "MATCH(\"zzz\",Data!I1:I600000,0)",
        ),
        (
            "VLOOKUP exact text, found at row 6",
            "VLOOKUP(\"ñandú\",Data!I1:K600000,3,FALSE)",
        ),
        (
            "SUMPRODUCT(--(range=\"beta\")) (600k ==)",
            "SUMPRODUCT(--(Data!I1:I600000=\"beta\"))",
        ),
        (
            "COUNTIF(range,\"<gamma\") (600k order)",
            "COUNTIF(Data!I1:I600000,\"<gamma\")",
        ),
        (
            "COUNTIF(range,\"alpha\") (pattern matcher)",
            "COUNTIF(Data!I1:I600000,\"alpha\")",
        ),
        ("SUM over 600k numbers (no text)", "SUM(Data!K1:K600000)"),
    ];
    for (i, (_, f)) in cases.iter().enumerate() {
        model.set_formula(s, &format!("D{}", i + 1), f);
    }
    // An input keeps everything live, so the run does the work.
    let spec = Spec {
        inputs: vec![BindingSpec::parse("Data!A1")],
        outputs: vec![BindingSpec::parse(&format!("Summary!D1:D{}", cases.len()))],
        fold: false,
        ..Spec::default()
    };
    let t = Instant::now();
    let (p, _) = compile(&model, &spec).unwrap();
    println!("compile (no fold) {:>9.1?}", t.elapsed());
    // Evaluate each formula alone to time it.
    for (i, (label, _)) in cases.iter().enumerate() {
        let spec = Spec {
            inputs: vec![BindingSpec::parse("Data!A1")],
            outputs: vec![BindingSpec::parse(&format!("Summary!D{}", i + 1))],
            fold: false,
            ..Spec::default()
        };
        let (p, _) = compile(&model, &spec).unwrap();
        let t = Instant::now();
        let out = p
            .execute(&serde_json::Value::Null, &Overrides::default())
            .unwrap();
        println!(
            "{label:<44} {:>9.1?}  -> {}",
            t.elapsed(),
            out.as_object()
                .and_then(|o| o.values().next())
                .map(ToString::to_string)
                .unwrap_or_default()
        );
    }
    let _ = p;
}
