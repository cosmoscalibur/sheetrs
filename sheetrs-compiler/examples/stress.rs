//! Scale check: `cargo run --release --example stress`.
use serde_json::json;
use sheetrs_compiler::compile::QueryLevel;
use sheetrs_compiler::{BindingSpec, Model, Overrides, Spec, Value, compile};
use std::time::Instant;

fn time<T>(label: &str, f: impl FnOnce() -> T) -> T {
    let t = Instant::now();
    let r = f();
    println!("  {label:<28} {:>9.1?}", t.elapsed());
    r
}

fn run(title: &str, m: &Model, inputs: &[&str], outputs: &[&str], input: serde_json::Value) {
    run_level(title, m, inputs, outputs, input, QueryLevel::Declared);
}

fn run_level(
    title: &str,
    m: &Model,
    inputs: &[&str],
    outputs: &[&str],
    input: serde_json::Value,
    query: QueryLevel,
) {
    println!("{title}");
    let spec = Spec {
        inputs: inputs.iter().map(|s| BindingSpec::parse(s)).collect(),
        outputs: outputs.iter().map(|s| BindingSpec::parse(s)).collect(),
        query,
        ..Spec::default()
    };
    let (p, r) = time("compile", || compile(m, &spec).unwrap());
    println!(
        "  kept {} cells, {} formulas (of {} reachable)",
        r.cells_kept, r.formulas_kept, r.cells_reachable
    );
    let out = time("run", || p.execute(&input, &Overrides::default()).unwrap());
    println!(
        "  -> {}",
        out.to_string().chars().take(100).collect::<String>()
    );
}

fn main() {
    let n = 100_000u32;
    // A: a long dependency chain from one input.
    let mut m = Model::new();
    let s = m.add_sheet("S");
    m.set(s, "A1", Value::Number(1.0));
    for i in 2..=n {
        m.set_formula(s, &format!("A{i}"), &format!("A{}+1", i - 1));
    }
    run(
        "chain of 100k formulas",
        &m,
        &["A1"],
        &[&format!("A{n}")],
        json!({"S!A1": 5}),
    );

    // B: running totals anchored at the top (quadratic range reads).
    let k = 5_000u32;
    let mut m = Model::new();
    let s = m.add_sheet("S");
    for i in 1..=k {
        m.set(s, &format!("A{i}"), Value::Number(f64::from(i)));
        m.set_formula(s, &format!("B{i}"), &format!("SUM($A$1:A{i})"));
    }
    run(
        "running SUM($A$1:An), 5k rows",
        &m,
        &["A1"],
        &[&format!("B1:B{k}")],
        json!({"S!A1": 1}),
    );

    // C: many lookups into one table.
    let t = 10_000u32;
    let mut m = Model::new();
    let s = m.add_sheet("S");
    for i in 1..=t {
        m.set(s, &format!("A{i}"), Value::Number(f64::from(i)));
        m.set(s, &format!("B{i}"), Value::Number(f64::from(i) * 2.0));
    }
    m.set(s, "D1", Value::Number(7.0));
    for i in 1..=20_000u32 {
        m.set_formula(
            s,
            &format!("E{i}"),
            &format!("VLOOKUP(D$1+{i},A$1:B${t},2,FALSE)"),
        );
    }
    m.set_formula(s, "F1", "SUM(E1:E20000)");
    run(
        "20k VLOOKUPs over a 10k table",
        &m,
        &["D1"],
        &["F1"],
        json!({"S!D1": 3}),
    );

    // D: wide independent formulas, all folded at build time.
    let mut m = Model::new();
    let s = m.add_sheet("S");
    for i in 1..=50_000u32 {
        m.set(s, &format!("A{i}"), Value::Number(f64::from(i)));
        m.set_formula(s, &format!("B{i}"), &format!("A{i}*2+ROUND(A{i}/3,2)"));
    }
    m.set_formula(s, "C1", "SUM(B1:B50000)");
    run("50k folded formulas", &m, &[], &["C1"], json!(null));
    // Inputs that may be any cell leave nothing to fold: the cost of `--no-fold`.
    let spec = Spec {
        outputs: vec![BindingSpec::parse("C1")],
        fold: false,
        ..Spec::default()
    };
    let (p, r) = time("compile, no folding", || compile(&m, &spec).unwrap());
    println!(
        "  kept {} cells, {} formulas",
        r.cells_kept, r.formulas_kept
    );
    time("run, no folding", || {
        p.execute(&json!(null), &Overrides::default()).unwrap()
    });
    run_level(
        "50k folded formulas, any range queryable",
        &m,
        &[],
        &["C1"],
        json!(null),
        QueryLevel::Any,
    );
}
