use crate::common::*;
use serde_json::{Value as Json, json};
use sheetrs_compiler::BindingSpec;
use sheetrs_compiler::ir::CellKind;
use sheetrs_compiler::options::Iteration;
use sheetrs_compiler::{Overrides, Spec};

fn model() -> sheetrs_compiler::Model {
    sheet(&[
        ("A1", json!(100)),             // input: price
        ("A2", json!(3)),               // input: qty
        ("B1", json!(0.21)),            // constant: tax rate
        ("B2", json!("=B1*2")),         // folded: depends on constants only
        ("C1", json!("=A1*A2")),        // intermediate (depends on inputs)
        ("C2", json!("=C1*(1+B1)+B2")), // final output
        ("D1", json!("=SUM(1,2,3)")),   // unreachable from the output
        ("D2", json!("=D1*1000")),      // unreachable too
    ])
}

#[test]
fn inputs_outputs_prune_and_fold() {
    let (p, r) = build(&model(), &spec(&["A1", "A2"], &["C2"]));
    // D1/D2 are dead; B2 is folded to a constant; C1 and C2 stay.
    assert_eq!(r.formulas_total, 5);
    assert_eq!(r.formulas_kept, 2);
    assert!(r.formulas_folded >= 1);
    let formulas = p
        .cells
        .iter()
        .filter(|c| matches!(c.kind, CellKind::Formula(_)))
        .count();
    assert_eq!(formulas, 2);
    assert!(
        p.cells.iter().all(|c| !(c.at.col == 3)),
        "column D must be pruned"
    );

    let out = p
        .execute(
            &json!({"Sheet1!A1": 100, "Sheet1!A2": 3}),
            &Overrides::default(),
        )
        .unwrap();
    assert_eq!(out["Sheet1!C2"], json!(300.0 * 1.21 + 0.42));
    // Defaults come from the source file when an input is omitted.
    let out = p
        .execute(&json!({"Sheet1!A2": 1}), &Overrides::default())
        .unwrap();
    assert_eq!(out["Sheet1!C2"], json!(100.0 * 1.21 + 0.42));
}

#[test]
fn unknown_and_malformed_inputs_are_rejected() {
    let (p, _) = build(&model(), &spec(&["A1"], &["C2"]));
    let e = p
        .execute(&json!({"nope": 1}), &Overrides::default())
        .unwrap_err();
    assert!(
        e.0.contains("unknown input `nope`") && e.0.contains("Sheet1!A1"),
        "{e}"
    );
    assert!(p.execute(&json!([1]), &Overrides::default()).is_err());
}

#[test]
fn everything_folds_without_inputs() {
    let (p, r) = build(&model(), &spec(&[], &["C2", "D2"]));
    assert_eq!(r.formulas_kept, 0);
    assert!(p.cells.iter().all(|c| matches!(c.kind, CellKind::Const(_))));
    let out = p.execute(&Json::Null, &Overrides::default()).unwrap();
    assert_eq!(out["Sheet1!D2"], json!(6000));
}

#[test]
fn range_bindings_use_json_matrices() {
    let m = sheet(&[
        ("A1", json!(0)),
        ("A2", json!(0)),
        ("A3", json!(0)),
        ("B1", json!("=A1*2")),
        ("B2", json!("=A2*2")),
        ("B3", json!("=A3*2")),
        ("C1", json!("=SUM(B1:B3)")),
    ]);
    let (p, _) = build(&m, &spec(&["A1:A3"], &["B1:B3", "C1"]));
    for input in [json!([1, 2, 3]), json!([[1], [2], [3]])] {
        let out = p
            .execute(&json!({ "Sheet1!A1:A3": input }), &Overrides::default())
            .unwrap();
        assert_eq!(out["Sheet1!B1:B3"], json!([[2], [4], [6]]));
        assert_eq!(out["Sheet1!C1"], json!(12));
    }
    let e = p
        .execute(&json!({"Sheet1!A1:A3": [1, 2]}), &Overrides::default())
        .unwrap_err();
    assert!(e.0.contains("expects 3"), "{e}");
}

#[test]
fn errors_and_blanks_in_json() {
    let m = sheet(&[
        ("A1", json!(0)),
        ("B1", json!("=A1/0")),
        ("C1", json!("=A2")),
    ]);
    let (p, _) = build(&m, &spec(&["A1"], &["B1", "C1"]));
    let out = p
        .execute(&json!({"Sheet1!A1": "hi"}), &Overrides::default())
        .unwrap();
    assert_eq!(out["Sheet1!B1"], err("#VALUE!"));
    // A formula that refers to a blank cell holds 0, as in Excel and LibreOffice.
    assert_eq!(out["Sheet1!C1"], json!(0));
}

#[test]
fn default_outputs_cover_formula_ranges() {
    let (p, _) = build(&model(), &spec(&[], &[]));
    assert_eq!(p.outputs.len(), 1);
    assert_eq!(p.outputs[0].name, "Sheet1!B1:D2");
    let out = p.execute(&Json::Null, &Overrides::default()).unwrap();
    assert_eq!(out["Sheet1!B1:D2"][1][1], json!(300.0 * 1.21 + 0.42)); // C2 in B1:D2
}

#[test]
fn circular_references_and_iteration() {
    let m = sheet(&[
        ("A1", json!("=B1+1")),
        ("B1", json!("=A1+1")),
        ("C1", json!("=A1")),
    ]);
    let (p, r) = build(&m, &spec(&[], &["A1", "C1"]));
    assert_eq!(r.cycles, 1);
    let out = p.execute(&Json::Null, &Overrides::default()).unwrap();
    assert_eq!(out["Sheet1!A1"], err("#CIRC!"));
    assert_eq!(out["Sheet1!C1"], err("#CIRC!"));

    // x = y/2 + 10, y = x  →  x = y = 20.
    let m = sheet(&[
        ("A1", json!("=B1/2+10")),
        ("B1", json!("=A1")),
        ("C1", json!("=A1+B1")),
    ]);
    let mut s = spec(&[], &["A1", "C1"]);
    s.iteration = Some(Iteration {
        max_iterations: 200,
        max_change: 1e-12,
    });
    let (p, _) = build(&m, &s);
    let out = p.execute(&Json::Null, &Overrides::default()).unwrap();
    assert!((out["Sheet1!A1"].as_f64().unwrap() - 20.0).abs() < 1e-9);
    assert!((out["Sheet1!C1"].as_f64().unwrap() - 40.0).abs() < 1e-9);
}

#[test]
fn cycles_through_inputs_survive_to_run_time() {
    // The cycle depends on an input, so it cannot be folded away.
    let m = sheet(&[
        ("A1", json!(1)),
        ("B1", json!("=C1*0.5+A1")),
        ("C1", json!("=B1")),
    ]);
    let mut s = spec(&["A1"], &["B1"]);
    s.iteration = Some(Iteration {
        max_iterations: 500,
        max_change: 1e-12,
    });
    let (p, r) = build(&m, &s);
    assert_eq!(r.formulas_kept, 2);
    assert_eq!(p.cycles.len(), 1);
    let out = p
        .execute(&json!({"Sheet1!A1": 5}), &Overrides::default())
        .unwrap();
    assert!((out["Sheet1!B1"].as_f64().unwrap() - 10.0).abs() < 1e-9);
}

#[test]
fn dynamic_references_resolve_at_run_time() {
    let m = sheet(&[
        ("A1", json!(2)), // input: which row to read
        ("B1", json!(10)),
        ("B2", json!(20)),
        ("B3", json!("=B2*2")), // computed, read dynamically
        ("C1", json!("=INDIRECT(\"B\"&A1)")),
        ("C2", json!("=OFFSET(B1,A1,0)")),
        ("C3", json!("=SUM(OFFSET(B1,0,0,A1,1))")),
        ("C4", json!("=INDIRECT(\"B3\")")), // static: resolved at build time
    ]);
    let (p, r) = build(&m, &spec(&["A1"], &["C1:C4"]));
    assert!(r.dynamic_refs);
    for (row, want) in [
        (1, [10, 20, 10, 40]),
        (2, [20, 40, 30, 40]),
        (3, [40, 0, 70, 40]),
    ] {
        let out = p
            .execute(&json!({"Sheet1!A1": row}), &Overrides::default())
            .unwrap();
        let got: Vec<Json> = out["Sheet1!C1:C4"]
            .as_array()
            .unwrap()
            .iter()
            .map(|r| r[0].clone())
            .collect();
        // OFFSET(B1,3,0) is B4, empty: the formula holds 0.
        let want: Vec<Json> = want.iter().map(|v| json!(v)).collect();
        assert_eq!(got, want, "row={row}");
    }
    // The static INDIRECT became a plain reference.
    let (p, r) = build(
        &sheet(&[("A1", json!("=INDIRECT(\"B1\")")), ("B1", json!(4))]),
        &spec(&[], &["A1"]),
    );
    assert!(!r.dynamic_refs);
    assert_eq!(
        p.execute(&Json::Null, &Overrides::default()).unwrap()["Sheet1!A1"],
        json!(4)
    );
}

#[test]
fn locale_is_fixed_at_build_time() {
    use sheetrs_compiler::locale::{LocaleOverride, from_tag};
    let regional = |decimal, group| LocaleOverride {
        decimal: Some(decimal),
        group: Some(group),
        date_order: Some("dmy".into()),
        ..LocaleOverride::default()
    };
    let m = sheet(&[
        ("A1", json!(1234.5)),
        ("B1", json!("=TEXT(A1,\"#,##0.00\")&\" \"&(1/2)")),
        ("C1", json!("=A1*2")),
    ]);
    let key = "Sheet1!B1";
    let with = |tag: &str, fold: bool| {
        let mut s = Spec {
            fold,
            ..spec(&[], &["B1"])
        };
        let over = match tag {
            "es-CO" => regional(',', '.'),
            "es-MX" => regional('.', ','),
            _ => LocaleOverride::default(),
        };
        s.locale = Some(from_tag(tag, &over).unwrap().0);
        let (p, _) = build(&m, &s);
        p.execute(&Json::Null, &Overrides::default()).unwrap()[key].clone()
    };

    // Folded or not, the build's locale is the one in force.
    for fold in [true, false] {
        assert_eq!(with("en", fold), json!("1,234.50 0.5"));
        assert_eq!(with("es", fold), json!("1.234,50 0,5"));
        assert_eq!(with("es-CO", fold), json!("1.234,50 0,5"));
        assert_eq!(with("es-MX", fold), json!("1,234.50 0.5"));
    }

    // Formulas unfolded and locale static are independent choices.
    let (_, r) = build(
        &m,
        &Spec {
            fold: false,
            ..spec(&[], &["C1"])
        },
    );
    assert_eq!(r.formulas_kept, 1);

    // The date system is not a run-time choice: it is the document's (build
    // time, for a model with no file behind it).
    let m = sheet(&[("A1", json!("=DATE(2024,1,1)"))]);
    let (p, _) = build(&m, &spec(&[], &["A1"]));
    assert_eq!(
        p.execute(&Json::Null, &Overrides::default()).unwrap()["Sheet1!A1"],
        json!(45292)
    );
}

#[test]
fn volatile_functions_are_not_folded() {
    let m = sheet(&[("A1", json!("=RAND()")), ("B1", json!("=TODAY()"))]);
    let (p, r) = build(&m, &spec(&[], &["A1", "B1"]));
    assert_eq!(r.formulas_kept, 2);
    let o = Overrides {
        rand_seed: Some(7),
        now: Some(45292.5),
        ..Overrides::default()
    };
    let x = p.execute(&Json::Null, &o).unwrap();
    assert_eq!(x, p.execute(&Json::Null, &o).unwrap()); // deterministic when seeded
    assert_eq!(x["Sheet1!B1"], json!(45292));
    let a = x["Sheet1!A1"].as_f64().unwrap();
    assert!((0.0..1.0).contains(&a));
}

#[test]
fn program_roundtrips_through_json() {
    let (p, _) = build(&model(), &spec(&["A1"], &["C2"]));
    let text = serde_json::to_string(&p).unwrap();
    let q: sheetrs_compiler::Program = serde_json::from_str(&text).unwrap();
    assert_eq!(p, q);
    let o = Overrides::default();
    assert_eq!(
        p.execute(&json!({"Sheet1!A1": 7}), &o).unwrap(),
        q.execute(&json!({"Sheet1!A1": 7}), &o).unwrap()
    );
}

#[test]
fn defined_names_bind_and_parse() {
    let mut m = sheet(&[("A1", json!(5)), ("B1", json!("=Rate*2"))]);
    m.names.insert("Rate".into(), "Sheet1!$A$1".into());
    let (p, _) = build(&m, &spec(&["Rate"], &["B1"]));
    assert_eq!(
        p.execute(&json!({"Sheet1!A1": 4}), &Overrides::default())
            .unwrap()["Sheet1!B1"],
        json!(8)
    );
}

#[test]
fn bad_targets_fail_early() {
    let e = sheetrs_compiler::compile(&model(), &spec(&["Nope!A1"], &["C2"])).unwrap_err();
    assert!(e.0.contains("cannot resolve"));
    let e = sheetrs_compiler::compile(&model(), &spec(&["A1", "A1"], &["C2"])).unwrap_err();
    assert!(e.0.contains("declared twice"));
}

#[test]
fn query_levels_keep_progressively_more() {
    use sheetrs_compiler::Program;
    use sheetrs_compiler::compile::QueryLevel;
    let mut m = sheet(&[
        ("A1", json!(2)),
        ("B1", json!("=A1*10")),
        ("C1", json!("=B1+1")),
        ("D1", json!("=A1+100")),
    ]);
    m.names.insert("Total".into(), "Sheet1!C1".into());
    let at = |level, ins: &[&str]| {
        let s = Spec {
            query: level,
            inputs: ins.iter().map(|t| BindingSpec::named("a", t)).collect(),
            outputs: vec![BindingSpec::named("ten", "B1")],
            ..Spec::default()
        };
        build(&m, &s).0
    };
    let ask = |p: &Program, inputs: &Json, q: &[&str]| {
        let q: Vec<String> = q.iter().map(|s| (*s).to_string()).collect();
        p.query(inputs, &Overrides::default(), &q, Default::default())
    };
    let none = Json::Null;

    // declared: the output, by name or by range; C1 and D1 are pruned away.
    let p = at(QueryLevel::Declared, &[]);
    assert_eq!(
        ask(&p, &none, &["ten"]).unwrap(),
        json!([{"name": "ten", "range": "Sheet1!B1", "values": [[20]]}])
    );
    assert_eq!(ask(&p, &none, &["B1"]).unwrap()[0]["values"], json!([[20]]));
    assert!(
        ask(&p, &none, &["C1"])
            .unwrap_err()
            .0
            .contains("outputs = declared")
    );

    // named: a defined name is read by name (any case) or by range; D1 is not.
    let p = at(QueryLevel::Named, &[]);
    assert_eq!(
        ask(&p, &none, &["total"]).unwrap(),
        json!([{"name": "Total", "range": "Sheet1!C1", "values": [[21]]}])
    );
    assert_eq!(
        ask(&p, &none, &["Sheet1!C1"]).unwrap()[0]["values"],
        json!([[21]])
    );
    assert!(ask(&p, &none, &["D1"]).is_err());

    // any: every cell, in request order, repeats allowed, following the inputs.
    let p = at(QueryLevel::Any, &["A1"]);
    let out = ask(&p, &json!({"a": 5}), &["D1", "Sheet1!A1:C1", "Total", "D1"]).unwrap();
    assert_eq!(out[0]["values"], json!([[105]]));
    assert_eq!(out[1]["values"], json!([[5, 50, 51]]));
    assert_eq!(
        (out[2]["name"].clone(), out[2]["values"].clone()),
        (json!("Total"), json!([[51]]))
    );
    assert_eq!(out[3], out[0]);
    assert!(ask(&p, &none, &["Z9:Z10"]).is_ok());
    assert!(ask(&p, &none, &["Nope!A1"]).is_err());
    // Open-ended ranges end where the data does.
    let out = ask(&p, &none, &["Sheet1!A:A", "Sheet1!1:1", "Sheet1!E:F"]).unwrap();
    assert_eq!(out[0], json!({"range": "Sheet1!A1", "values": [[2]]}));
    assert_eq!(out[1]["range"], "Sheet1!A1:D1");
    assert_eq!(out[2], json!({"range": "", "values": []}));
    // Columns instead of rows.
    let cols = p
        .query(
            &none,
            &Overrides::default(),
            &["A1:C1".into()],
            sheetrs_compiler::run::MajorDimension::Columns,
        )
        .unwrap();
    assert_eq!(cols[0]["values"], json!([[2], [20], [21]]));
}

#[test]
fn requests_are_inputs_and_ranges() {
    use sheetrs_compiler::compile::QueryLevel;
    let m = sheet(&[
        ("A1", json!(2)),
        ("B1", json!("=A1*10")),
        ("C1", json!("=B1+1")),
    ]);
    let build_at = |level| {
        let s = Spec {
            query: level,
            inputs: vec![BindingSpec::named("a", "A1")],
            outputs: vec![
                BindingSpec::named("ten", "B1"),
                BindingSpec::named("eleven", "C1"),
            ],
            ..Spec::default()
        };
        build(&m, &s).0
    };
    let call = |p: &sheetrs_compiler::Program, r: Json| p.handle(&r, &Overrides::default());

    // Without ranges: the outputs declared at build time, under their names.
    let p = build_at(QueryLevel::Declared);
    assert_eq!(
        call(&p, json!({"inputs": {"a": 3}})).unwrap(),
        json!({"outputs": {"ten": 30, "eleven": 31}})
    );
    assert_eq!(
        call(&p, Json::Null).unwrap(),
        json!({"outputs": {"ten": 20, "eleven": 21}})
    );
    // With ranges: exactly those, and no more.
    assert_eq!(
        call(&p, json!({"inputs": {"a": 3}, "ranges": ["eleven"]})).unwrap(),
        json!({"valueRanges": [{"name": "eleven", "range": "Sheet1!C1", "values": [[31]]}]})
    );
    // Mistakes are named.
    for (req, want) in [
        (json!({"a": 3}), "unknown key `a`"),
        (json!({"inputs": {"zzz": 1}}), "unknown input `zzz`"),
        (json!({"ranges": []}), "non-empty array"),
        (json!({"ranges": [1]}), "array of strings"),
        (json!({"majorDimension": "DIAGONAL"}), "majorDimension"),
        (json!([1]), "JSON object"),
    ] {
        let e = call(&p, req.clone()).unwrap_err().0;
        assert!(e.contains(want), "{req}: {e}");
    }

    // `any` has no default answer: ranges are required.
    let p = build_at(QueryLevel::Any);
    let e = call(&p, json!({"inputs": {"a": 3}})).unwrap_err().0;
    assert!(e.contains("must list `ranges`"), "{e}");
    assert_eq!(
        call(&p, json!({"inputs": {"a": 3}, "ranges": ["B1:C1"]})).unwrap(),
        json!({"valueRanges": [{"range": "Sheet1!B1:C1", "values": [[30, 31]]}]})
    );
}

#[test]
fn settings_sheet_errors_and_defaults() {
    use sheetrs_compiler::Value;
    use sheetrs_compiler::config::Config;
    let mut m = sheet(&[("A1", json!(1)), ("B1", json!("=A1+1"))]);
    let s = m.add_sheet("settings");
    for (c, h) in ["key", "value", "description", "meta_type"]
        .iter()
        .enumerate()
    {
        m.set(s, &format!("{}1", (b'A' + c as u8) as char), Value::text(h));
    }
    m.set(s, "A2", Value::text("version"));
    m.set(s, "D2", Value::text("info"));
    // Declared outputs are mandatory when the sheet exists.
    let e = Config::default().spec(&m).unwrap_err().0;
    assert!(e.contains("declares no output"), "{e}");
    // ... unless the caller asks for ranges.
    let any = Config {
        outputs: Some("any".into()),
        ..Config::default()
    };
    let spec = any.spec(&m).unwrap();
    assert_eq!(spec.skip_sheet, Some(s));
    assert_eq!(spec.meta.len(), 1);
    let (p, _) = build(&m, &spec);
    assert!(p.outputs.is_empty());
    let out = p
        .query(
            &Json::Null,
            &Overrides::default(),
            &["B1".into()],
            Default::default(),
        )
        .unwrap();
    assert_eq!(out[0]["values"], json!([[2]]));
    // The settings sheet cannot be read.
    let e = p
        .query(
            &Json::Null,
            &Overrides::default(),
            &["settings!A1".into()],
            Default::default(),
        )
        .unwrap_err()
        .0;
    assert!(e.contains("cannot be read"), "{e}");
}

#[test]
fn declarations_are_validated_together() {
    let m = sheet(&[
        ("A1", json!(5)),
        ("B1", json!("=A1+1")),
        ("B2", json!("=A1+2")),
        ("C1", json!(7)),
    ]);
    let named = |ins: &[(&str, &str)], outs: &[(&str, &str)]| Spec {
        inputs: ins.iter().map(|(n, t)| BindingSpec::named(n, t)).collect(),
        outputs: outs.iter().map(|(n, t)| BindingSpec::named(n, t)).collect(),
        ..Spec::default()
    };
    // Inputs that hold formulas, and outputs with nothing in them, are all
    // listed in one error.
    let e = sheetrs_compiler::compile(
        &m,
        &named(
            &[("a", "A1"), ("b", "B1:B2"), ("c", "C1:C1"), ("d", "B2")],
            &[("x", "A1"), ("y", "E1:F9"), ("z", "H1")],
        ),
    )
    .unwrap_err()
    .0;
    assert!(
        e.contains("input `b` (B1:B2) holds formulas in B1, B2"),
        "{e}"
    );
    assert!(e.contains("input `d` (B2) holds formulas in B2"), "{e}");
    assert!(!e.contains("input `a`") && !e.contains("input `c`"), "{e}");
    assert!(e.contains("output `y` (E1:F9) is empty"), "{e}");
    assert!(e.contains("output `z` (H1) is empty"), "{e}");
    assert!(!e.contains("output `x`"), "{e}");
    assert_eq!(e.lines().count(), 4, "{e}");

    // A constant keeps its value as the default; declaring it does not make it
    // required; an empty input is allowed and defaults to blank.
    let m = sheet(&[("A1", json!(5)), ("B1", json!("=A1*2+N(A2)"))]);
    let s = named(&[("a", "A1"), ("blank", "A2")], &[("out", "B1")]);
    let (p, _) = build(&m, &s);
    let run = |j: Json| p.execute(&j, &Overrides::default()).unwrap()["out"].clone();
    assert_eq!(run(json!({})), json!(10));
    assert_eq!(run(json!({"a": 7})), json!(14));
    assert_eq!(run(json!({"blank": 1})), json!(11));
    assert_eq!(run(json!({"a": null})), json!(0), "null sets a blank");
}

#[test]
fn macros_are_refused_unless_strict_is_off() {
    use sheetrs_compiler::config::Config;
    let mut m = sheet(&[("A1", json!(1)), ("B1", json!("=A1+1"))]);
    m.macros = true;
    let strict = Config::default().spec(&m).unwrap_err().0;
    assert!(
        strict.contains("macros") && strict.contains("--no-strict"),
        "{strict}"
    );
    let lax = Config {
        strict: Some(false),
        ..Config::default()
    };
    let spec = lax.spec(&m).unwrap();
    assert!(
        spec.notes
            .iter()
            .any(|n| n.starts_with("warning: the workbook contains macros"))
    );
    let (p, _) = build(&m, &spec);
    assert!(p.diagnostics.iter().any(|d| d.contains("contains macros")));
    // No macros, no complaint.
    m.macros = false;
    assert!(Config::default().spec(&m).is_ok());
}

#[test]
fn a_formula_ending_in_a_blank_keeps_an_empty_result() {
    // B1 =A1 with A1 empty. B1 shows 0 and is no longer blank, but it is an empty result
    // that still equals "" and is not a number.
    let build_for = |fold: bool| {
        let mut m = sheetrs_compiler::Model::new();
        let s = m.add_sheet("S");
        m.set_formula(s, "B1", "A1");
        for (cell, f) in [
            ("C1", "B1=\"\""),
            ("D1", "B1=0"),
            ("E1", "ISNUMBER(B1)"),
            ("F1", "ISBLANK(B1)"),
            ("G1", "B1&\"x\""),
            ("H1", "IF(B1=\"\",\"empty\",\"not\")"),
        ] {
            m.set_formula(s, cell, f);
        }
        let spec = Spec {
            fold,
            ..spec(&[], &["B1:H1"])
        };
        let (p, _) = build(&m, &spec);
        p.execute(&Json::Null, &Overrides::default()).unwrap()["S!B1:H1"][0].clone()
    };
    for fold in [true, false] {
        assert_eq!(
            build_for(fold),
            json!([0, true, true, false, false, "0x", "empty"]),
            "fold={fold}"
        );
    }
}

#[test]
fn conditional_sums_read_only_the_selected_cells() {
    // B1 holds the SUMIF and sits in its own sum range, but A1 does not select
    // it: no cycle, as in LibreOffice and Excel.
    let m = sheet(&[
        ("A1", json!(1)),
        ("A2", json!(2)),
        ("A3", json!(2)),
        ("B2", json!(10)),
        ("B3", json!(20)),
        ("B1", json!("=SUMIF(A1:A3,2,B1:B3)")),
        ("C1", json!("=SUMIFS(B1:B3,A1:A3,2)")),
        ("D1", json!("=AVERAGEIF(A1:A3,2,B1:B3)")),
    ]);
    let (p, r) = build(&m, &spec(&[], &["B1:D1"]));
    assert_eq!(r.cycles, 1, "the static graph still sees the loop");
    let got = p.execute(&Json::Null, &Overrides::default()).unwrap();
    assert_eq!(got["Sheet1!B1:D1"], json!([[30, 30, 15]]));
}

#[test]
fn conditional_counts_accept_array_criteria() {
    assert_eq!(
        calc("=SUMPRODUCT((COUNTIF({1;2;3},{2;5;1})=0)*1)"),
        json!(1)
    );
    assert_eq!(calc("=SUMPRODUCT(COUNTIFS({1;2;2;3},{2;3;4}))"), json!(3));
}

#[test]
fn a_one_cell_array_formula_over_a_range_shows_its_first_element() {
    let mut m = sheet(&[("A1", json!(7)), ("A2", json!(8)), ("B2", json!(9))]);
    m.cells.insert(
        sheetrs_engine::addr::CellRef::new(0, 0, 2),
        sheetrs_compiler::model::ModelCell {
            value: sheetrs_compiler::Value::Empty,
            formula: Some("A1:B2".into()),
            array: true,
        },
    );
    let (p, _) = build(&m, &spec(&[], &["C1"]));
    assert_eq!(
        p.execute(&Json::Null, &Overrides::default()).unwrap()["Sheet1!C1"],
        json!(7)
    );
}

#[test]
fn an_xlsx_is_read_by_its_format_and_computed_like_any_other_file() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/assets/loan.xlsx");
    let m = sheetrs_compiler::loader::load(&path).unwrap();
    // The format decides what the file carries: an XLSX has no case-sensitivity setting and
    // its dates count from 1899-12-30 unless it says `date1904`.
    assert!(m.from_xlsx);
    assert_eq!(
        m.date_system,
        sheetrs_compiler::options::DateSystem::NULL_1899
    );
    assert_eq!(m.settings.string_ref_excel, Some(true));
}

#[test]
fn formula_text_follows_libreoffice_notation_and_locale() {
    let m = sheet(&[
        ("A1", json!(5)),
        ("A2", json!("=SUM(A1:A1,2.5)")),
        ("B1", json!("=_xlfn.FORMULATEXT(A2)")),
        ("B2", json!("=_xlfn.FORMULATEXT(A1)")),
        ("B3", json!("=_xlfn.ISFORMULA(A2)")),
        ("B4", json!("=_xlfn.ISFORMULA(A1)")),
    ]);
    let run = |tag: &str| {
        let mut s = spec(&[], &["B1:B4"]);
        s.locale = Some(
            sheetrs_compiler::locale::from_tag(tag, &Default::default())
                .unwrap()
                .0,
        );
        let (p, _) = build(&m, &s);
        p.execute(&Json::Null, &Overrides::default()).unwrap()["Sheet1!B1:B4"].clone()
    };
    assert_eq!(
        run("en"),
        json!([["=SUM(A1:A1,2.5)"], [{"error": "#N/A"}], [true], [false]])
    );
    // The separators follow the decimal one, as in LibreOffice.
    assert_eq!(run("es")[0], json!(["=SUM(A1:A1;2,5)"]));
}

#[test]
fn let_unique_and_long_text() {
    // LET names stand for their values in what follows (nested, shadowing, the
    // `_xlpm.` prefix Excel writes).
    assert_eq!(calc("=LET(x,2,y,x*3,x+y)"), json!(8));
    assert_eq!(
        calc("=_xlfn.LET(_xlpm.v,ROUND(123.4,-1),_xlpm.v+1)"),
        json!(121)
    );
    // LibreOffice keeps the outer binding (Excel would shadow it: 22).
    assert_eq!(calc("=LET(x,1,LET(x,x+10,x*2))"), json!(2));
    // UNIQUE keeps the first of each, comparing text without regard to case.
    assert_eq!(
        calc("=TEXTJOIN(\",\",TRUE,UNIQUE({\"a\";\"B\";\"b\";\"c\";\"a\"}))"),
        json!("a,B,c")
    );
    assert_eq!(calc("=SUM(UNIQUE({1;1;2;3;3}))"), json!(6));
    assert_eq!(calc("=SUM(UNIQUE({1;1;2;3;3},FALSE,TRUE))"), json!(2));
    assert_eq!(calc("=COLUMNS(UNIQUE({1,1,2},TRUE))"), json!(2));
    // LibreOffice 26.8 reads it as an unknown function (the LONGTEXT feature, off by default,
    // joins constant pieces: see the next tests).
    assert_eq!(
        calc("=_xlfn._longtext(\"abc\",\"def\")"),
        json!({"error": "#NAME?"})
    );
}

#[test]
fn the_longtext_feature_joins_constant_pieces_only() {
    let mut m = sheetrs_compiler::Model::new();
    let s = m.add_sheet("Sheet1");
    m.set_formula(s, "A1", "=_xlfn._longtext(\"ab\",\"cd\")");
    m.set_formula(s, "A2", "=_xlfn._longtext(\"ab\",B1)");
    let mut sp = spec(&[], &["A1", "A2"]);
    sp.features.longtext = true;
    let (p, _) = build(&m, &sp);
    let out = p.execute(&Json::Null, &Overrides::default()).unwrap();
    assert_eq!(out["Sheet1!A1"], json!("abcd"));
    assert_eq!(out["Sheet1!A2"], json!({"error": "#NAME?"}));
    let note: Vec<_> = p
        .diagnostics
        .iter()
        .filter(|d| d.contains("_longtext"))
        .collect();
    assert_eq!(note.len(), 1, "{:?}", p.diagnostics);
    assert!(note[0].starts_with("Sheet1!A2: "), "{}", note[0]);
}

#[test]
fn a_longtext_wrapper_is_reported_per_cell() {
    // Without the feature (the default), as LibreOffice 26.8 reads it.
    let mut m = sheetrs_compiler::Model::new();
    let s = m.add_sheet("Sheet1");
    m.set_formula(s, "A1", "=_xlfn._longtext(\"ab\",\"cd\")");
    m.set_formula(s, "A2", "=1+1");
    let (p, _) = build(&m, &spec(&[], &["A1", "A2"]));
    let note: Vec<_> = p
        .diagnostics
        .iter()
        .filter(|d| d.contains("_longtext"))
        .collect();
    assert_eq!(note.len(), 1, "{:?}", p.diagnostics);
    assert!(note[0].starts_with("Sheet1!A1: "), "{}", note[0]);
    let out = p.execute(&Json::Null, &Overrides::default()).unwrap();
    assert_eq!(out["Sheet1!A1"], json!({"error": "#NAME?"}));
}

#[test]
fn error_literals_are_capitals_only_in_libreoffice() {
    assert_eq!(calc("#REF!"), err("#REF!"));
    // LibreOffice reads `#ref!` as an unknown name.
    assert_eq!(calc("#ref!"), err("#NAME?"));
    // An XLSX import in LibreOffice 26.8 also reads a bare `#REF!` as an unknown name
    // (the other error literals are fine).
    let mut m = sheetrs_compiler::Model::new();
    m.from_xlsx = true;
    let s = m.add_sheet("Sheet1");
    m.set_formula(s, "A1", "#REF!+5");
    m.set_formula(s, "A2", "#N/A");
    let (p, _) = build(&m, &spec(&[], &["A1:A2"]));
    assert_eq!(
        p.execute(&Json::Null, &Overrides::default()).unwrap()["Sheet1!A1:A2"],
        json!([[err("#NAME?")], [err("#N/A")]])
    );
}

#[test]
fn exact_lookups_over_big_ranges_use_an_index_with_the_scan_semantics() {
    // 300 rows: text keys (the first of a duplicate wins, case does not matter),
    // numbers, and an empty key cell; every case must answer as the scan does.
    let mut cells: Vec<(String, Json)> = Vec::new();
    for r in 1..=300 {
        let key = match r {
            5 | 9 => json!("Dup"),
            7 => json!(42),
            8 => continue, // an empty key cell
            _ => json!(format!("k{r}")),
        };
        cells.push((format!("A{r}"), key));
        cells.push((format!("B{r}"), json!(r * 10)));
    }
    let mut formulas = vec![
        ("D1", "=VLOOKUP(\"dup\",A1:B300,2,FALSE)"),
        ("D2", "=VLOOKUP(42,A1:B300,2,FALSE)"),
        ("D3", "=VLOOKUP(\"K250\",A1:B300,2,FALSE)"),
        ("D4", "=VLOOKUP(\"nope\",A1:B300,2,FALSE)"),
        ("D5", "=VLOOKUP(0,A1:B300,2,FALSE)"),
        ("D6", "=VLOOKUP(\"k2*\",A1:B300,2,FALSE)"),
        // The same empty-key lookup over a range too small for the index.
        ("D7", "=VLOOKUP(0,A1:B20,2,FALSE)"),
    ];
    let owned: Vec<(&str, Json)> = cells
        .iter()
        .map(|(a, v)| (a.as_str(), v.clone()))
        .chain(formulas.drain(..).map(|(a, f)| (a, json!(f))))
        .collect();
    let m = sheet(&owned);
    let (p, _) = build(&m, &spec(&[], &["D1:D7"]));
    let got = p.execute(&Json::Null, &Overrides::default()).unwrap();
    assert_eq!(
        got["Sheet1!D1:D7"],
        json!([[50], [70], [2500], [{"error": "#N/A"}], [0], [20], [0]])
    );
}

#[test]
fn indirect_and_address_follow_the_documents_reference_syntax() {
    let run = |excel: bool, formula: &str| {
        let mut m = sheet(&[("A1", json!(7))]);
        m.settings.string_ref_excel = Some(excel);
        let s = m.add_sheet("Other");
        m.set_formula(s, "A1", formula);
        let (p, _) = build(&m, &spec(&[], &["Other!A1"]));
        p.execute(&Json::Null, &Overrides::default()).unwrap()["Other!A1"].clone()
    };
    // An XLSX (or an XLSX saved as ODS): `Sheet!A1`.
    assert_eq!(run(true, "=INDIRECT(\"Sheet1!A1\")"), json!(7));
    assert_eq!(
        run(true, "=INDIRECT(\"Sheet1.A1\")"),
        json!({"error": "#REF!"})
    );
    assert_eq!(
        run(true, "=ADDRESS(1,1,1,1,\"My sheet\")"),
        json!("'My sheet'!$A$1")
    );
    // A document created in LibreOffice: `Sheet.A1`.
    assert_eq!(run(false, "=INDIRECT(\"Sheet1.A1\")"), json!(7));
    assert_eq!(
        run(false, "=INDIRECT(\"$Sheet1.$A$1:Sheet1.A1\")"),
        json!(7)
    );
    assert_eq!(
        run(false, "=INDIRECT(\"Sheet1!A1\")"),
        json!({"error": "#REF!"})
    );
    assert_eq!(
        run(false, "=ADDRESS(1,1,1,1,\"My sheet\")"),
        json!("'My sheet'.$A$1")
    );
    assert_eq!(run(false, "=ADDRESS(1,1,1,0,\"S\")"), json!("S!R1C1"));
}

/// An XLSX with one sheet `S` holding `sheet_data`.
fn xlsx_with(sheet_data: &str) -> Vec<u8> {
    use std::io::Write;
    let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_string()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.to_string()),
        ("xl/workbook.xml", r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_string()),
        ("xl/_rels/workbook.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#.to_string()),
        ("xl/worksheets/sheet1.xml", format!(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{sheet_data}</sheetData></worksheet>"#)),
    ];
    for (name, body) in parts {
        zip.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(body.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// A shared formula whose cell would read outside the sheet is `#REF!` there, as LibreOffice
/// computes it: the workbook builds, and the cell is `#REF!` when it runs.
#[test]
fn a_shared_formula_moved_outside_the_sheet_builds_and_is_a_ref_error() {
    let dir = tempfile::tempdir().unwrap();
    // `=XFD2` in XFC1, shared with XFD1, whose copy would read column XFE.
    let path = dir.path().join("beyond.xlsx");
    std::fs::write(
        &path,
        xlsx_with(
            r#"<row r="2"><c r="XFD2"><v>5</v></c></row><row r="1"><c r="XFC1"><f t="shared" ref="XFC1:XFD1" si="0">XFD2</f><v>5</v></c><c r="XFD1"><f t="shared" si="0"/><v>0</v></c></row>"#,
        ),
    )
    .unwrap();
    let m = sheetrs_compiler::loader::load(&path).unwrap();
    let (p, _) = build(&m, &spec(&[], &["XFC1", "XFD1"]));
    let out = p.execute(&Json::Null, &Overrides::default()).unwrap();
    assert_eq!(out["S!XFC1"], json!(5));
    assert_eq!(out["S!XFD1"], err("#REF!"));
}
