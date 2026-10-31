//! What the engine does not support stops the build: it is a replacement of LibreOffice or it
//! says it is not one for the workbook.

use crate::common::*;
use serde_json::json;
use sheetrs_compiler::{Spec, Value, compile};

fn build_with(f: &str) -> Result<sheetrs_compiler::Report, String> {
    let mut m = sheetrs_compiler::Model::new();
    let s = m.add_sheet("Sheet1");
    m.set_formula(s, "A1", f);
    compile(&m, &spec(&[], &["A1"]))
        .map(|(_, r)| r)
        .map_err(|e| e.0)
}

#[test]
fn a_function_of_libreoffice_the_engine_lacks_stops_the_build() {
    let e = build_with("WEBSERVICE(\"x\")+FILTERXML(\"a\",\"b\")").unwrap_err();
    assert!(
        e.contains("WEBSERVICE") && e.contains("FILTERXML") && e.contains("Sheet1!A1"),
        "{e}"
    );
    // A name LibreOffice does not know either is `#NAME?` there too.
    assert_eq!(calc("NOPE(1)"), err("#NAME?"));
    assert!(build_with("NOPE(1)").is_ok());
    // `sheetc inspect` counts them instead.
    let mut m = sheetrs_compiler::Model::new();
    let s = m.add_sheet("Sheet1");
    m.set_formula(s, "A1", "WEBSERVICE(\"x\")");
    let spec = Spec {
        inspect: true,
        ..spec(&[], &["A1"])
    };
    let (_, r) = compile(&m, &spec).unwrap();
    assert!(r.unsupported.contains("WEBSERVICE"));
}

#[test]
fn syntax_the_parser_does_not_read_stops_the_build() {
    for f in ["[1]Sheet1!B3", "Table1[Col]", "A1:A5#", "SUM(1,"] {
        let e = build_with(f).unwrap_err();
        assert!(e.contains("syntax") && e.contains(f), "{f}: {e}");
    }
}

#[test]
fn a_text_code_with_a_language_the_engine_has_no_names_for_stops_the_build() {
    let e = build_with("TEXT(45367,\"[$-407]dddd\")").unwrap_err();
    assert!(e.contains("[$-407]dddd"), "{e}");
    for ok in [
        "TEXT(45367,\"[$-409]dddd\")",
        "TEXT(45367,\"[$-40A]dddd\")",
        "TEXT(45367,\"[$€-407]#,##0\")",
    ] {
        assert!(build_with(ok).is_ok(), "{ok}");
    }
    // A tag LibreOffice rejects is an error value, as there.
    assert_eq!(calc("TEXT(45367,\"[$-en-US]dddd\")"), err("#VALUE!"));
}

#[test]
fn a_document_option_the_engine_does_not_support_stops_the_build() {
    let mut m = sheetrs_compiler::Model::new();
    let s = m.add_sheet("Sheet1");
    m.set(s, "A1", Value::Number(1.0));
    m.unsupported_options
        .push("the document uses 'precision as shown'".into());
    let e = compile(&m, &spec(&[], &["A1"])).unwrap_err().0;
    assert!(e.contains("precision as shown"), "{e}");
}

#[test]
fn a_date_outside_the_range_of_libreoffice_stops_the_build() {
    for f in [
        "YEAR(-700000)",
        "TEXT(20000000,\"yyyy\")",
        "EDATE(1,-100000)",
    ] {
        let e = build_with(f).unwrap_err();
        assert!(
            e.contains("0001-01-01") && e.contains("Sheet1!A1"),
            "{f}: {e}"
        );
    }
    for ok in ["DATE(32767,12,31)", "YEAR(1)", "YEAR(11274306)"] {
        assert!(build_with(ok).is_ok(), "{ok}");
    }
}

#[test]
fn xlookup_binary_search_over_a_range_stops_the_build() {
    for f in [
        "XLOOKUP(2,B1:B5,C1:C5,\"NF\",0,2)",
        "XLOOKUP(2,B1:B5,C1:C5,\"NF\",1,-2)",
    ] {
        let e = build_with(f).unwrap_err();
        assert!(e.contains("XLOOKUP") && e.contains("Sheet1!A1"), "{f}: {e}");
    }
    for ok in [
        "XLOOKUP(2,B1:B5,C1:C5,\"NF\",0,1)",
        "XLOOKUP(2,B1:B5,C1:C5,\"NF\",1,-1)",
        "XLOOKUP(2,{1,2,3},{4,5,6},\"NF\",0,2)",
    ] {
        assert!(build_with(ok).is_ok(), "{ok}");
    }
}

#[test]
fn distributions_where_libreoffice_is_not_reproduced_stop_the_build() {
    // F.INV(1, …) is the end of LibreOffice's search range; more than 1E7 degrees of freedom or
    // shape is where its continued fraction loses digits.
    for f in [
        "F.INV(1,5,8)",
        "T.DIST(2,100000000,TRUE)",
        "F.DIST(2,5,1000000000,TRUE)",
        "BETA.DIST(0.5,1E9,1E9,TRUE)",
        "T.INV(0.9,2E7)",
    ] {
        let e = build_with(f).unwrap_err();
        assert!(e.contains("Sheet1!A1"), "{f}: {e}");
    }
    for ok in [
        "F.INV(0.9,5,8)",
        // The legacy FINV is the right tail: 1 is zero there.
        "FINV(1,5,8)",
        "T.DIST(2,10000000,TRUE)",
        "F.DIST(2,5,1E7,TRUE)",
        "BETA.DIST(0.5,1E7,1E7,TRUE)",
        // LibreOffice's non-convergence is an error value, as there.
        "CHISQ.DIST(10000000,10000000,TRUE)",
    ] {
        assert!(build_with(ok).is_ok(), "{ok}");
    }
    assert_eq!(calc("CHISQ.DIST(10000000,10000000,TRUE)"), err("#VALUE!"));
}

#[test]
fn sorting_by_keys_with_an_error_is_not_reproduced() {
    let mut m = sheetrs_compiler::Model::new();
    let s = m.add_sheet("Sheet1");
    for (i, v) in [1.0, 2.0, 3.0].into_iter().enumerate() {
        m.set(s, &format!("B{}", i + 1), Value::Number(v));
    }
    m.set_formula(s, "C1", "1");
    m.set_formula(s, "C2", "NA()");
    m.set_formula(s, "C3", "3");
    // The keys are a range with an error: the rows are not placed as LibreOffice places them.
    m.set_formula(s, "A1", "INDEX(SORTBY(B1:B3,C1:C3),1)");
    let e = compile(&m, &spec(&[], &["A1"])).unwrap_err().0;
    assert!(e.contains("SORTBY") && e.contains("Sheet1!A1"), "{e}");
    // The same range without an error, and `SORT` of the range with it (the errors last, as 0).
    m.set_formula(s, "C2", "2");
    assert!(compile(&m, &spec(&[], &["A1"])).is_ok());
    m.set_formula(s, "C2", "NA()");
    m.set_formula(s, "A1", "INDEX(SORT(C1:C3),3)");
    assert!(compile(&m, &spec(&[], &["A1"])).is_ok());
}

#[test]
fn errortype_of_an_error_with_several_internal_codes_is_not_reproduced() {
    // LibreOffice answers 502, 519 or others for what the engine knows only as #VALUE!.
    for f in ["ERRORTYPE(\"x\"*1)", "ERRORTYPE(VALUE(\"x\"))"] {
        let e = build_with(f).unwrap_err();
        assert!(
            e.contains("ERRORTYPE") && e.contains("Sheet1!A1"),
            "{f}: {e}"
        );
    }
    // The errors with one code are answered.
    assert_eq!(calc("ERRORTYPE(NA())"), json!(32767));
    assert_eq!(calc("ERRORTYPE(1/0)"), json!(532));
    assert_eq!(calc("ERRORTYPE(5)"), err("#N/A"));
}

#[test]
fn counting_the_records_of_a_database_with_field_zero_is_not_reproduced() {
    let mut m = sheetrs_compiler::Model::new();
    let s = m.add_sheet("Sheet1");
    m.set(s, "B1", Value::text("Age"));
    m.set(s, "B2", Value::Number(30.0));
    m.set(s, "C1", Value::text("Age"));
    m.set(s, "C2", Value::Number(30.0));
    for f in ["DCOUNT(B1:B2,0,C1:C2)", "DCOUNTA(B1:B2,0,C1:C2)"] {
        m.set_formula(s, "A1", f);
        let e = compile(&m, &spec(&[], &["A1"])).unwrap_err().0;
        assert!(e.contains("field 0") && e.contains("Sheet1!A1"), "{f}: {e}");
    }
    m.set_formula(s, "A1", "DCOUNT(B1:B2,1,C1:C2)");
    assert!(compile(&m, &spec(&[], &["A1"])).is_ok());
}

#[test]
fn regression_statistics_that_are_rounding_noise_and_equal_bins_are_not_reproduced() {
    // A perfect fit in the logarithms: LibreOffice answers F = 4E31 from the noise of its
    // computation.
    let e = build_with("INDEX(LOGEST({1;2;4;8},{1;2;3;4},TRUE,TRUE),4,1)").unwrap_err();
    assert!(e.contains("LOGEST") && e.contains("Sheet1!A1"), "{e}");
    // The same fit without statistics, and a line that is exact in the numbers, are answered.
    assert!(build_with("INDEX(LOGEST({1;2;4;8},{1;2;3;4}),1,1)").is_ok());
    assert!(build_with("INDEX(LINEST({0;1;2;3},{1;2;3;4},TRUE,TRUE),4,1)").is_ok());
    let e = build_with("INDEX(FREQUENCY({5;3;8;1},{3;3;6}),1,1)").unwrap_err();
    assert!(e.contains("FREQUENCY") && e.contains("equal bins"), "{e}");
}

#[test]
fn tbill_yield_over_the_last_days_of_february_is_not_reproduced() {
    let e = build_with("TBILLYIELD(DATE(2024,2,29),DATE(2025,2,28),98)").unwrap_err();
    assert!(e.contains("TBILLYIELD") && e.contains("February"), "{e}");
    // Past a year it is an error as in LibreOffice, and other dates are answered.
    assert!(build_with("TBILLYIELD(DATE(2023,2,28),DATE(2025,2,28),98)").is_ok());
    assert!(build_with("TBILLYIELD(DATE(2024,5,15),DATE(2024,11,15),98)").is_ok());
}

#[test]
fn text_functions_of_libreoffice_outside_what_the_engine_reproduces_stop_the_build() {
    // A regular expression outside the reproduced subset (it can match the empty text, or is
    // anchored): the same refusal as the search of the criteria functions.
    for f in [
        "_xlfn.ORG.LIBREOFFICE.REGEX(\"abc\",\"a*\",\"X\")",
        "_xlfn.ORG.LIBREOFFICE.REGEX(\"abc\",\"^a\",\"Z\")",
        "_xlfn.ORG.LIBREOFFICE.REGEX(\"abc\",\"(?i)B\")",
    ] {
        let e = build_with(f).unwrap_err();
        assert!(
            e.contains("regular expression") && e.contains("Sheet1!A1"),
            "{f}: {e}"
        );
    }
    assert!(build_with("_xlfn.ORG.LIBREOFFICE.REGEX(\"abc\",\"(b)\",\"[$1]\")").is_ok());
    // A cut through a character above the plane gives half of a surrogate pair there.
    let e = build_with("RIGHTB(\"😀\",1)").unwrap_err();
    assert!(e.contains("RIGHTB") && e.contains("surrogate"), "{e}");
    assert!(build_with("LEFTB(\"😀\",2)").is_ok());
}

#[test]
fn a_text_code_that_is_not_known_at_build_time_stops_the_build() {
    let mut m = sheetrs_compiler::Model::new();
    let s = m.add_sheet("Sheet1");
    m.set(s, "B1", Value::text("dd/mm/yyyy"));
    m.set(s, "B2", Value::text("[$-407]dddd"));
    m.set_formula(s, "A1", "TEXT(45367,B1)");
    // The code is a constant: it is checked.
    assert!(compile(&m, &spec(&[], &["A1"])).is_ok());
    m.set_formula(s, "A1", "TEXT(45367,B2)");
    let e = compile(&m, &spec(&[], &["A1"])).unwrap_err().0;
    assert!(e.contains("TEXT code"), "{e}");
    // The code is an input: the build cannot tell, so it warns, and a run that gives a code with
    // such a tag gets the error value of what is not implemented.
    m.set_formula(s, "A1", "TEXT(45367,B1)");
    let (p, r) = compile(&m, &spec(&["B1"], &["A1"])).unwrap();
    assert!(
        r.warnings
            .iter()
            .any(|w| w.contains("Sheet1!A1") && w.contains("TEXT")),
        "{:?}",
        r.warnings
    );
    let run = |code: &str| {
        p.execute(
            &json!({"Sheet1!B1": code}),
            &sheetrs_compiler::Overrides::default(),
        )
        .unwrap()["Sheet1!A1"]
            .clone()
    };
    assert_eq!(run("yyyy"), json!("2024"));
    assert_eq!(run("[$-407]dddd"), err("#NOT_IMPLEMENTED!"));
}

#[test]
fn a_regular_expression_outside_the_subset_stops_the_build() {
    use sheetrs_compiler::options::{Search, SearchMode};
    let regex = Search {
        mode: SearchMode::Regex,
        whole_cell: true,
        case_sensitive: false,
    };
    let build = |f: &str, inputs: &[&str]| {
        let mut m = sheetrs_compiler::Model::new();
        let s = m.add_sheet("Sheet1");
        m.set(s, "B1", Value::text("abc"));
        m.set(s, "B2", Value::text("a1"));
        m.set(s, "C1", Value::text("a*"));
        m.set_formula(s, "A1", f);
        let sp = Spec {
            search: Some(regex),
            ..spec(inputs, &["A1"])
        };
        compile(&m, &sp).map(|_| ()).map_err(|e| e.0)
    };
    for f in [
        "COUNTIF(B1:B2,\"a*\")",
        "COUNTIF(B1:B2,\"^a\")",
        "SEARCH(\"\\bfoo\",B1)",
        "SEARCH(\"(?=a)a\",B1)",
        "SEARCH(\"\\p{L}+\",B1)",
        "MATCH(\"(a)\\1\",B1:B2,0)",
        "COUNTIF(B1:B2,C1)",
    ] {
        let e = build(f, &[]).expect_err(f);
        assert!(
            e.contains("Sheet1!A1") && e.contains("regular expression"),
            "{f}: {e}"
        );
    }
    for ok in [
        "COUNTIF(B1:B2,\"a.+\")",
        "COUNTIF(B1:B2,\"[a-c]+\")",
        "SEARCH(\"(b|c)+\",B1)",
        "COUNTIF(B1:B2,\">=a\")",
        "COUNTIF(B1:B2,5)",
        "SEARCH(\"b\",B1)",
    ] {
        assert!(build(ok, &[]).is_ok(), "{ok}");
    }
    // A pattern from an input cell cannot be checked: a warning, not an error.
    assert!(build("COUNTIF(B1:B2,C1)", &["C1"]).is_ok());
    // Without the regular expression mode, the same text is a wildcard pattern.
    let mut m = sheetrs_compiler::Model::new();
    let s = m.add_sheet("Sheet1");
    m.set(s, "B1", Value::text("abc"));
    m.set_formula(s, "A1", "COUNTIF(B1:B1,\"a*\")");
    assert!(compile(&m, &spec(&[], &["A1"])).is_ok());
}

#[test]
fn what_the_build_cannot_tell_warns_and_gives_an_error_value_when_met() {
    let mut m = sheetrs_compiler::Model::new();
    let s = m.add_sheet("Sheet1");
    for (i, v) in [1.0, 2.0, 3.0].iter().enumerate() {
        m.set(s, &format!("B{}", i + 1), Value::Number(*v));
        m.set(s, &format!("C{}", i + 1), Value::Number(10.0 * v));
    }
    m.set(s, "D1", Value::Number(1.0));
    m.set_formula(s, "A1", "XLOOKUP(2,B1:B3,C1:C3,\"NF\",0,D1)");
    let (p, r) = compile(&m, &spec(&["D1"], &["A1"])).unwrap();
    assert!(
        r.warnings.iter().any(|w| w.contains("XLOOKUP")),
        "{:?}",
        r.warnings
    );
    let run = |mode: i64| {
        p.execute(
            &json!({"Sheet1!D1": mode}),
            &sheetrs_compiler::Overrides::default(),
        )
        .unwrap()["Sheet1!A1"]
            .clone()
    };
    assert_eq!(run(1), json!(20));
    assert_eq!(run(2), err("#NOT_IMPLEMENTED!"));
    assert_eq!(run(-2), err("#NOT_IMPLEMENTED!"));
    // Known at build time, it is an error: a constant cell as the mode.
    m.set(s, "D1", Value::Number(2.0));
    assert!(compile(&m, &spec(&[], &["A1"])).is_err());
}

#[test]
fn dates_and_patterns_that_only_the_inputs_make_unsupported_give_the_error_value() {
    use sheetrs_compiler::options::{Search, SearchMode};
    let mut m = sheetrs_compiler::Model::new();
    let s = m.add_sheet("Sheet1");
    m.set(s, "B1", Value::Number(45367.0));
    m.set(s, "C1", Value::text("abc"));
    m.set(s, "D1", Value::text("b+"));
    m.set_formula(s, "A1", "YEAR(B1)");
    m.set_formula(s, "A2", "SEARCH(D1,C1)");
    let sp = Spec {
        search: Some(Search {
            mode: SearchMode::Regex,
            whole_cell: true,
            case_sensitive: false,
        }),
        ..spec(&["B1", "D1"], &["A1", "A2"])
    };
    let (p, r) = compile(&m, &sp).unwrap();
    assert!(
        r.warnings.iter().any(|w| w.contains("SEARCH")),
        "{:?}",
        r.warnings
    );
    let run = |b: f64, d: &str| {
        p.execute(
            &json!({"Sheet1!B1": b, "Sheet1!D1": d}),
            &sheetrs_compiler::Overrides::default(),
        )
        .unwrap()
    };
    let ok = run(45367.0, "b+");
    assert_eq!(
        (ok["Sheet1!A1"].clone(), ok["Sheet1!A2"].clone()),
        (json!(2024), json!(2))
    );
    let bad = run(-9_000_000.0, "^b");
    assert_eq!(bad["Sheet1!A1"], err("#NOT_IMPLEMENTED!"));
    assert_eq!(bad["Sheet1!A2"], err("#NOT_IMPLEMENTED!"));
}

#[test]
fn fourier_by_rows_of_an_array_written_in_the_formula_stops_the_build() {
    // LibreOffice transposes that array in the formula each time it calculates.
    let e = build_with("SUM(FOURIER({1,2;3,4},FALSE))").unwrap_err();
    assert!(e.contains("FOURIER") && e.contains("transposes"), "{e}");
    // By columns, a computed array and a range are as in LibreOffice.
    for ok in [
        "SUM(FOURIER({1,2;3,4},TRUE))",
        "SUM(FOURIER(TRANSPOSE({1,2;3,4}),FALSE))",
    ] {
        assert!(build_with(ok).is_ok(), "{ok}");
    }
}

#[test]
fn forecast_ets_that_is_random_or_reads_beyond_its_arrays_stops_the_build() {
    // A prediction interval with seasons is computed from random scenarios.
    let e = build_with("FORECAST.ETS.CONFINT(9,{1;2;3;4;5;6;7;8},{1;2;3;4;5;6;7;8},0.95,2)")
        .unwrap_err();
    assert!(e.contains("FORECAST.ETS") && e.contains("random"), "{e}");
    // A target between the last timeline value and the next step reads past a vector.
    let e = build_with("FORECAST.ETS.CONFINT(4.5,{1;2;3;5},{1;2;3;4},0.95,0)").unwrap_err();
    assert!(e.contains("beyond"), "{e}");
    // Without seasons and on a step the interval is computed.
    assert!(build_with("FORECAST.ETS.CONFINT(6,{1;2;3;5},{1;2;3;4},0.95,0)").is_ok());
}

#[test]
fn cell_and_info_of_the_document_or_the_system_stop_the_build() {
    for f in [
        "CELL(\"format\",B1)",
        "CELL(\"width\",B1)",
        "CELL(\"protect\",B1)",
    ] {
        let e = build_with(f).unwrap_err();
        assert!(e.contains(&f[..4]), "{f}: {e}");
    }
    // What a cell holds, where it is and what a program knows of itself are answered.
    for ok in [
        "CELL(\"address\",B3)",
        "CELL(\"type\",B3)",
        "INFO(\"recalc\")",
        "INFO(\"numfile\")",
    ] {
        assert!(build_with(ok).is_ok(), "{ok}");
    }
    assert_eq!(calc("CELL(\"address\",B3)"), json!("$B$3"));
}
