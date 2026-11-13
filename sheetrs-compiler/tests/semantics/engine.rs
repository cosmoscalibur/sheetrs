use crate::common::*;
use serde_json::{Value as Json, json};

#[test]
fn operators_evaluate_with_spreadsheet_precedence_and_coercion() {
    assert_eq!(calc("1+2*3"), json!(7));
    assert_eq!(calc("-2^2"), json!(4));
    assert_eq!(calc("2^3^2"), json!(64)); // left associative, like spreadsheets
    assert_eq!(calc("10/4"), json!(2.5));
    assert_eq!(calc("50%+1"), json!(1.5));
    assert_eq!(calc("1/0"), err("#DIV/0!"));
    assert_eq!(calc("0^0"), json!(1));
    assert_eq!(calc("\"a\"+1"), err("#VALUE!"));
    assert_eq!(calc("\"3\"+1"), json!(4));
    assert_eq!(calc("TRUE+1"), json!(2));
}

#[test]
fn text_and_comparison() {
    assert_eq!(calc("\"a\"&1+1"), json!("a2"));
    assert_eq!(calc("\"abc\"=\"ABC\""), json!(true));
    assert_eq!(calc("1<\"a\""), json!(true)); // numbers sort before text
    assert_eq!(calc("\"a\"<TRUE"), json!(false)); // TRUE is the number 1
    assert_eq!(calc("LEFT(\"hello\",2)&RIGHT(\"hello\",3)"), json!("hello"));
    assert_eq!(calc("SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)"), json!("a-b+c"));
    assert_eq!(calc("TRIM(\"  a   b \")"), json!("a b"));
    assert_eq!(calc("PROPER(\"hELLO wORLD\")"), json!("Hello World"));
    assert_eq!(calc("TEXT(0.256,\"0.0%\")"), json!("25.6%"));
    assert_eq!(calc("TEXT(1234.5,\"#,##0.00\")"), json!("1,234.50"));
    assert_eq!(calc("FIND(\"b\",\"abc\")"), json!(2));
    assert_eq!(calc("SEARCH(\"B*\",\"abc\")"), json!(2));
    assert_eq!(calc("TEXTJOIN(\",\",TRUE,\"a\",\"\",\"b\")"), json!("a,b"));
}

#[test]
fn conditional_functions_do_not_evaluate_the_branch_they_do_not_take() {
    assert_eq!(calc("IF(TRUE,1,1/0)"), json!(1));
    assert_eq!(calc("IF(FALSE,1)"), json!(false));
    assert_eq!(calc("IFERROR(1/0,\"x\")"), json!("x"));
    assert_eq!(calc("IFERROR(5,1/0)"), json!(5));
    assert_eq!(calc("IFNA(NA(),7)"), json!(7));
    assert_eq!(calc("IFNA(1/0,7)"), err("#DIV/0!"));
    assert_eq!(calc("CHOOSE(2,10,20,1/0)"), json!(20));
    assert_eq!(calc("IFS(1>2,\"a\",2>1,\"b\")"), json!("b"));
    assert_eq!(calc("SWITCH(2,1,\"a\",2,\"b\",\"z\")"), json!("b"));
    assert_eq!(calc("AND(TRUE,1,0)"), json!(false));
    assert_eq!(calc("OR(FALSE,0,2)"), json!(true));
}

#[test]
fn math_functions() {
    assert_eq!(calc("ROUND(1.005,2)"), json!(1.01));
    assert_eq!(calc("ROUND(-2.5,0)"), json!(-3));
    assert_eq!(calc("ROUNDUP(1.21,1)"), json!(1.3));
    assert_eq!(calc("MOD(-1,3)"), json!(2));
    assert_eq!(calc("INT(-1.5)"), json!(-2));
    assert_eq!(calc("SQRT(-1)"), libre("Err:502"));
    assert_eq!(calc("CEILING(2.1,0.5)"), json!(2.5));
    assert_eq!(calc("SUMPRODUCT({1,2,3},{4,5,6})"), json!(32));
    assert_eq!(calc("SUM({1,2;3,4})"), json!(10));
    assert_eq!(calc("MEDIAN(1,3,2,10)"), json!(2.5));
    assert_eq!(calc("LARGE({5,1,9},2)"), json!(5));
    assert_eq!(calc("PMT(0.01,12,-1000)"), json!(88.848_788_678_341_66));
}

#[test]
fn dates_reach_the_year_32767_like_libreoffice() {
    assert_eq!(calc("DATE(10000,1,1)"), json!(2_958_466));
    assert_eq!(calc("DATE(12345,1,1)"), json!(3_814_960));
    assert_eq!(calc("DATE(32767,12,31)"), json!(11_274_306));
    assert_eq!(calc("YEAR(3815019)"), json!(12345));
    assert_eq!(calc("TEXT(3815019,\"yyyy-mm-dd\")"), json!("12345-03-01"));
    assert_eq!(calc("DATEVALUE(\"3/1/12345\")"), json!(3_815_019));
    assert_eq!(calc("DATEVALUE(\"3/1/32768\")"), libre("Err:502"));
}

#[test]
fn one_row_aggregates() {
    let m = sheet(&[
        ("A1", json!(1)),
        ("B1", json!("x")),
        ("D1", json!(2.5)),
        ("A2", json!(1)),
        ("B2", json!("=1/0")),
        ("C2", json!(3)),
        ("A3", json!(4)),
        ("B3", json!("=A3*2")),
        ("A5", json!("=SUM(A1:H1)")),
        ("B5", json!("=SUM(A2:C2)")),
        ("C5", json!("=SUM(A3:C3)")),
        ("D5", json!("=SUM(A9:C9)")),
        ("E5", json!("=AVERAGE(A1:D1)")),
        ("F5", json!("=SUM(A1:D1,A3:C3)")),
    ]);
    let (p, _) = build(&m, &spec(&[], &["A5:F5"]));
    let out = p
        .calculate(
            &Json::Null,
            &sheetrs_engine::calculate::Overrides::default(),
        )
        .unwrap();
    assert_eq!(
        out["Sheet1!A5:F5"][0],
        json!([3.5, err("#DIV/0!"), 12, 0, 1.75, 15.5])
    );
}

#[test]
fn ranges_and_aggregates() {
    let m = sheet(&[
        ("A1", json!(1)),
        ("A2", json!(2)),
        ("A3", json!("x")),
        ("A4", json!(4)),
        ("B1", json!("a")),
        ("B2", json!("b")),
        ("B3", json!("a")),
        ("B4", json!("a")),
        ("C1", json!("=SUM(A1:A4)")),
        ("C2", json!("=AVERAGE(A1:A4)")),
        ("C3", json!("=COUNT(A1:A4)")),
        ("C4", json!("=COUNTA(A1:A4)")),
        ("C5", json!("=SUMIF(B1:B4,\"a\",A1:A4)")),
        ("C6", json!("=COUNTIF(A1:A4,\">1\")")),
        ("C7", json!("=SUMIFS(A1:A4,B1:B4,\"a\",A1:A4,\">=1\")")),
        ("C8", json!("=MAX(A:A)")),
        ("C9", json!("=SUMPRODUCT((B1:B4=\"a\")*(A1:A4>0))")),
        ("C10", json!("=SUM(A1:A3)+\"2\"")),
        ("C11", json!("=COUNTIF(B1:B4,\"?\")")),
        ("C12", json!("=COUNTBLANK(A1:A6)")),
    ]);
    let (p, _) = build(&m, &spec(&[], &["C1:C12"]));
    let out = p
        .calculate(
            &Json::Null,
            &sheetrs_engine::calculate::Overrides::default(),
        )
        .unwrap();
    let col: Vec<Json> = out["Sheet1!C1:C12"]
        .as_array()
        .unwrap()
        .iter()
        .map(|r| r[0].clone())
        .collect();
    assert_eq!(
        col,
        vec![
            json!(7),
            json!(7.0 / 3.0),
            json!(3),
            json!(4),
            json!(5), // 1 + 4 (A3 is text, ignored by SUMIF on numbers; B3="a")
            json!(2),
            json!(5),
            json!(4),
            json!(3),
            json!(5),
            json!(4),
            json!(2),
        ]
    );
}

#[test]
fn vlookup_and_index_match_find_the_row_of_a_key() {
    let m = sheet(&[
        ("A1", json!("k1")),
        ("B1", json!(10)),
        ("A2", json!("k2")),
        ("B2", json!(20)),
        ("A3", json!("k3")),
        ("B3", json!(30)),
        ("D1", json!("=VLOOKUP(\"k2\",A1:B3,2,FALSE)")),
        ("D2", json!("=INDEX(B1:B3,MATCH(\"k3\",A1:A3,0))")),
        ("D3", json!("=XLOOKUP(\"k1\",A1:A3,B1:B3)")),
        ("D4", json!("=VLOOKUP(\"zz\",A1:B3,2,FALSE)")),
        ("D5", json!("=IFERROR(VLOOKUP(\"zz\",A1:B3,2,FALSE),0)")),
        ("D6", json!("=MATCH(25,B1:B3,1)")),
        ("D7", json!("=SUM(INDEX(B1:B3,0))")),
        ("D8", json!("=ROWS(A1:B3)*COLUMNS(A1:B3)")),
    ]);
    let (p, _) = build(&m, &spec(&[], &["D1:D8"]));
    let out = p
        .calculate(
            &Json::Null,
            &sheetrs_engine::calculate::Overrides::default(),
        )
        .unwrap();
    let col: Vec<Json> = out["Sheet1!D1:D8"]
        .as_array()
        .unwrap()
        .iter()
        .map(|r| r[0].clone())
        .collect();
    assert_eq!(
        col,
        vec![
            json!(20),
            json!(30),
            json!(10),
            err("#N/A"),
            json!(0),
            json!(2),
            json!(60),
            json!(6)
        ]
    );
}

#[test]
fn dates_and_date_systems() {
    assert_eq!(calc("DATE(2024,1,31)"), json!(45322));
    assert_eq!(
        calc("YEAR(45292)&\"-\"&MONTH(45292)&\"-\"&DAY(45292)"),
        json!("2024-1-1")
    );
    assert_eq!(calc("EDATE(DATE(2024,1,31),1)"), json!(45351)); // 2024-02-29
    assert_eq!(calc("EOMONTH(DATE(2024,1,15),0)"), json!(45322));
    assert_eq!(calc("WEEKDAY(DATE(2024,1,1))"), json!(2));
    assert_eq!(
        calc("DATEDIF(DATE(2020,1,15),DATE(2024,1,14),\"Y\")"),
        json!(3)
    );
    assert_eq!(calc("DAYS(DATE(2024,3,1),DATE(2024,2,1))"), json!(29));
    assert_eq!(
        calc("TEXT(DATE(2024,12,25),\"dd/mm/yyyy\")"),
        json!("25/12/2024")
    );
    // The real calendar: 1900 has no 29th of February.
    assert_eq!(calc("DATE(1900,3,1)-DATE(1900,2,28)"), json!(1));

    let mut s = spec(&[], &["A1"]);
    s.date_system = Some(sheetrs_engine::options::DateSystem::Null1904);
    let mut m = sheetrs_compiler::Model::new();
    let sh = m.add_sheet("Sheet1");
    m.set_formula(sh, "A1", "DATE(2024,1,1)");
    let (p, _) = build(&m, &s);
    assert_eq!(
        p.calculate(
            &Json::Null,
            &sheetrs_engine::calculate::Overrides::default()
        )
        .unwrap()["Sheet1!A1"],
        json!(43830)
    );
}

#[test]
fn prefixes_name_the_same_functions() {
    assert_eq!(calc("EASTERSUNDAY(2024)"), json!(45382));
    assert_eq!(calc("COM.MICROSOFT.CONCAT(\"a\",\"b\")"), json!("ab"));
    assert_eq!(calc("_xlfn.CONCAT(\"a\",\"b\")"), json!("ab"));
    // The IR names the function with its namespace.
    let mut m = sheetrs_compiler::Model::new();
    let s = m.add_sheet("S");
    m.set_formula(s, "A1", "EASTERSUNDAY(2024)+SUM(1)");
    m.set_formula(s, "B1", "A1");
    let (p, _) = build(
        &m,
        &sheetrs_compiler::Spec {
            fold: false,
            ..spec(&[], &["B1"])
        },
    );
    let json = serde_json::to_string(&p).unwrap();
    assert!(
        json.contains("core::EASTERSUNDAY") && json.contains("core::SUM"),
        "{json}"
    );
}

#[test]
fn unsupported_functions_are_reported() {
    let m = sheet(&[("A1", json!("=NOSUCHFN(1)+1"))]);
    let (p, r) = build(&m, &spec(&[], &["A1"]));
    assert!(r.unsupported.contains("NOSUCHFN"));
    assert_eq!(
        p.calculate(
            &Json::Null,
            &sheetrs_engine::calculate::Overrides::default()
        )
        .unwrap()["Sheet1!A1"],
        err("#NAME?")
    );
    assert!(
        p.diagnostics
            .iter()
            .any(|d| d.contains("Sheet1!A1") && d.contains("NOSUCHFN"))
    );
}

/// Approximate lookups over an array that is not a range follow `LibreOffice`'s own search for
/// them; each answer here is one `LibreOffice` 26.8 gave (15,000 more were compared when it was
/// written).
#[test]
fn lookups_over_arrays_that_are_not_sorted_as_libreoffice() {
    for (f, want) in [
        ("MATCH(5,{9,5,2},1)", json!(2)),
        ("MATCH(4,{4,2,6,8,10},1)", json!(2)),
        ("MATCH(3,{3,2,3,3},1)", json!(4)),
        ("MATCH(3,{1,8},-1)", err("#N/A")),
        ("MATCH(6,{6,9},-1)", json!(1)),
        ("MATCH(3,{2;6;4;8},1)", json!(1)),
        ("MATCH(\"c\",{\"a\",3,\"b\",\"d\"},1)", json!(3)),
        ("MATCH(3,{1,2;3,4},1)", libre("Err:539")),
        ("MATCH(3,{1,2,3},2)", libre("Err:539")),
        ("MATCH(3,{1,2,3,4},0.5)", json!(3)),
        ("LOOKUP(6,{2,6,4,8},{100,101,102,103})", json!(101)),
        ("LOOKUP(4,{6,4,2},{100,101,102})", json!(102)),
        ("LOOKUP(3,{1,2;3,4})", json!(4)),
        ("VLOOKUP(5,{4,\"x\";2,\"y\";6,\"z\"},2,TRUE)", json!("y")),
        ("VLOOKUP(3,{1,\"x\";2,\"y\";4,\"z\"},2,TRUE)", json!("y")),
        ("HLOOKUP(\"b\",{\"a\",\"c\";1,2},2,TRUE)", json!(1)),
    ] {
        assert_eq!(calc(f), want, "{f}");
    }
}

/// `XLOOKUP` over an array that is not a range, in the modes that look for the nearest value or
/// search in halves; `LibreOffice` 26.8 gave these (its walk down skips the last element).
#[test]
fn xlookup_over_arrays_as_libreoffice() {
    for (f, want) in [
        ("XLOOKUP(6,{1},{100},\"NF\",-1,-1)", json!("NF")),
        (
            "XLOOKUP(\"c\",{\"a\",\"c\"},{100,101},\"NF\",-1,-1)",
            json!(100),
        ),
        ("XLOOKUP(1,{6,6},{100,101},\"NF\",1,-1)", json!(100)),
        (
            "XLOOKUP(7,{\"B\",4,5},{100,101,102},\"NF\",1,1)",
            json!(100),
        ),
        (
            "XLOOKUP(TRUE,{1,\"B\",1,1,\"b\"},{100,101,102,103,104},\"NF\",-1,-1)",
            json!(103),
        ),
    ] {
        assert_eq!(calc(f), want, "{f}");
    }
}

#[test]
fn match_over_a_range_wants_a_vector_and_a_known_type() {
    let m = sheet(&[
        ("D1", json!(1)),
        ("D2", json!(2)),
        ("F1", json!(1)),
        ("G1", json!(2)),
        ("F2", json!(3)),
        ("G2", json!(4)),
        ("A1", json!("=MATCH(2,D1:D2,2)")),
        ("A2", json!("=MATCH(3,F1:G2,1)")),
        ("A3", json!("=LOOKUP(1,F1:G2)")),
        ("A4", json!("=MATCH(2,D1:D2,0.5)")),
    ]);
    let (p, _) = build(&m, &spec(&[], &["A1", "A2", "A3", "A4"]));
    let r = p
        .calculate(
            &serde_json::Value::Null,
            &sheetrs_engine::Overrides::default(),
        )
        .unwrap();
    assert_eq!(r["Sheet1!A1"], libre("Err:504"));
    assert_eq!(r["Sheet1!A2"], libre("Err:504"));
    // A square range is searched down its first column.
    assert_eq!(r["Sheet1!A3"], json!(2));
    assert_eq!(r["Sheet1!A4"], json!(2));
}

/// Text of every script orders as `LibreOffice` orders it (ICU's collation): each answer is one
/// `LibreOffice` 26.8 gave.
#[test]
fn text_orders_as_libreoffice_in_every_script() {
    for (a, b, lt) in [
        ("日本", "漢字", true),
        ("ア", "a", false),
        ("a", "あ", true),
        ("😀", "a", true),
        ("a", "😀", false),
        ("😀", "😁", true),
        ("😁", "😀", false),
        ("👍", "❤", false),
        ("مرحبا", "שלום", false),
        ("שלום", "привет", false),
        ("привет", "ελλάδα", false),
        ("ελλάδα", "a", false),
        ("हिन्दी", "한국어", true),
        ("한국어", "日本", true),
        ("é", "f", true),
        ("z", "é", false),
        ("ä", "b", true),
        ("Ä", "a", false),
        ("ß", "ss", false),
        ("ch", "d", true),
        ("ll", "m", true),
        ("😀", "漢", true),
    ] {
        let f = format!("\"{a}\"<\"{b}\"");
        assert_eq!(calc(&f), json!(lt), "{f}");
    }
    // English has no letter ñ after n (Spanish has: the oracle tests cover it).
    assert_eq!(calc("\"nz\"<\"ñ\""), json!(false));
}

/// What `LibreOffice` 26.8 answers for the inputs of functions that have several errors
/// (probed with its UNO API).
#[test]
fn error_codes_per_input_as_libreoffice() {
    for (f, want) in [
        (
            "TEXTJOIN(\",\",TRUE,TEXTSPLIT(\"\",\",\"))",
            libre("Err:504"),
        ),
        (
            "TEXTJOIN({\"-\",\"+\"},TRUE,\"a\",\"b\",\"c\")",
            json!("a-b+c"),
        ),
        ("TEXTBEFORE(\"abc\",\"b\",1,0,0,1/0)", err("#DIV/0!")),
        ("ROWS(5)", libre("Err:504")),
        ("ROWS(TRUE)", libre("Err:504")),
        ("COLUMNS(\"x\")", libre("Err:504")),
        ("ROWS(1/0)", err("#DIV/0!")),
        ("REPT(\"b\",-1)", libre("Err:502")),
        ("REPT(\"b\",268435456)", libre("Err:513")),
        ("LEN(REPT(\"ab\",40000))", json!(80000)),
        ("FACT(171)", err("#VALUE!")),
        ("FACT(-1)", libre("Err:502")),
        ("IMPOWER(\"abc\",\"x\")", err("#VALUE!")),
        ("IMPOWER(\"abc\",2)", libre("Err:502")),
        ("SHEET(1)", libre("Err:504")),
        ("SHEET(\"Nope\")", libre("Err:502")),
        ("SHEET(1/0)", err("#DIV/0!")),
        ("CHISQ.INV(0.5,3401607)", libre("Err:523")),
        ("CHIINV(0.5,1E10)", libre("Err:523")),
        ("CHISQINV(0.5,1E10)", libre("Err:523")),
        ("CHISQ.INV.RT(2,10)", libre("Err:502")),
        ("_xlfn._xlws.FILTER(5,1)", libre("Err:504")),
        ("_xlfn._xlws.FILTER({1;2},{1;\"a\"})", libre("Err:536")),
        ("_xlfn._xlws.FILTER({1;2},{0;0})", libre("Err:533")),
        ("_xlfn._xlws.FILTER({1;2},{1;0;1})", libre("Err:504")),
        ("_xlfn._xlws.FILTER({1;2},{1,2;3,4})", err("#VALUE!")),
        ("IMABS(\"3+4i\",1)", libre("Err:504")),
        ("VDB(2400,300,10,0)", libre("Err:511")),
    ] {
        assert_eq!(calc(f), want, "{f}");
    }
}

/// What `LibreOffice` 26.8 answers where the engine used to give the code of another case (each
/// probed per input, see `scripts/lo_probe.py` and `FIDELITY.md`); the oracle rows of
/// `scripts/error_code_formulas.txt` hold the same and many more.
#[test]
fn error_codes_probed_with_libreoffice() {
    for (f, want) in [
        // An error written in an array is refused when the formula is compiled, wherever it is.
        ("{1;#N/A}", libre("Err:539")),
        ("IF(FALSE,{1;#N/A},3)", libre("Err:539")),
        ("IFERROR(SUM({#N/A}),\"x\")", libre("Err:539")),
        ("SUM({1,2;3,#N/A})", libre("Err:539")),
        // What is not a constant, or not in rows of one width, is refused the same way.
        ("{1;A1}", libre("Err:539")),
        ("{1,2+3}", libre("Err:539")),
        ("{1,SUM(1,2)}", libre("Err:539")),
        ("{1,2;3}", libre("Err:539")),
        ("{1,--2}", libre("Err:539")),
        ("{1,(2)}", libre("Err:539")),
        ("{1,5%}", libre("Err:539")),
        ("IF(FALSE,{1,A1},3)", libre("Err:539")),
        ("SUM({1,-2,+3})", json!(2)),
        ("{1,\"x\"}", json!(1)),
        // A plain value where a range goes.
        ("INDEX(5,1)", err("#REF!")),
        ("_xlfn.XMATCH(2,5)", libre("Err:504")),
        ("MATCH(5,5,0)", libre("Err:504")),
        ("SUMIFS(5,5,5)", libre("Err:504")),
        ("SUMIF(5,\">1\")", libre("Err:504")),
        ("COUNTIF(5,\">1\")", libre("Err:504")),
        // Sets of different sizes.
        ("CORREL({1;2;3},{1;2})", libre("Err:502")),
        ("SUMX2MY2({1;2;3},{1;2})", err("#VALUE!")),
        ("SUMPRODUCT({1;2;3},{1;2})", err("#VALUE!")),
        // FOURIER: too many columns or rows is `Err:502`, a blank or a text a plain `#VALUE!`.
        (
            "_xlfn.ORG.LIBREOFFICE.FOURIER({1,2,3;4,5,6},TRUE)",
            libre("Err:502"),
        ),
        (
            "_xlfn.ORG.LIBREOFFICE.FOURIER({1;\"a\";3;4},TRUE)",
            err("#VALUE!"),
        ),
        // Securities: the checks come before the answer `#VALUE!` of the two it does not calculate.
        (
            "ODDFPRICE(DATE(2024,2,15),DATE(2027,7,1),DATE(2024,1,1),DATE(2024,7,1),0.05,0.08,100,2,0)",
            err("#VALUE!"),
        ),
        (
            "ODDFPRICE(DATE(2024,2,15),DATE(2027,7,1),DATE(2024,1,1),DATE(2024,7,1),-0.01,0.08,100,2,0)",
            libre("Err:502"),
        ),
        (
            "ODDFPRICE(DATE(2024,2,15),DATE(2027,7,1),DATE(2024,1,1),DATE(2024,7,1),0.05,0.08,100,3,0)",
            libre("Err:502"),
        ),
        (
            "ODDFYIELD(DATE(2024,2,15),DATE(2027,7,1),DATE(2024,1,1),DATE(2024,7,1),0.05,0,100,2,0)",
            libre("Err:502"),
        ),
        (
            "COUPNUM(DATE(2024,1,15),DATE(2029,6,30),2,\"x\")",
            libre("Err:502"),
        ),
        ("COUPNUM(DATE(2024,1,15),DATE(2029,6,30),2,\"\")", json!(11)),
        (
            "COUPPCD(DATE(2024,1,15),DATE(2029,6,30),\"x\")",
            err("#VALUE!"),
        ),
        // Where the text to find does not fit from the start, `FINDB` says `Err:502`.
        ("FINDB(\"a\",\"abc\",4)", libre("Err:502")),
        ("FINDB(\"a\",\"abc\",3)", err("#VALUE!")),
        ("SEARCH(\"a\",\"abc\",0)", libre("Err:502")),
        ("FIND(\"a\",\"abc\",0)", err("#VALUE!")),
        // `F.INV(1, …)`: where the search of LibreOffice ends, or does not converge.
        ("F.INV(1,5,8)", json!(5.0 * 2f64.powi(268))),
        ("F.INV(1,1,1)", libre("Err:523")),
        // The number of arguments, from the table probed for every function.
        ("SUM()", json!(0)),
        ("PRODUCT()", json!(0)),
        ("AVERAGE()", err("#DIV/0!")),
        ("ROUND(2.5)", json!(3)),
        ("ABS()", libre("Err:511")),
        ("ABS(1,2)", libre("Err:508")),
        ("ATAN2(1,2,3)", libre("Err:504")),
        ("IMABS()", libre("Err:504")),
        ("BAHTTEXT()", libre("Err:511")),
        ("SUMIFS(A1:A2,A1:A2,\">1\",A1:A2)", libre("Err:511")),
        ("NORMDIST(1,0,1)", json!(0.841_344_746_068_542_9)),
        ("LOGINV(0.5,1)", json!(std::f64::consts::E)),
    ] {
        let got = calc(f);
        // `calc` evaluates in an empty workbook: the cell references above are blank there.
        assert_eq!(got, want, "{f}");
    }
}

#[test]
fn a_range_is_met_by_the_cell_of_the_formula_in_any_sheet() {
    use sheetrs_compiler::Model;
    use sheetrs_engine::Value;
    let mut m = Model::new();
    let s = m.add_sheet("Sheet1");
    let d = m.add_sheet("Data");
    for (i, n) in [5.0, 3.0, 8.0, 1.0].into_iter().enumerate() {
        m.set(d, &format!("A{}", i + 1), Value::Number(n));
        m.set(d, &format!("B{}", i + 1), Value::Number(n * 10.0));
    }
    // A column of formulas over a range of another sheet: each meets its own row.
    for row in 1..=6 {
        m.set_formula(s, &format!("C{row}"), "Data!A1:A4*2");
        m.set_formula(s, &format!("D{row}"), "IFERROR(Data!A1:A3,0)");
    }
    // A block is met by the row and the column; a row of cells by the column; a whole column by the
    // row, where it holds nothing.
    m.set_formula(s, "B2", "Data!A1:B4");
    m.set_formula(s, "B3", "Data!A1:B1");
    m.set_formula(s, "B4", "Data!B1:B4+1");
    m.set_formula(s, "B7", "Data!A:A*2");
    m.set_formula(s, "A4", "Data!B1:B4");
    let (p, _) = build(&m, &spec(&[], &["A4", "B2:B4", "B7", "C1:C6", "D1:D6"]));
    let out = p
        .calculate(
            &Json::Null,
            &sheetrs_engine::calculate::Overrides::default(),
        )
        .unwrap();
    let col = |key: &str| -> Vec<Json> {
        match &out[key] {
            Json::Array(rows) => rows.iter().map(|r| r[0].clone()).collect(),
            one => vec![one.clone()],
        }
    };
    assert_eq!(
        col("Sheet1!C1:C6"),
        vec![
            json!(10),
            json!(6),
            json!(16),
            json!(2),
            err("#VALUE!"),
            err("#VALUE!")
        ]
    );
    // What the formula does not meet is an error that `IFERROR` catches.
    assert_eq!(
        col("Sheet1!D1:D6"),
        vec![json!(5), json!(3), json!(8), json!(0), json!(0), json!(0)]
    );
    // `Data!A1:B4` at B2 is `Data!B2`; `Data!A1:B1` at B3 is `Data!B1` (any row of the formula);
    // `Data!B1:B4+1` at B4 is `Data!B4+1`.
    assert_eq!(col("Sheet1!B2:B4"), vec![json!(30), json!(50), json!(11)]);
    assert_eq!(col("Sheet1!B7"), vec![json!(0)]);
    // A column of one column is met by the row alone: `Data!B1:B4` at A4 is `Data!B4`.
    assert_eq!(col("Sheet1!A4"), vec![json!(10)]);
}

#[test]
fn arguments_are_arrays_where_libreoffice_takes_arrays() {
    let m = sheet(&[
        ("A1", json!(5)),
        ("A2", json!(3)),
        ("A3", json!(8)),
        ("A4", json!(1)),
        ("B1", json!(50)),
        ("B2", json!(30)),
        ("B3", json!(80)),
        ("B4", json!(10)),
    ]);
    let run = |formula: &str| {
        let mut m = m.clone();
        let s = m.sheet_id("Sheet1").unwrap();
        m.set_formula(s, "D1", formula);
        let (p, _) = build(&m, &spec(&[], &["D1"]));
        p.calculate(
            &Json::Null,
            &sheetrs_engine::calculate::Overrides::default(),
        )
        .unwrap()["Sheet1!D1"]
            .clone()
    };
    for (formula, want) in [
        // The arguments `LibreOffice` takes as arrays see the whole range.
        ("CORREL(A1:A4*1,B1:B4)", json!(1)),
        ("SLOPE(B1:B4,A1:A4*1)", json!(10)),
        ("SUMPRODUCT(A1:A4>3)", json!(2)),
        ("SUMPRODUCT((A1:A4>3)*B1:B4)", json!(130)),
        ("_xlfn._xlws.FILTER(A1:A4,A1:A4>3)", json!(5)),
        ("SUM(_xlfn._xlws.FILTER(A1:A4,A1:A4>2))", json!(16)),
        ("_xlfn.XLOOKUP(8,A1:A4*1,B1:B4)", json!(80)),
        ("_xlfn.XMATCH(8,A1:A4*1)", json!(3)),
        // The others meet the cell of the formula, which is in row 1: `A1*2` is 10, a value where
        // an array goes (`Err:504`, `#REF!` for `INDEX`).
        ("INDEX(A1:A4*2,2)", err("#REF!")),
        ("_xlfn._xlws.SORT(A1:A4*2)", libre("Err:504")),
        ("_xlfn.UNIQUE(A1:A4*2)", libre("Err:504")),
        ("_xlfn.TAKE(A1:A4*2,1)", libre("Err:504")),
        ("_xlfn.HSTACK(A1:A4*2,1)", json!(10)),
        ("SUM(A1:A4*2)", json!(10)),
        ("MAX(A1:A4*2)", json!(10)),
        ("IFERROR(A1:A4,0)", json!(5)),
        // A range of several cells where the Analysis add-in takes a number.
        ("DEC2BIN(1,A1:A4)", libre("Err:502")),
        ("DEC2BIN(A1:A4,5)", json!("00101")),
    ] {
        assert_eq!(run(formula), want, "{formula}");
    }
}

#[test]
fn database_functions_take_arrays_written_in_the_formula() {
    let table = "{\"Name\",\"Age\";\"Ann\",30;\"Bob\",25;\"Cid\",30;\"Dee\",41}";
    for (criteria, want) in [
        // A text criterion works as in a range, except that the wildcards of a table written in
        // the formula do not: the whole text of a cell.
        ("{\"Age\";\">26\"}", json!(101)),
        ("{\"Age\";\"=30\"}", json!(60)),
        ("{\"Name\";\"dee\"}", json!(41)),
        ("{\"Name\";\"D*e\"}", json!(0)),
        ("{\"Age\";\">26\";\"<26\"}", json!(126)),
        // A number, a boolean or an empty text in an array is no condition; with no condition at
        // all no record is taken from such a table.
        ("{\"Age\";30}", json!(0)),
        ("{\"Age\";\"\"}", json!(0)),
        ("{\"Age\",\"Name\";30,\"\";\"=25\",\"Bob\"}", json!(25)),
        // A name that is not a header of the table.
        ("{\"Nope\";1}", libre("Err:504")),
    ] {
        let f = format!("DSUM({table},\"Age\",{criteria})");
        assert_eq!(calc(&f), want, "{f}");
    }
    // The field is a position that has to be one of the table.
    assert_eq!(
        calc(&format!("DSUM({table},0,{{\"Age\";\">26\"}})")),
        libre("Err:504")
    );
    assert_eq!(
        calc(&format!("DSUM({table},3,{{\"Age\";\">26\"}})")),
        err("#VALUE!")
    );
    assert_eq!(
        calc(&format!("DGET({table},\"Name\",{{\"Age\";\">40\"}})")),
        json!("Dee")
    );
    // The same criteria over a range: a number there is ignored and means every record.
    let m = sheet(&[
        ("A1", json!("Name")),
        ("B1", json!("Age")),
        ("A2", json!("Ann")),
        ("B2", json!(30)),
        ("A3", json!("Bob")),
        ("B3", json!(25)),
        ("D1", json!("=DSUM(A1:B3,\"Age\",{\"Age\";30})")),
        ("D2", json!("=DSUM(A1:B3,\"Age\",{\"Name\";\"B*\"})")),
    ]);
    let (p, _) = build(&m, &spec(&[], &["D1:D2"]));
    let out = p
        .calculate(
            &Json::Null,
            &sheetrs_engine::calculate::Overrides::default(),
        )
        .unwrap();
    assert_eq!(out["Sheet1!D1:D2"], json!([[55], [25]]));
}

#[test]
fn xlookup_wants_a_row_or_a_column_and_a_result_as_long_as_it() {
    for (formula, want) in [
        ("_xlfn.XLOOKUP(1,{1,2;3,4},{1,2;3,4})", libre("Err:504")),
        (
            "_xlfn.XLOOKUP(9,{1,2;3,4},{1,2;3,4},\"none\")",
            libre("Err:504"),
        ),
        ("_xlfn.XLOOKUP(2,{1;2;3},{10;20})", libre("Err:504")),
        ("_xlfn.XLOOKUP(2,{1;2;3},{10;20;30;40})", libre("Err:504")),
        ("_xlfn.XLOOKUP(2,{1;2;3},7)", libre("Err:504")),
        ("_xlfn.XLOOKUP(2,{1;2;3},{10;20;30})", json!(20)),
    ] {
        assert_eq!(calc(formula), want, "{formula}");
    }
}

/// An array written in a formula may leave elements out, and a boolean in it is a number; the
/// answers are those `LibreOffice` 26.8 gave (`scripts/residual_formulas.txt` holds the rest, which
/// the oracle compares: `LibreOffice` writes an empty element back as `#N/A`, so these cannot go there).
#[test]
fn arrays_with_elements_left_out_and_booleans_as_libreoffice() {
    for (f, want) in [
        ("{}", json!(0)),
        ("{1,2,}", json!(1)),
        ("{,1}", json!(0)),
        ("SUM({1,2,})", json!(3)),
        ("SUM({1,;2,3})", json!(6)),
        ("COLUMNS({1,2,})", json!(3)),
        ("ROWS({1,;2,3})", json!(2)),
        ("COLUMNS({1,;2,3})", json!(2)),
        ("{{1,2}}", libre("Err:533")),
        ("{1,{2}}", libre("Err:533")),
        ("SUM({{1,2}})", libre("Err:533")),
        ("{1,2;}", libre("Err:539")),
        ("{;1,2}", libre("Err:539")),
        ("{TRUE,FALSE}", json!(1)),
        ("SUM({TRUE,FALSE})", json!(1)),
        ("INDEX({TRUE,FALSE},1,2)", json!(0)),
        ("SUMPRODUCT({TRUE,TRUE}*{1,1})", json!(2)),
    ] {
        assert_eq!(calc(f), want, "{f}");
    }
}

/// The counts of arguments `LibreOffice` takes that the engine did not run, and what it answers.
#[test]
fn argument_counts_libreoffice_takes() {
    for (f, want) in [
        ("ROWS()", json!(0)),
        ("COLUMNS()", json!(0)),
        ("ROWS({1;2},{1;2;3})", json!(5)),
        ("COLUMNS({1,2},{1,2,3})", json!(5)),
        ("ROWS({1;2},5)", libre("Err:504")),
        ("CHOOSE(1)", libre("Err:502")),
        ("CHOOSE(\"a\")", err("#VALUE!")),
        ("IF(TRUE)", json!(true)),
        ("IF(0)", json!(false)),
        ("IF(\"a\")", err("#VALUE!")),
        ("IFS()", err("#N/A")),
        ("IFS(TRUE)", libre("Err:511")),
        ("IFS(FALSE)", err("#N/A")),
        ("IFS(FALSE,1,TRUE)", err("#N/A")),
        ("INDEX({1,2;3,4})", json!(1)),
        ("INDEX(5)", err("#REF!")),
        ("_xlfn.TAKE({1,2;3,4})", libre("Err:502")),
        ("_xlfn.DROP(5)", libre("Err:502")),
        ("_xlfn.TEXTBEFORE(\"a-b\")", err("#N/A")),
        ("_xlfn.TEXTBEFORE(\"\")", libre("Err:504")),
        ("_xlfn.TEXTSPLIT(\"a-b\")", json!("a-b")),
        ("_xlfn.TEXTSPLIT(\"\")", libre("Err:504")),
        ("NETWORKDAYS(45292,45299,45293,1)", libre("Err:502")),
        ("NETWORKDAYS(45292,45299,45293,\"a\")", libre("Err:504")),
        ("_xlfn.NORM.S.INV(0.5,1)", err("#VALUE!")),
        ("_xlfn.RANDARRAY(1,1,1,1,TRUE,TRUE)", libre("Err:509")),
    ] {
        assert_eq!(calc(f), want, "{f}");
    }
}

/// Ties, errors among the keys of a lookup, an index outside the table, and the error of the set
/// functions, as `LibreOffice` 26.8 gave them.
#[test]
fn modes_lookups_and_sets_as_libreoffice() {
    for (f, want) in [
        // `MODE` takes the smallest of the tied values, `MODE.SNGL` the first.
        ("MODE({3;3;2;2;1})", json!(2)),
        ("MODE({5;1;1;5;3;3})", json!(1)),
        ("_xlfn.MODE.SNGL({3;3;2;2;1})", json!(3)),
        // An error among the keys that the search reaches is the answer.
        ("HLOOKUP(5,{8,\"a\",5,2}*1,1,FALSE)", err("#VALUE!")),
        ("HLOOKUP(5,{8,5,\"a\",2}*1,1,FALSE)", json!(5)),
        ("VLOOKUP(5,{8;\"a\";5;2}*1,1,FALSE)", json!(5)),
        ("VLOOKUP(5,{5;\"a\"}*1,1,TRUE)", err("#VALUE!")),
        ("VLOOKUP(5,{3;8;\"a\"}*1,1,TRUE)", json!(3)),
        // An index outside the table is an invalid argument.
        ("VLOOKUP(5,{8,5;1,2},3,FALSE)", libre("Err:502")),
        ("HLOOKUP(5,{8,5;1,2},0,FALSE)", libre("Err:502")),
        // The set and the position that are out of range are `#VALUE!`; text for the set is `Err:504`.
        ("LARGE({1;2;3},5)", err("#VALUE!")),
        ("SMALL({1;2;3},0)", err("#VALUE!")),
        ("LARGE(\"a\",1)", libre("Err:504")),
        ("SERIESSUM(1E200,1,1,{1;1;1})", libre("Err:502")),
    ] {
        assert_eq!(calc(f), want, "{f}");
    }
}

/// A reference over several sheets is the same area on each sheet from the first to the last, as
/// `LibreOffice` 26.8 reads it (`scripts/probe_3d.py` asked it of every function): the functions that
/// go through the cells add them up, the others refuse it with their own error, and it is not a value.
#[test]
fn references_over_several_sheets_as_libreoffice() {
    let mut m = sheetrs_compiler::Model::new();
    let data = m.add_sheet("Data");
    let mid = m.add_sheet("Mid");
    let wide = m.add_sheet("Wide");
    let calc = m.add_sheet("Calc");
    for (sheet, values) in [
        (data, [5.0, 3.0, 8.0]),
        (mid, [1.0, 2.0, 3.0]),
        (wide, [8.0, 4.0, 11.0]),
    ] {
        for (i, v) in values.iter().enumerate() {
            m.set(
                sheet,
                &format!("A{}", i + 1),
                sheetrs_engine::Value::Number(*v),
            );
        }
    }
    let cases = [
        // The sheets in between are in it, however it is written.
        ("SUM(Data!A1:Wide!A3)", json!(45)),
        ("SUM(Data:Wide!A1:A3)", json!(45)),
        ("SUM(Wide:Data!A1:A3)", json!(45)),
        ("SUM(Data!A1:Wide!A1)", json!(14)),
        ("SUM(Data!A1:Mid!A1)", json!(6)),
        ("SUM(Data!A1:Data!A3)", json!(16)),
        ("SUM(Data:Wide!A1:A3,1)", json!(46)),
        ("COUNT(Data:Wide!A1:A3)", json!(9)),
        ("AVERAGE(Data!A1:Wide!A1)", json!(14.0 / 3.0)),
        ("MAX(Data:Wide!A1:A3)", json!(11)),
        ("MIN(Data:Wide!A1:A3)", json!(1)),
        ("PRODUCT(Data!A1:Wide!A1)", json!(40)),
        ("LARGE(Data:Wide!A1:A3,2)", json!(8)),
        ("MEDIAN(Data:Wide!A1:A3)", json!(4)),
        // The shape of the reference adds up over the sheets.
        ("ROWS(Data:Wide!A1:A3)", json!(9)),
        ("COLUMNS(Data!A1:Wide!B1)", json!(6)),
        ("_xlfn.SHEETS(Data:Wide!A1)", json!(3)),
        ("ROW(Data:Wide!A5)", json!(5)),
        // What takes one sheet refuses it with its own error, and it is no value.
        ("SUMIF(Data:Wide!A1:A3,\">3\")", libre("Err:504")),
        ("COUNTIF(Data:Wide!A1:A3,\">3\")", libre("Err:504")),
        ("MATCH(5,Data:Wide!A1:A3,0)", libre("Err:504")),
        ("SUMPRODUCT(Data:Wide!A1:A3)", libre("Err:504")),
        ("VLOOKUP(5,Data:Wide!A1:B3,2,FALSE)", libre("Err:504")),
        ("INDEX(Data:Wide!A1:A3,1)", err("#REF!")),
        ("ABS(Data:Wide!A1)", err("#VALUE!")),
        ("Data!A1:Wide!A1", err("#VALUE!")),
        ("Data:Wide!A1+1", err("#VALUE!")),
        ("SUM(Data!A1:Wide!A1+1)", err("#VALUE!")),
        // The reference goes through `IF` and `CHOOSE` to the function that reads their result.
        ("SUM(IF(TRUE,Data:Wide!A1:A3,0))", json!(45)),
        ("SUM(CHOOSE(2,0,Data!A1:Wide!A3))", json!(45)),
        ("ROWS(IF(TRUE,Data:Wide!A1:A3,0))", json!(9)),
        // The first sheet quoted, the last quoted, both, and a column with no row on the last sheet.
        ("SUM('Data':'Wide'!A1:A3)", json!(45)),
        ("SUM('Data':Wide!A1)", json!(14)),
        ("SUM(Data:'Wide'!A1)", json!(14)),
        ("SUM(Data!A2:Wide!A)", json!(9)),
        ("COLUMNS(Data!A2:Wide!B)", json!(6)),
        // A file keeps a range of every row on the sheets as whole columns.
        ("SUM(Data:Wide!A:A)", json!(45)),
    ];
    for (i, (f, _)) in cases.iter().enumerate() {
        m.set_formula(calc, &format!("A{}", i + 1), f);
    }
    let outs: Vec<String> = (1..=cases.len()).map(|i| format!("Calc!A{i}")).collect();
    let outs: Vec<&str> = outs.iter().map(String::as_str).collect();
    let (p, _) = build(&m, &spec(&[], &outs));
    let got = p
        .calculate(
            &Json::Null,
            &sheetrs_engine::calculate::Overrides::default(),
        )
        .unwrap();
    for (i, (f, want)) in cases.iter().enumerate() {
        let got = &got[format!("Calc!A{}", i + 1)];
        if let (Some(a), Some(b)) = (want.as_f64(), got.as_f64()) {
            assert!((a - b).abs() < 1e-9, "{f}: {got}");
        } else {
            assert_eq!(got, want, "{f}");
        }
    }
}

/// A text with an array that is no formula is the error `LibreOffice` gives the cell, and what is left
/// open is closed (the cases of `sheetrs_formula::lax`, run).
#[test]
fn arrays_that_are_no_formula_and_what_is_left_open() {
    for (f, want) in [
        ("{1,2}{3}", libre("Err:509")),
        ("{1,2}}", libre("Err:508")),
        ("{1 2}", libre("Err:539")),
        ("{\"x\",'y'}", libre("Err:539")),
        ("{1,2},3", libre("Err:509")),
        ("{", libre("Err:512")),
        ("{1,2}{", libre("Err:512")),
        ("{1,2}:A1", libre("Err:502")),
        ("{1,2", json!(1)),
        ("SUM({1,2", json!(3)),
        ("SUM({1,2},{3", json!(6)),
        ("{1,2}+{3", json!(4)),
    ] {
        assert_eq!(calc(f), want, "{f}");
    }
}

/// An array written in the formula is taken one element at a time by the function that takes one
/// value there (`LibreOffice` 26.8, `scripts/probe_lift.py`), and a function that takes arrays takes it
/// whole; the elements of an array are numbers.
#[test]
fn arrays_in_the_arguments_of_functions_as_libreoffice() {
    for (f, want) in [
        ("SUM(IF({1;0;1},{5;9;9},{7;7;7}))", json!(21)),
        ("SUM(ABS({-1;-2}))", json!(3)),
        ("SUM(CHOOSE({1;2},10,20))", json!(30)),
        ("INDEX(IF({1;0},{5;9},\"a\"),2)", json!("a")),
        ("IF({1;0})", json!(1)),
        ("SUM(IF({1;0;1},10))", json!(20)),
        ("SUM(IF({TRUE;FALSE;TRUE},1,2))", json!(4)),
        ("SUM(ROUND({1.23;4.56},1))", json!(5.8)),
        ("SUM(LEN({\"ab\";\"cde\"}))", json!(5)),
        ("SUM(SQRT({4;9}))", json!(5)),
        ("SUM(MOD({5;7},3))", json!(3)),
        ("SUM(POWER({2;3},2))", json!(13)),
        ("SUM(IFERROR(1/{1;0},0))", json!(1)),
        ("SUM(INDEX({1;2;3},{1;3}))", json!(4)),
        ("SUM(LOOKUP({2;3},{1;2;3},{10;20;30}))", json!(50)),
        ("SUM(VLOOKUP({1;2},{1,10;2,20},2,FALSE))", json!(30)),
        ("SUM(MATCH({2;3},{1;2;3},0))", json!(5)),
        ("SUM(DATE(2020,{1;2},1))", json!(87693)),
        ("SUM(NOT({TRUE;FALSE}))", json!(1)),
        ("SUM(SIGN({-3;4}))", json!(0)),
        ("SUM(AND({1;0}))", json!(0)),
        // A function that takes an array takes it whole.
        ("MAX(ABS({-1;-5}))", json!(5)),
        ("COUNT(ABS({-1;-2}))", json!(2)),
        ("SUMPRODUCT(ABS({-2;-3}))", json!(5)),
        ("ROWS(ABS({-1;-2}))", json!(2)),
        ("ROWS(SUM({1;2}))", libre("Err:504")),
        // The elements of an array are numbers once they are picked out.
        ("INDEX({TRUE,FALSE}=1,1,1)", json!(1)),
        ("INDEX({TRUE,FALSE}=1,1,2)", json!(0)),
        ("INDEX({1,2}>1,1,2)", json!(1)),
        ("SUM({1,2}>1)", json!(1)),
        ("ISLOGICAL({TRUE,FALSE})", json!(true)),
        ("ISLOGICAL({1,2}>1)", json!(false)),
        ("ISLOGICAL(INDEX({1,2}>1,1,2))", json!(false)),
        ("TYPE({TRUE,FALSE})", json!(64)),
        ("TYPE({1,2}>1)", json!(64)),
        ("INDEX(IF({1,0},TRUE,FALSE),1,1)", json!(1)),
        ("INDEX(NOT({TRUE,FALSE}),1,1)", json!(0)),
    ] {
        assert_eq!(calc(f), want, "{f}");
    }
}

/// The cell that holds what `IF` or `CHOOSE` chose element by element from branches shows booleans
/// when the last branch that ran set a boolean format (`TRUE()`, `NOT(…)`, a comparison), and numbers
/// when it was a function or an operator that gives a number (`SUM(1)`, `DATE(…)`, `5+1`); a
/// constant, a reference or a text leaves what the branch before it set, and a branch no element
/// chooses does not run (`LibreOffice` 26.8, typed and imported). What a function takes of that array
/// is the numbers: `INDEX`, `ISLOGICAL`, `IFERROR`.
#[test]
fn the_boolean_format_of_what_if_chooses_in_an_array_as_libreoffice() {
    for (f, want) in [
        ("IF({1;0},TRUE())", json!(true)),
        ("IF({0;1},TRUE())", json!(false)),
        ("IF({1;0},TRUE(),FALSE())", json!(true)),
        ("IF({0;1;1},TRUE(),FALSE())", json!(false)),
        ("IF({1;0},0,TRUE())", json!(false)),
        // A boolean format on a number that is not 0 or 1: the cell holds the number (`5` in an XLSX,
        // `boolean-value="5"` in an ODS) and shows TRUE.
        ("IF({0;1},TRUE(),5)", json!(5)),
        ("IF({1;0},TRUE(),0)", json!(true)),
        ("IF({1;0},TRUE(),5)", json!(true)),
        ("IF({1;0},NOT(0),5)", json!(true)),
        ("IF({1;0},1=1,5)", json!(true)),
        ("IF({1;0},SUM(1),TRUE())", json!(true)),
        ("IF({0;1},TRUE(),SUM(1))", json!(true)),
        ("IF({1;0},TRUE(),SUM(1))", json!(1)),
        ("IF({1;0},TRUE(),DATE(2020,1,1))", json!(1)),
        ("IF({0;1},TRUE(),DATE(2020,1,1))", json!(43831)),
        // A branch no element chooses does not run: the number stays a number.
        ("IF({1;1},5,TRUE())", json!(5)),
        ("IF({0;0},TRUE(),0)", json!(0)),
        ("CHOOSE({1;2},TRUE(),FALSE())", json!(true)),
        ("CHOOSE({2;1},TRUE(),SUM(1))", json!(true)),
        ("CHOOSE({1;2},TRUE(),SUM(1))", json!(1)),
        // What a function takes of the array is numbers.
        ("INDEX(IF({1;0},TRUE()),1)", json!(1)),
        ("ISLOGICAL(IF({1;0},TRUE()))", json!(false)),
        ("ISNUMBER(IF({1;0},TRUE()))", json!(true)),
        ("SUM(IF({1;0},TRUE()))", json!(1)),
        ("IFERROR(IF({1;0},TRUE()),5)", json!(1)),
        ("IF({1;0},TRUE())+0", json!(1)),
    ] {
        assert_eq!(calc(f), want, "{f}");
    }
}

/// A formula takes the number format of what it ends with: a function that sets none leaves the boolean
/// format of its argument on the number it gives (`VALUE(TRUE())` is the number 1 that its cell shows as
/// `TRUE`, `+TRUE()` too), one that sets a number format gives the number (`ABS(TRUE())`, `SUM(TRUE())`,
/// `TRUE()+0`), a minus keeps the format of an operand that is a boolean (`-TRUE()`) and not that of a
/// number that kept it (`-VALUE(TRUE())`), and `IF`, `CHOOSE` and `IFERROR` give the format of the
/// branch they give (`LibreOffice` 26.8, typed and imported; `scripts/probe_logical.py`).
#[test]
fn a_single_value_keeps_the_boolean_format_as_libreoffice() {
    for (f, want) in [
        ("VALUE(TRUE())", json!(true)),
        ("VALUE(FALSE())", json!(false)),
        ("VALUE(1=1)", json!(true)),
        ("VALUE(NOT(1))", json!(false)),
        ("VALUE(AND(1,1))", json!(true)),
        ("VALUE(ISNUMBER(1))", json!(true)),
        ("VALUE(VALUE(TRUE()))", json!(true)),
        ("VALUE(IF(1,TRUE(),5))", json!(true)),
        ("VALUE(IF(1,5,TRUE()))", json!(5)),
        ("VALUE(2)", json!(2)),
        ("VALUE(1)", json!(1)),
        ("VALUE(0)", json!(0)),
        ("+TRUE()", json!(true)),
        ("+VALUE(TRUE())", json!(true)),
        ("(VALUE(TRUE()))", json!(true)),
        ("-TRUE()", json!(-1)),
        ("-VALUE(TRUE())", json!(-1)),
        ("VALUE(-TRUE())", json!(-1)),
        ("IF(1,VALUE(TRUE()),5)", json!(true)),
        ("IF(0,5,VALUE(TRUE()))", json!(true)),
        ("IF(1,VALUE(TRUE()),VALUE(FALSE()))", json!(true)),
        ("IF(VALUE(TRUE()),0,1)", json!(0)),
        ("IF(TRUE(),0,1)", json!(0)),
        ("IF(TRUE(),1,0)", json!(1)),
        ("CHOOSE(1,VALUE(TRUE()),5)", json!(true)),
        ("CHOOSE(TRUE(),0,1)", json!(0)),
        ("IFERROR(VALUE(TRUE()),5)", json!(true)),
        ("IFERROR(1/0,VALUE(TRUE()))", json!(true)),
        ("IFERROR(0,TRUE())", json!(0)),
        ("IFERROR(TRUE(),0)", json!(true)),
        ("ABS(TRUE())", json!(1)),
        ("ABS(VALUE(TRUE()))", json!(1)),
        ("VALUE(ABS(TRUE()))", json!(1)),
        ("N(TRUE())", json!(1)),
        ("SUM(TRUE())", json!(1)),
        ("SUM(VALUE(TRUE()))", json!(1)),
        ("TRUE()+0", json!(1)),
        ("TRUE()*1", json!(1)),
        ("VALUE(TRUE())+0", json!(1)),
        ("VALUE(TRUE())+1", json!(2)),
        ("VALUE(TRUE())^1", json!(1)),
        ("VALUE(TRUE())+VALUE(5)", json!(6)),
        ("VALUE(TRUE())<1", json!(false)),
        // In an expression it is the number: the format is only the way the cell shows it.
        ("ISLOGICAL(VALUE(TRUE()))", json!(false)),
        ("ISLOGICAL(ABS(TRUE()))", json!(false)),
        ("TYPE(VALUE(TRUE()))", json!(1)),
        ("VALUE(TRUE())&\"\"", json!("1")),
    ] {
        assert_eq!(calc(f), want, "{f}");
    }
}

/// An error written in the formula in the arguments of the functions that read arrays is the array of one
/// element that holds it, and an error element next to a text is left out where `LibreOffice` leaves it out
/// (`LibreOffice` 26.8, imported and typed; `scripts/probe_literal_error.py --strict`).
#[test]
fn an_error_written_where_arrays_are_read_as_libreoffice() {
    for (f, want) in [
        // a text leaves the pair out, whatever the other element is
        (r#"SUMPRODUCT("a",#NAME?)"#, json!(0)),
        (r#"SUMPRODUCT("a",{1}/0)"#, json!(0)),
        (r#"SUMX2MY2("a",{1}/0)"#, json!(0)),
        ("SUMPRODUCT(1,#NAME?)", json!({"error": "#NAME?"})),
        // the last error of the arrays is the answer
        ("SUMPRODUCT(#NAME?,#NUM!)", json!({"error": "#NUM!"})),
        ("SUMPRODUCT(#NUM!,#NAME?)", json!({"error": "#NAME?"})),
        // the sizes are checked first
        ("SUMPRODUCT(#NAME?,{1;2;3})", json!({"error": "#VALUE!"})),
        // an empty cell is 0 and a text is `#VALUE!` in SUMXMY2, an error before it
        ("SUMXMY2(Z1,1)", json!(1)),
        (r#"SUMXMY2("a",1)"#, json!({"error": "#VALUE!"})),
        (r#"SUMXMY2("a",{1}/0)"#, json!({"error": "#DIV/0!"})),
        (r#"COVAR("a",{1}/0)"#, json!({"error": "#VALUE!"})),
        // lists of arrays leave the error out
        ("_xlfn.VSTACK(1,#NAME?)", json!(1)),
        ("_xlfn.HSTACK(1,#NAME?)", json!(1)),
        // a delimiter written as an error is left out of a text, and answered next to a reference
        (r#"_xlfn.TEXTSPLIT("a",#NAME?)"#, json!("a")),
        ("_xlfn.TEXTSPLIT(Z1,#NAME?)", json!({"error": "#NAME?"})),
        // the lanes of CHOOSECOLS are values
        ("_xlfn.CHOOSECOLS(Z1:Z3,#NAME?)", json!({"error": "#NAME?"})),
    ] {
        assert_eq!(calc(f), want, "{f}");
    }
}

/// The element that the missing branch of an `IF` gives in an array is an empty text that `COUNTA`
/// does not count, and the cell that holds the array shows `FALSE` (`LibreOffice` 26.8, typed and
/// imported): `LEN` of it is 0, `TYPE` 2, it equals `""`, `+0` is `#VALUE!`, and `AND` or `OR` of an
/// array with a text in it is `Err:502`; a minus keeps the boolean format of its operand, which is
/// the number -1 (`-TRUE()`), and `TRUE` written without parentheses is the boolean in an XLSX (it is
/// the number 1 where `LibreOffice` types it: it enters the formula as `1`).
#[test]
fn the_missing_branch_of_an_if_in_an_array_as_libreoffice() {
    let value = json!({"error": "#VALUE!"});
    let err502 = json!({"error": "#VALUE!", "code": "Err:502"});
    for (f, want) in [
        ("IF({0;1},1)", json!(false)),
        ("INDEX(IF({0;1},1),1)", json!("")),
        ("LEN(INDEX(IF({0;1},1),1))", json!(0)),
        ("TYPE(INDEX(IF({0;1},1),1))", json!(2)),
        ("INDEX(IF({0;1},1),1)+0", value.clone()),
        ("INDEX(IF({0;1},1),1)*2", value.clone()),
        ("INDEX(IF({0;1},1),1)=\"\"", json!(true)),
        ("INDEX(IF({0;1},1),1)&\"x\"", json!("x")),
        ("ISTEXT(INDEX(IF({0;1},1),1))", json!(true)),
        ("ISNUMBER(INDEX(IF({0;1},1),1))", json!(false)),
        ("ISBLANK(INDEX(IF({0;1},1),1))", json!(false)),
        ("T(INDEX(IF({0;1},1),1))", json!("")),
        ("COUNTA(IF({0;1},1))", json!(1)),
        ("COUNTA({\"\";1})", json!(2)),
        ("COUNTBLANK(IF({1;0},5))", json!(1)),
        ("COUNT(IF({1;0},5))", json!(1)),
        ("SUM(IF({1;0},5))", json!(5)),
        ("AND(IF({1;0},TRUE()))", err502.clone()),
        ("OR(IF({0;0},TRUE()))", err502.clone()),
        ("AND({1;\"a\"})", err502.clone()),
        ("OR({0;\"a\"})", err502),
        ("AND(IF({1;1},TRUE()))", json!(true)),
        ("AND(IF({1;0},TRUE(),FALSE()))", json!(false)),
        ("-TRUE()", json!(-1)),
        ("-IF(1,TRUE())", json!(-1)),
        ("IF({1;0},TRUE,FALSE)", json!(true)),
        ("IF(1,TRUE,FALSE)", json!(true)),
    ] {
        assert_eq!(calc(f), want, "{f}");
    }
}
