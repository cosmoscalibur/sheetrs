use crate::common::*;
use serde_json::{Value as Json, json};

#[test]
fn arithmetic_and_precedence() {
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
fn lazy_forms() {
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
    assert_eq!(calc("SQRT(-1)"), err("#VALUE!")); // Err:502
    assert_eq!(calc("CEILING(2.1,0.5)"), json!(2.5));
    assert_eq!(calc("SUMPRODUCT({1,2,3},{4,5,6})"), json!(32));
    assert_eq!(calc("SUM({1,2;3,4})"), json!(10));
    assert_eq!(calc("MEDIAN(1,3,2,10)"), json!(2.5));
    assert_eq!(calc("LARGE({5,1,9},2)"), json!(5));
    assert_eq!(calc("PMT(0.01,12,-1000)"), json!(88.84878867834166));
}

#[test]
fn dates_reach_the_year_32767_like_libreoffice() {
    assert_eq!(calc("DATE(10000,1,1)"), json!(2958466));
    assert_eq!(calc("DATE(12345,1,1)"), json!(3814960));
    assert_eq!(calc("DATE(32767,12,31)"), json!(11274306));
    assert_eq!(calc("YEAR(3815019)"), json!(12345));
    assert_eq!(calc("TEXT(3815019,\"yyyy-mm-dd\")"), json!("12345-03-01"));
    assert_eq!(calc("DATEVALUE(\"3/1/12345\")"), json!(3815019));
    assert_eq!(calc("DATEVALUE(\"3/1/32768\")"), err("#VALUE!"));
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
    let out = p.execute(&Json::Null, &Default::default()).unwrap();
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
    let out = p.execute(&Json::Null, &Default::default()).unwrap();
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
fn lookups() {
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
    let out = p.execute(&Json::Null, &Default::default()).unwrap();
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
    s.date_system = Some(sheetrs_compiler::options::DateSystem::NULL_1904);
    let mut m = sheetrs_compiler::Model::new();
    let sh = m.add_sheet("Sheet1");
    m.set_formula(sh, "A1", "DATE(2024,1,1)");
    let (p, _) = build(&m, &s);
    assert_eq!(
        p.execute(&Json::Null, &Default::default()).unwrap()["Sheet1!A1"],
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
        p.execute(&Json::Null, &Default::default()).unwrap()["Sheet1!A1"],
        err("#NAME?")
    );
    assert!(
        p.diagnostics
            .iter()
            .any(|d| d.contains("Sheet1!A1") && d.contains("NOSUCHFN"))
    );
}

/// Approximate lookups over an array that is not a range follow LibreOffice's own search for
/// them; each answer here is one LibreOffice 26.8 gave (15,000 more were compared when it was
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
        ("MATCH(3,{1,2;3,4},1)", err("#VALUE!")),
        ("MATCH(3,{1,2,3},2)", err("#VALUE!")),
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
/// search in halves; LibreOffice 26.8 gave these (its walk down skips the last element).
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
        .execute(
            &serde_json::Value::Null,
            &sheetrs_compiler::Overrides::default(),
        )
        .unwrap();
    assert_eq!(r["Sheet1!A1"], err("#VALUE!"));
    assert_eq!(r["Sheet1!A2"], err("#VALUE!"));
    // A square range is searched down its first column.
    assert_eq!(r["Sheet1!A3"], json!(2));
    assert_eq!(r["Sheet1!A4"], json!(2));
}

/// Text of every script orders as LibreOffice orders it (ICU's collation): each answer is one
/// LibreOffice 26.8 gave.
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
