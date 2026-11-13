//! Find & Replace against `LibreOffice` itself.
//!
//! `tests/assets/replace/` holds a workbook made with `LibreOffice` 26.8 in both formats
//! (`input.*`) and, for each case of `cases.json`, the workbook `LibreOffice` leaves after running
//! `replaceAll` with the same search (`expected_<case>.*`, saved by `LibreOffice` in the same format).
//! The editor runs each case on the input and the two workbooks are read back and compared cell by
//! cell: the value, and the formula when there is one (the cached result of a formula that was
//! rewritten is not compared: the editor drops it and the application computes it).

use serde_json::Value as Json;
use sheetrs_edit::BookEdit;
use sheetrs_io::Book;
use sheetrs_io::reader::{CellValue, Workbook};
use sheetrs_protocol::{Find, Mode, Reference, Replace, Scope, SearchIn};
use sheetrs_query::BookQuery;
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/replace")
        .join(name)
}

/// The value, the formula, and the number format when the case compares it.
type Cells = BTreeMap<(String, u32, u32), (CellValue, Option<String>, Option<String>)>;

/// A number format as the two files write it: `LibreOffice` saves `mm/dd/yy` in one and
/// `MM/DD/YY` in the other, and puts the language of the format before some.
fn same_format(format: Option<&str>) -> String {
    let lower = format.unwrap_or("General").to_lowercase();
    lower
        .replace("[$-409]", "")
        .replace(['\\', '"'], "")
        .replace("[red]", "")
}

fn cells(workbook: &Workbook, formats: bool) -> Cells {
    let mut out = BTreeMap::new();
    for sheet in &workbook.sheets {
        for ((row, col), cell) in sheet.cells.iter() {
            if matches!(cell.value, CellValue::Empty) && cell.formula.is_none() {
                continue;
            }
            // A boolean `LibreOffice` writes with a format is the number 1 in one file and `TRUE`
            // in the other.
            let value = match &cell.value {
                CellValue::Boolean(b) => CellValue::Number(f64::from(u8::from(*b))),
                v => v.clone(),
            };
            out.insert(
                (sheet.name.clone(), row, col),
                (
                    value,
                    cell.formula.as_deref().map(str::to_string),
                    formats.then(|| same_format(cell.num_fmt.as_deref())),
                ),
            );
        }
    }
    out
}

/// The text of every note of a workbook, by sheet and cell, as `sheetrs query` reads it.
fn notes(path: &Path) -> BTreeMap<(String, String), String> {
    let query: sheetrs_protocol::Query =
        serde_json::from_value(serde_json::json!({"properties": ["comment"]})).unwrap();
    Book::from_file(path)
        .unwrap()
        .query(&query)
        .unwrap()
        .cells
        .into_iter()
        .filter_map(|c| {
            c.comment
                .clone()
                .map(|n| ((c.sheet.to_string(), c.cell()), n))
        })
        .collect()
}

fn same(a: &CellValue, b: &CellValue) -> bool {
    match (a, b) {
        (CellValue::Number(x), CellValue::Number(y)) => (x - y).abs() < 1e-9,
        (a, b) => format!("{a:?}") == format!("{b:?}"),
    }
}

fn spec(case: &Json) -> (Vec<Scope>, Replace) {
    let text = case["text"].as_str().unwrap().to_string();
    let find = Find {
        text,
        mode: match case["mode"].as_str() {
            Some("regex") => Mode::Regex,
            Some("wildcards") => Mode::Wildcards,
            _ => Mode::Literal,
        },
        match_case: case["match_case"].as_bool().unwrap_or(false),
        entire_cells: case["entire_cells"].as_bool().unwrap_or(false),
        search_in: match case["search_in"].as_str() {
            Some("formula") => SearchIn::Formula,
            Some("comment") => SearchIn::Comment,
            _ => SearchIn::Value,
        },
        include_formulas: case["include_formulas"].as_bool().unwrap_or(false),
        formatted: case["formatted"].as_bool().unwrap_or(false),
        filtered: case["filtered"].as_bool().unwrap_or(false),
    };
    let scopes = case["scope"]
        .as_array()
        .map(|a| {
            a.iter()
                .map(|s| serde_json::from_value(s.clone()).unwrap())
                .collect()
        })
        .unwrap_or_default();
    (
        scopes,
        Replace {
            find,
            with: case["with_"].as_str().unwrap().to_string(),
        },
    )
}

fn cases() -> Vec<Json> {
    serde_json::from_str(&std::fs::read_to_string(asset("cases.json")).unwrap()).unwrap()
}

fn run(extension: &str) {
    let dir = tempfile::tempdir().unwrap();
    for case in cases() {
        let name = case["name"].as_str().unwrap();
        let stem = case["input"].as_str().unwrap_or("input");
        let source = Book::from_file(asset(&format!("{stem}.{extension}"))).unwrap();
        let (scopes, replace) = spec(&case);
        let out = dir.path().join(format!("{name}.{extension}"));
        let mut editor = source.clone();
        editor.replace(scopes, [replace]);
        let report = editor
            .write(&out)
            .unwrap_or_else(|e| panic!("{extension} {name}: {e}"));
        assert!(
            report.skipped.is_empty(),
            "{extension} {name}: {:?}",
            report.skipped
        );
        let formats = case["formats"].as_bool().unwrap_or(false);
        let got = cells(&Book::from_file(&out).unwrap().workbook().unwrap(), formats);
        let want = cells(
            &Book::from_file(asset(&format!("expected_{name}.{extension}")))
                .unwrap()
                .workbook()
                .unwrap(),
            formats,
        );
        let mut different = Vec::new();
        for key in got
            .keys()
            .chain(want.keys())
            .collect::<std::collections::BTreeSet<_>>()
        {
            let (g, w) = (got.get(key), want.get(key));
            let equal = match (g, w) {
                // A formula's value is not compared (the editor leaves it to be recalculated)
                // unless LibreOffice made it an error, which is what the text typed meant.
                (Some((gv, gf, gn)), Some((wv, wf, wn))) => {
                    gn == wn
                        && gf == wf
                        && (gf.is_some() && !matches!(wv, CellValue::Error(_)) || same(gv, wv))
                }
                _ => false,
            };
            if !equal {
                different.push(format!("{key:?}: got {g:?}, LibreOffice {w:?}"));
            }
        }
        assert!(
            different.is_empty(),
            "{extension} {name}:\n{}",
            different.join("\n")
        );
        // The notes are what `LibreOffice` leaves, replaced or not.
        assert_eq!(
            notes(&out),
            notes(&asset(&format!("expected_{name}.{extension}"))),
            "{extension} {name}: the notes"
        );
    }
}

#[test]
fn an_xlsx_is_replaced_as_libreoffice_replaces_it() {
    run("xlsx");
}

#[test]
fn an_ods_is_replaced_as_libreoffice_replaces_it() {
    run("ods");
}

fn input(extension: &str) -> Book {
    Book::from_file(asset(&format!("input.{extension}"))).unwrap()
}

#[test]
fn a_scope_that_names_a_missing_sheet_and_a_pattern_that_is_not_reproduced_are_errors() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut editor = input(extension).clone();
        editor.replace(
            [Scope::Sheet("Nope".into())],
            [Replace {
                find: Find::literal("a"),
                with: "b".into(),
            }],
        );
        assert!(matches!(editor.write(&out), Err(e) if e.to_string().contains("Nope")));
        assert!(!out.exists());
        let mut editor = input(extension).clone();
        editor.replace(
            [],
            [Replace {
                find: Find {
                    mode: Mode::Regex,
                    ..Find::literal("\\bfoo")
                },
                with: "b".into(),
            }],
        );
        assert!(
            editor
                .write(&out)
                .unwrap_err()
                .to_string()
                .contains("unsupported search")
        );
    }
}

#[test]
fn a_replacement_reads_the_workbook_as_it_is_whatever_order_the_deletions_were_pushed_in() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let mut outputs = Vec::new();
        for replace_first in [true, false] {
            let out = dir.path().join(format!("out{replace_first}.{extension}"));
            let mut editor = input(extension).clone();
            let replacement = |editor: &mut Book| {
                editor.replace(
                    [Scope::Range("T!A3".parse::<Reference>().unwrap())],
                    [Replace {
                        find: Find::literal("1"),
                        with: "one".into(),
                    }],
                );
            };
            if replace_first {
                replacement(&mut editor);
            }
            editor
                .delete_rows(["T!1".parse().unwrap()])
                .delete_rows(["T!5".parse().unwrap()]);
            if !replace_first {
                replacement(&mut editor);
            }
            let plan = editor.explain();
            // The replacement is the first pass, against the workbook as it is; the deletions are
            // one pass after it, in the rows of the workbook.
            assert_eq!(plan.matches("pass ").count() / 2, 1, "{extension}: {plan}");
            assert!(
                plan.find("replace \"1\"").unwrap() < plan.find("delete rows").unwrap(),
                "{plan}"
            );
            assert!(plan.contains("delete rows T!1:1, T!5:5"), "{plan}");
            // What the deletions leave stale is calculated again as the last pass.
            assert!(
                plan.contains("pass 3 of 3:\n  recalculate every formula"),
                "{plan}"
            );
            let report = editor.write(&out).unwrap();
            assert_eq!(
                report.replaced, 1,
                "{extension}: T!A3 (1234.5) was replaced"
            );
            outputs.push(std::fs::read(&out).unwrap());
        }
        assert_eq!(
            outputs[0], outputs[1],
            "{extension}: the order changed the result"
        );
    }
}

#[test]
fn what_the_replacement_changes_is_what_the_report_counts() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut editor = input(extension).clone();
        editor.replace(
            [],
            [Replace {
                find: Find::literal("hello"),
                with: "bye".into(),
            }],
        );
        let report = editor.write(&out).unwrap();
        // `Hello World hello`, `Hello U` (sheet U).
        assert_eq!(
            (report.replaced, report.skipped.len()),
            (2, 0),
            "{extension}"
        );
    }
}

/// A formula that refers to another workbook is searched as `LibreOffice` shows it (the file in a
/// URL, resolved against the document) and keeps its reference when it is replaced: the formula of
/// `Summary!A8` of `tests/assets/book/book.*` is `'file:///folder/ext.xlsx'#$Sheet1.B3`.
#[test]
fn a_formula_with_an_external_reference_is_replaced_and_keeps_the_reference() {
    let dir = tempfile::tempdir().unwrap();
    for extension in ["xlsx", "ods"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/assets/book")
            .join(format!("book.{extension}"));
        let source = Book::from_file(&path).unwrap();
        let find = |text: &str, with: &str| {
            let find = Find {
                text: text.to_string(),
                mode: Mode::Literal,
                match_case: true,
                entire_cells: false,
                search_in: SearchIn::Formula,
                include_formulas: false,
                formatted: false,
                filtered: false,
            };
            (
                vec![serde_json::from_value(serde_json::json!({"range": "Summary!A8"})).unwrap()],
                Replace {
                    find,
                    with: with.to_string(),
                },
            )
        };
        // The cell and the text of the formula as `LibreOffice` shows it.
        let folder = std::fs::canonicalize(path.parent().unwrap()).unwrap();
        let (scopes, replace) = find(&format!("{}/ext", folder.display()), "unused");
        let out = dir.path().join(format!("none.{extension}"));
        let mut editor = source.clone();
        editor.replace(scopes, [replace]);
        // The URL is in the text that is searched.
        assert_eq!(editor.write(&out).unwrap().replaced, 1, "{extension}");
        let (scopes, replace) = find("B3", "B4");
        let out = dir.path().join(format!("b4.{extension}"));
        let mut editor = source.clone();
        editor.replace(scopes, [replace]);
        editor.write(&out).unwrap();
        let workbook = Book::from_file(&out).unwrap().workbook().unwrap();
        let formula = workbook
            .sheet("Summary")
            .and_then(|s| s.cell(7, 0))
            .and_then(|c| c.as_formula().map(str::to_string));
        assert_eq!(formula.as_deref(), Some("[1]Sheet1!B4"), "{extension}");
        assert_eq!(workbook.external_workbooks.len(), 1, "{extension}");
    }
}

fn pair(find: &str, with: &str) -> Replace {
    Replace {
        find: Find::literal(find),
        with: with.into(),
    }
}

/// A sweep (several replacements in one call) is the same plan as one call per replacement: each
/// element is an operation, in order, that sees what the one before it left; so the order matters.
#[test]
fn a_sweep_of_replacements_is_one_operation_per_pair_in_order() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let bytes = |name: &str, build: &dyn Fn(&mut Book)| {
            let out = dir.path().join(format!("{name}.{extension}"));
            let mut editor = input(extension).clone();
            build(&mut editor);
            editor.write(&out).unwrap();
            std::fs::read(out).unwrap()
        };
        let sweep = bytes("sweep", &|e| {
            e.replace([], [pair("hello", "bye"), pair("bye", "gone")]);
        });
        let chained = bytes("chained", &|e| {
            e.replace([], [pair("hello", "bye")]);
            e.replace([], [pair("bye", "gone")]);
        });
        assert_eq!(sweep, chained, "{extension}: a sweep is the chained calls");
        let reversed = bytes("reversed", &|e| {
            e.replace([], [pair("bye", "gone"), pair("hello", "bye")]);
        });
        assert_ne!(
            sweep, reversed,
            "{extension}: the order of the pairs matters"
        );
        let explain = {
            let mut editor = input(extension).clone();
            editor.replace([], [pair("hello", "bye"), pair("bye", "gone")]);
            editor.explain()
        };
        assert_eq!(
            explain.matches("replace").count(),
            2,
            "{extension}: explain shows one line per pair: {explain}"
        );
    }
}
