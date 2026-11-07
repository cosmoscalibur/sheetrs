//! Find & Replace against `LibreOffice` itself.
//!
//! `tests/assets/replace/` holds a workbook made with `LibreOffice` 24.2.7 in both formats
//! (`input.*`) and, for each case of `cases.json`, the workbook `LibreOffice` leaves after running
//! `replaceAll` with the same search (`expected_<case>.*`, saved by `LibreOffice` in the same format).
//! The editor runs each case on the input and the two workbooks are read back and compared cell by
//! cell: the value, and the formula when there is one (the cached result of a formula that was
//! rewritten is not compared: the editor drops it and the application computes it).

use serde_json::Value as Json;
use sheetrs_edit::Editor;
use sheetrs_io::reader::{CellValue, Source, Workbook, read_workbook};
use sheetrs_protocol::{Find, Mode, Reference, Replace, Scope, SearchIn};
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

fn asset(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/replace")
        .join(name)
}

type Cells = BTreeMap<(String, u32, u32), (CellValue, Option<String>)>;

fn cells(workbook: &Workbook) -> Cells {
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
                (value, cell.formula.as_deref().map(str::to_string)),
            );
        }
    }
    out
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
            _ => SearchIn::Value,
        },
        include_formulas: case["include_formulas"].as_bool().unwrap_or(false),
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

/// What `LibreOffice` does and the editor does not, on purpose: a search by value reaches a formula
/// cell in `LibreOffice` whatever it is told, and in the editor only with `include_formulas`.
const LIBREOFFICE_REACHES_FORMULAS: &[&str] = &["value_excludes_formulas"];

fn run(extension: &str) {
    let dir = tempfile::tempdir().unwrap();
    for case in cases() {
        let name = case["name"].as_str().unwrap();
        let source = Source::from_file(asset(&format!("input.{extension}"))).unwrap();
        let (scopes, replace) = spec(&case);
        let out = dir.path().join(format!("{name}.{extension}"));
        let mut editor = Editor::new(&source);
        editor.replace(scopes, replace);
        let report = editor
            .save(&out)
            .unwrap_or_else(|e| panic!("{extension} {name}: {e}"));
        assert!(
            report.skipped.is_empty(),
            "{extension} {name}: {:?}",
            report.skipped
        );
        let got = cells(&read_workbook(&out).unwrap());
        if LIBREOFFICE_REACHES_FORMULAS.contains(&name) {
            let input =
                cells(&read_workbook(asset(&format!("input.{extension}")).as_path()).unwrap());
            assert_eq!(got.len(), input.len(), "{extension} {name}");
            assert_eq!(report.replaced, 0, "{extension} {name}");
            continue;
        }
        let want = cells(&read_workbook(asset(&format!("expected_{name}.{extension}"))).unwrap());
        let mut different = Vec::new();
        for key in got
            .keys()
            .chain(want.keys())
            .collect::<std::collections::BTreeSet<_>>()
        {
            let (g, w) = (got.get(key), want.get(key));
            let equal = match (g, w) {
                (Some((gv, gf)), Some((wv, wf))) => gf == wf && (gf.is_some() || same(gv, wv)),
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

fn input(extension: &str) -> Source {
    Source::from_file(asset(&format!("input.{extension}"))).unwrap()
}

fn formula_search(text: &str, with: &str) -> Replace {
    Replace {
        find: Find {
            search_in: SearchIn::Formula,
            ..Find::literal(text)
        },
        with: with.to_string(),
    }
}

#[test]
fn a_replacement_that_leaves_no_formula_is_reported_and_the_cell_is_left_as_it_was() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut editor = Editor::new(&input(extension));
        editor.replace([], formula_search("B1", "(("));
        let report = editor.save(&out).unwrap();
        // LibreOffice keeps `=((*2` as a formula with an error; the editor does not apply it.
        let skipped: Vec<_> = report.skipped.iter().map(|s| s.cell.as_str()).collect();
        assert_eq!(skipped, ["A6", "A10", "A12"], "{extension}");
        let before = cells(&read_workbook(asset(&format!("input.{extension}")).as_path()).unwrap());
        let after = cells(&read_workbook(&out).unwrap());
        for at in ["A6", "A10", "A12"] {
            let row = at[1..].parse::<u32>().unwrap() - 1;
            let key = ("T".to_string(), row, 0);
            assert_eq!(after[&key].1, before[&key].1, "{extension} {at}");
        }
        // The constant that holds `b1` is text and does change.
        assert_eq!(report.replaced, 1, "{extension}");
        assert_eq!(
            after[&("T".to_string(), 10, 0)].0,
            CellValue::Text("a1 (( A1".into()),
            "{extension}"
        );
    }
}

#[test]
fn a_formula_that_names_a_sheet_that_does_not_exist_is_reported() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let mut editor = Editor::new(&input(extension));
        editor.replace([], formula_search("U Sheet", "Nope"));
        let report = editor
            .save(dir.path().join(format!("out.{extension}")))
            .unwrap();
        assert_eq!(report.skipped.len(), 1, "{extension}");
        assert!(report.skipped[0].reason.contains("`Nope`"), "{extension}");
    }
}

#[test]
fn a_scope_that_names_a_missing_sheet_and_a_pattern_that_is_not_reproduced_are_errors() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut editor = Editor::new(&input(extension));
        editor.replace(
            [Scope::Sheet("Nope".into())],
            Replace {
                find: Find::literal("a"),
                with: "b".into(),
            },
        );
        assert!(matches!(editor.save(&out), Err(e) if e.to_string().contains("Nope")));
        assert!(!out.exists());
        let mut editor = Editor::new(&input(extension));
        editor.replace(
            [],
            Replace {
                find: Find {
                    mode: Mode::Regex,
                    ..Find::literal("^a")
                },
                with: "b".into(),
            },
        );
        assert!(
            editor
                .save(&out)
                .unwrap_err()
                .to_string()
                .contains("unsupported search")
        );
    }
}

#[test]
fn a_replacement_reads_the_workbook_as_the_changes_before_it_leave_it() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut editor = Editor::new(&input(extension));
        editor
            .delete_rows(["T!1:2".parse().unwrap()])
            .replace(
                [Scope::Range(Reference::parse("T!A1").unwrap())],
                Replace {
                    find: Find::literal("1"),
                    with: "one".into(),
                },
            )
            .delete_rows(["T!1".parse().unwrap()]);
        let plan = editor.explain().unwrap();
        assert_eq!(plan.matches("pass ").count(), 3, "{extension}: {plan}");
        assert!(plan.contains("replace \"1\" with \"one\""), "{plan}");
        assert!(plan.contains("recalculate when opened"), "{plan}");
        editor.save(&out).unwrap();
        // Row 3 of the input ("1234.5") is row 1 after the first deletion; the last deletion
        // removes it again, so no `one` is left, and the cell below is the one of the input's row 4.
        let after = cells(&read_workbook(&out).unwrap());
        assert!(
            !after
                .values()
                .any(|(v, _)| matches!(v, CellValue::Text(t) if t.contains("one"))),
            "{extension}"
        );
    }
}

#[test]
fn what_the_replacement_changes_is_what_the_report_counts() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut editor = Editor::new(&input(extension));
        editor.replace(
            [],
            Replace {
                find: Find::literal("hello"),
                with: "bye".into(),
            },
        );
        let report = editor.save(&out).unwrap();
        // `Hello World hello`, `Hello U` (sheet U).
        assert_eq!(
            (report.replaced, report.skipped.len()),
            (2, 0),
            "{extension}"
        );
    }
}

#[test]
fn the_command_replaces_in_place_and_says_what_it_did() {
    for extension in ["xlsx", "ods"] {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join(format!("book.{extension}"));
        std::fs::copy(asset(&format!("input.{extension}")), &path).unwrap();
        let run = |args: &[&str]| {
            std::process::Command::new(env!("CARGO_BIN_EXE_sheetedit"))
                .arg(&path)
                .args(args)
                .output()
                .unwrap()
        };
        let dry = run(&[
            "--replace",
            "hello",
            "--with",
            "bye",
            "--scope",
            "U Sheet",
            "--dry-run",
            "--in-place",
        ]);
        let text = String::from_utf8_lossy(&dry.stdout).to_string();
        assert!(
            text.contains("replace \"hello\" with \"bye\" (literal, in value) in sheet U Sheet"),
            "{text}"
        );
        let done = run(&["--replace", "hello", "--with", "bye", "--in-place"]);
        let text = String::from_utf8_lossy(&done.stdout).to_string();
        assert!(
            done.status.success(),
            "{text}{}",
            String::from_utf8_lossy(&done.stderr)
        );
        assert!(text.contains("Replaced in 2 cells"), "{text}");
        let after = cells(&read_workbook(&path).unwrap());
        assert_eq!(
            after[&("T".to_string(), 0, 0)].0,
            CellValue::Text("bye World bye".into()),
            "{extension}"
        );
        // `--with` without `--replace` is a usage error.
        assert!(!run(&["--with", "x", "--in-place"]).status.success());
    }
}
