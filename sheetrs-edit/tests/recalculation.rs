//! A conversion of formulas to values reads the results the workbook caches. After a change that
//! leaves them stale (a row deleted: the sum over a column reads one row less) they are not the ones
//! the application shows, so the editor calculates them again, with the engine, before it converts.
//!
//! `LibreOffice` calculates as it goes, so the workbooks it leaves after a sequence of steps
//! (`expected_<name>.*`, made by `scripts/gen_recalculation_fixtures.py`) are what the editor has
//! to give when the same steps are saved one at a time (each step reads the file the one before
//! wrote, marked to be recalculated), in each format on its own. The operations of one plan are not
//! steps: they are all written against the workbook as it is, and the planner runs them in an order
//! of its own (the conversion before the deletions), so one plan gives what the steps give in that
//! order.

use sheetrs_edit::BookEdit;
use sheetrs_io::Book;
use sheetrs_io::reader::CellValue;
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

const EXTENSIONS: [&str; 2] = ["xlsx", "ods"];

/// The sequences of `scripts/gen_recalculation_fixtures.py`, as `sheetrs edit` writes the scopes.
const SEQUENCES: &[(&str, &[&str])] = &[
    ("value_only", &["formula_to_value F"]),
    (
        "delete_rows_then_value",
        &["delete_rows T!2", "formula_to_value F"],
    ),
    (
        "delete_everything_then_value",
        &[
            "delete_rows T!2",
            "delete_columns U!B",
            "delete_sheets V",
            "delete_names Gone",
            "formula_to_value F",
        ],
    ),
    (
        "value_in_the_middle",
        &[
            "delete_rows T!2",
            "formula_to_value F!1:8",
            "delete_rows T!1",
        ],
    ),
];

fn structure(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/structure")
        .join(name)
}

/// One operation as a spec: `delete_rows T!2`, `formula_to_value F`.
fn apply(editor: &mut Book, spec: &str) {
    let (verb, rest) = spec.split_once(' ').unwrap();
    match verb {
        "delete_rows" => editor.delete_rows([rest.parse().unwrap()]),
        "delete_columns" => editor.delete_columns([rest.parse().unwrap()]),
        "delete_sheets" => editor.delete_sheets([rest]),
        "delete_names" => editor.delete_names([rest]),
        "formula_to_value" => editor.formula_to_value([rest.parse().unwrap()]),
        other => panic!("unknown operation {other}"),
    };
}

fn editor(source: &Book, specs: &[&str]) -> Book {
    let mut editor = source.clone();
    for spec in specs {
        apply(&mut editor, spec);
    }
    editor
}

type Cells = BTreeMap<(String, u32, u32), (Option<String>, CellValue)>;

/// The formula with each range of one cell (`A1:A1`) as that cell: `LibreOffice` writes the
/// reference of one cell that way once a deletion leaves a range one cell wide.
fn one_cell_ranges(formula: &str) -> String {
    let range = regex::Regex::new(r"(\$?[A-Z]+\$?[0-9]+):(\$?[A-Z]+\$?[0-9]+)").unwrap();
    range
        .replace_all(formula, |c: &regex::Captures| {
            if c[1] == c[2] {
                c[1].to_string()
            } else {
                c[0].to_string()
            }
        })
        .into_owned()
}

/// Every cell with its formula (a reference that went reads `T!#REF!` in an XLSX and `#REF!` in an
/// ODS: both are `#REF!`) and its value.
fn cells(path: &Path) -> Cells {
    let mut cells = BTreeMap::new();
    for sheet in &Book::from_file(path).unwrap().workbook().unwrap().sheets {
        for (at, cell) in sheet.cells.iter() {
            cells.insert(
                (sheet.name.clone(), at.0, at.1),
                (
                    cell.formula
                        .as_deref()
                        .map(|f| one_cell_ranges(&f.replace("T!#REF!", "#REF!"))),
                    cell.value.clone(),
                ),
            );
        }
    }
    cells
}

fn same_value(a: &CellValue, b: &CellValue) -> bool {
    match (a, b) {
        (CellValue::Number(x), CellValue::Number(y)) => (x - y).abs() <= 1e-9 * x.abs().max(1.0),
        _ => a == b,
    }
}

/// What differs between two workbooks, cell by cell. A formula that is still a formula has its text compared and
/// not its cached result: a workbook the plan marks to be recalculated when it is opened holds none
/// that is meant to be read.
fn differences(got: &Cells, want: &Cells) -> Vec<String> {
    got.keys()
        .chain(want.keys())
        .collect::<std::collections::BTreeSet<_>>()
        .into_iter()
        .filter_map(|key| match (got.get(key), want.get(key)) {
            (Some((gf, gv)), Some((wf, wv)))
                if gf == wf && (gf.is_some() || same_value(gv, wv)) =>
            {
                None
            }
            (g, w) => Some(format!("{key:?}: got {g:?}, expected {w:?}")),
        })
        .collect()
}

#[test]
fn a_conversion_after_a_change_gives_what_libreoffice_gives() {
    let mut problems = Vec::new();
    for extension in EXTENSIONS {
        for (name, specs) in SEQUENCES {
            let dir = tempfile::tempdir().unwrap();
            let out = one_at_a_time(extension, specs, dir.path());
            let wanted = structure(&format!("expected_{name}.{extension}"));
            let found = differences(&cells(&out), &cells(&wanted));
            if !found.is_empty() {
                problems.push(format!(
                    "{extension} {name}: {} cells differ\n  {}",
                    found.len(),
                    found.join("\n  ")
                ));
            }
        }
    }
    assert!(problems.is_empty(), "{}", problems.join("\n"));
}

/// The sequence one operation at a time, each step saved to a file that the next one reads: what a
/// person does with a command line, and what the Python binding does when each call saves.
fn one_at_a_time(extension: &str, specs: &[&str], dir: &Path) -> PathBuf {
    let mut current = structure(&format!("book.{extension}"));
    for (i, spec) in specs.iter().enumerate() {
        let next = dir.join(format!("step{i}.{extension}"));
        editor(&Book::from_file(&current).unwrap(), &[spec])
            .write(&next)
            .unwrap();
        current = next;
    }
    current
}

/// The order the planner runs the operations in: the conversions, then the deletions of rows, of
/// columns, of names and of sheets.
fn planners_order<'a>(specs: &[&'a str]) -> Vec<&'a str> {
    let rank = |spec: &str| {
        [
            "formula_to_value",
            "delete_rows",
            "delete_columns",
            "delete_names",
            "delete_sheets",
        ]
        .iter()
        .position(|verb| spec.starts_with(verb))
        .unwrap()
    };
    let mut ordered = specs.to_vec();
    ordered.sort_by_key(|spec| rank(spec));
    ordered
}

#[test]
fn one_plan_gives_what_its_operations_saved_one_at_a_time_in_the_planners_order_give() {
    for extension in EXTENSIONS {
        let source = Book::from_file(structure(&format!("book.{extension}"))).unwrap();
        for (name, specs) in SEQUENCES {
            let dir = tempfile::tempdir().unwrap();
            let plan = dir.path().join(format!("plan.{extension}"));
            // Whatever order they are pushed in.
            let mut reversed = specs.to_vec();
            reversed.reverse();
            editor(&source, &reversed).write(&plan).unwrap();
            let steps = one_at_a_time(extension, &planners_order(specs), dir.path());
            let found = differences(&cells(&plan), &cells(&steps));
            assert!(
                found.is_empty(),
                "{extension} {name}: one plan and the steps differ\n  {}",
                found.join("\n  ")
            );
        }
    }
}

/// A workbook that says its results are stale (an XLSX marked to recalculate when it is opened; an
/// ODS with no cached result in its formulas) is calculated again before a conversion reads it.
#[test]
fn a_conversion_on_a_workbook_saved_stale_calculates_it_first() {
    for extension in EXTENSIONS {
        let dir = tempfile::tempdir().unwrap();
        let stale = dir.path().join(format!("stale.{extension}"));
        Book::from_file(structure(&format!("book.{extension}")))
            .unwrap()
            .delete_rows(["T!2".parse().unwrap()])
            .write(&stale)
            .unwrap();
        // The values are the ones of the workbook as LibreOffice leaves it.
        let out = dir.path().join(format!("out.{extension}"));
        editor(&Book::from_file(&stale).unwrap(), &["formula_to_value F"])
            .write(&out)
            .unwrap();
        let wanted = structure(&format!("expected_delete_rows_then_value.{extension}"));
        let found = differences(&cells(&out), &cells(&wanted));
        assert!(found.is_empty(), "{extension}: {}", found.join("\n  "));
    }
}

/// A workbook the engine cannot calculate (here, an external reference) cannot have its stale
/// results calculated again: the conversion says so and writes nothing, instead of the stale values.
#[test]
fn a_workbook_that_cannot_be_calculated_is_not_converted_with_stale_values() {
    for extension in EXTENSIONS {
        let source = Book::from_file(
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("tests/assets/book")
                .join(format!("book.{extension}")),
        )
        .unwrap();
        let dir = tempfile::tempdir().unwrap();
        // The deletion is one step, saved; the conversion is the next, on a workbook whose results
        // are stale.
        let deleted = dir.path().join(format!("deleted.{extension}"));
        editor(&source, &[])
            .delete_rows(["Data!10:11".parse().unwrap()])
            .write(&deleted)
            .unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let error = editor(&Book::from_file(&deleted).unwrap(), &[])
            .formula_to_value(["Data!D:E".parse().unwrap()])
            .write(&out)
            .unwrap_err();
        assert!(
            format!("{error:#}").contains("could not be calculated again"),
            "{extension}: {error:#}"
        );
        assert!(!out.exists(), "{extension}: nothing is written");
    }
}
