//! A plan gives the workbook that all its operations give at once, each written against the
//! workbook as it is (not one at a time on the result of the one before), whatever order they were
//! pushed in and however the planner groups them into passes over the file.
//!
//! The fixture is the workbook of `structure_libreoffice`: sheets `T`, `U`, `V` with numbers and
//! formulas, and sheet `F` with formulas over `T`. It holds no reference across sheets
//! (`Sheet1:Sheet3!A1`), which are not supported (issue #75).

use crate::plan::{self, Engine};
use sheetrs_io::Book;
use sheetrs_io::write::Operation;
use sheetrs_io::write::{WorkbookModifications, modify_workbook};
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

type Snapshot = BTreeMap<(String, u32, u32), (Option<String>, String)>;

const EXTENSIONS: [&str; 2] = ["xlsx", "ods"];

fn asset(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join(format!("tests/assets/structure/book.{extension}"))
}

/// Every cell of the workbook with its formula and value. A reference that went reads `T!#REF!`
/// in an XLSX and `#REF!` in an ODS.
fn snapshot(path: &Path) -> Snapshot {
    let book = Book::from_file(path).unwrap().workbook().unwrap();
    let mut cells = BTreeMap::new();
    for sheet in &book.sheets {
        for (at, cell) in sheet.cells.iter() {
            cells.insert(
                (sheet.name.clone(), at.0, at.1),
                (
                    cell.formula
                        .as_deref()
                        .map(|f| f.replace("T!#REF!", "#REF!")),
                    format!("{:?}", cell.value),
                ),
            );
        }
    }
    cells
}

/// One operation from a spec such as `delete_rows T!4:5`, `delete_columns U!B`, `formula_to_value F`, `delete_sheets U`, `delete_names N`;
/// several ranges are separated by a comma and are simultaneous.
fn op(spec: &str) -> Operation {
    let (verb, rest) = spec.split_once(' ').unwrap_or((spec, ""));
    let parts = || rest.split(',').map(str::trim);
    match verb {
        "delete_rows" => Operation::DeleteRows(parts().map(|s| s.parse().unwrap()).collect()),
        "delete_columns" => Operation::DeleteColumns(parts().map(|s| s.parse().unwrap()).collect()),
        "formula_to_value" => {
            Operation::FormulaToValue(parts().map(|s| s.parse().unwrap()).collect())
        }
        "delete_sheets" => Operation::DeleteSheets(parts().map(String::from).collect()),
        "delete_names" => Operation::DeleteNames(parts().map(String::from).collect()),
        other => panic!("unknown operation {other}"),
    }
}

/// The workbook with the operations pushed in the order of `specs`.
fn book(extension: &str, specs: &[&str]) -> Book {
    let mut book = Book::from_file(asset(extension)).unwrap();
    for spec in specs {
        book.push(op(spec));
    }
    book
}

/// The engine of these plans calculates nothing: the workbook comes back as it is. The plans here are
/// about how operations group into passes, and the reference writes the operations straight through
/// the writer, so what a recalculation would change is left to the tests with the engine
/// (`tests/recalculation.rs`).
fn engine() -> Engine {
    Engine {
        recalculate: Some(unchanged),
    }
}

fn unchanged(book: &Book) -> Result<Vec<u8>, String> {
    match (book.path(), book.bytes()) {
        (Some(path), _) => std::fs::read(path).map_err(|e| e.to_string()),
        (None, Some(data)) => Ok(data.to_vec()),
        (None, None) => unreachable!(),
    }
}

fn run(extension: &str, book: &Book) -> Result<Snapshot, String> {
    let dir = tempfile::tempdir().unwrap();
    let output = dir.path().join(format!("out.{extension}"));
    plan::execute(book, &output, &engine()).map_err(|e| format!("{e:#}"))?;
    Ok(snapshot(&output))
}

/// What all the operations ask of the writer at once, against the workbook as it is: the definition
/// of the result of a plan of deletions and conversions.
fn at_once(extension: &str, book: &Book) -> Result<Snapshot, String> {
    let mut mods = WorkbookModifications::default();
    for operation in book.operations() {
        match operation {
            Operation::DeleteSheets(names) => mods
                .delete_sheets
                .get_or_insert_with(Default::default)
                .extend(names.iter().cloned()),
            Operation::DeleteNames(names) => mods
                .delete_names
                .get_or_insert_with(Default::default)
                .extend(names.iter().cloned()),
            Operation::FormulaToValue(scopes) => {
                mods.formula_to_value.extend(scopes.iter().cloned());
            }
            Operation::DeleteRows(rows) => {
                mods.delete_rows
                    .extend(rows.iter().cloned().map(Into::into));
            }
            Operation::DeleteColumns(columns) => {
                mods.delete_columns
                    .extend(columns.iter().cloned().map(Into::into));
            }
            _ => unreachable!("these plans hold deletions and conversions"),
        }
    }
    let dir = tempfile::tempdir().unwrap();
    let output = dir.path().join(format!("out.{extension}"));
    modify_workbook(&asset(extension), &output, &mods).map_err(|e| format!("{e:#}"))?;
    Ok(snapshot(&output))
}

/// A small deterministic generator, so that a failing sequence can be run again.
struct Lcg(u64);

impl Lcg {
    fn next(&mut self, below: usize) -> usize {
        self.0 = self
            .0
            .wrapping_mul(6_364_136_223_846_793_005)
            .wrapping_add(1_442_695_040_888_963_407);
        ((self.0 >> 33) as usize) % below
    }
}

const POOL: &[&str] = &[
    "delete_rows T!2",
    "delete_rows T!4:5",
    "delete_rows T!1",
    "delete_rows T!3, T!9",
    "delete_rows U!3",
    "delete_rows V!2:3",
    "delete_rows F!5",
    "delete_columns T!D",
    "delete_columns U!B",
    "delete_columns V!A",
    "formula_to_value F",
    "formula_to_value T!C:C",
    "formula_to_value U",
    "formula_to_value F!3:6",
    "delete_sheets U",
    "delete_sheets V",
    "delete_names NameOne",
    "delete_names Gone",
];

#[test]
fn a_plan_gives_what_its_operations_give_at_once_against_the_workbook() {
    // Sets of up to five operations of the pool, in both formats, in the order drawn. Where the
    // writer cannot make a change (a merged area that would lose its first cell), both stop.
    let mut random = Lcg(0x5EED);
    let mut agreed = 0;
    for _ in 0..120 {
        let length = 1 + random.next(5);
        let specs: Vec<&str> = (0..length).map(|_| POOL[random.next(POOL.len())]).collect();
        for extension in EXTENSIONS {
            let book = book(extension, &specs);
            let (planned, reference) = (run(extension, &book), at_once(extension, &book));
            if let (Ok(_), Ok(_)) = (&planned, &reference) {
                agreed += 1;
            }
            assert_eq!(
                planned.is_ok(),
                reference.is_ok(),
                "{extension}: {specs:?}: {:?} / {:?}",
                planned.as_ref().err(),
                reference.as_ref().err()
            );
            if planned.is_ok() {
                assert_eq!(planned, reference, "{extension}: {specs:?}");
            }
        }
    }
    // The test means something only if the sets ran.
    assert!(agreed > 100, "{agreed}");
}

#[test]
fn a_set_of_deletions_and_conversions_is_one_pass_however_many_operations_it_has() {
    // The passes of the file, not the recalculation of what they left stale at the end.
    let passes = |specs: &[&str]| {
        plan::passes(book("xlsx", specs).operations())
            .iter()
            .filter(|pass| *pass != &[Operation::Recalculate])
            .count()
    };
    assert_eq!(
        passes(&[
            "delete_rows T!2",
            "delete_columns U!B",
            "delete_sheets V",
            "delete_names NameOne"
        ]),
        1
    );
    assert_eq!(
        passes(&["formula_to_value F", "delete_rows T!2", "delete_rows T!3"]),
        1
    );
    // Rows 2 and 3 touch: one band, in the one pass.
    assert_eq!(
        passes(&["formula_to_value F", "delete_rows T!2", "delete_rows T!3"]),
        1
    );
    assert_eq!(passes(&["delete_rows T!2", "formula_to_value F"]), 1);
    assert_eq!(
        passes(&[
            "delete_rows T!2",
            "formula_to_value F",
            "delete_rows T!1",
            "formula_to_value U"
        ]),
        1
    );
}

#[test]
fn a_conversion_gives_the_values_the_workbook_has_whatever_the_order_it_is_given_in() {
    // F!A2 reads T!A4, which the deletion of row 4 rewrites to `#REF!`: the conversion is of the
    // workbook as it is, so it keeps the value, whether it is pushed before the deletion or after.
    for extension in EXTENSIONS {
        for specs in [
            ["formula_to_value F", "delete_rows T!4:5"],
            ["delete_rows T!4:5", "formula_to_value F"],
        ] {
            let result = run(extension, &book(extension, &specs)).unwrap();
            let (formula, value) = &result[&("F".to_string(), 1, 0)];
            assert_eq!(
                (formula, value.as_str()),
                (&None, "Number(4.0)"),
                "{extension} {specs:?}"
            );
        }
    }
}

#[test]
fn the_order_the_operations_were_pushed_in_changes_nothing() {
    fn permutations(items: &[&str]) -> Vec<Vec<String>> {
        if items.len() <= 1 {
            return vec![items.iter().map(ToString::to_string).collect()];
        }
        let mut all = Vec::new();
        for i in 0..items.len() {
            let mut rest = items.to_vec();
            let head = rest.remove(i);
            for mut tail in permutations(&rest) {
                tail.insert(0, head.to_string());
                all.push(tail);
            }
        }
        all
    }
    // Every order of five operations: 120 pushes with one canonical plan and one result.
    let specs = [
        "delete_rows T!2",
        "delete_rows U!4:5, V!6",
        "formula_to_value F",
        "delete_sheets U",
        "delete_columns T!D",
    ];
    for extension in EXTENSIONS {
        let reference_plan = plan::canonical(book(extension, &specs).operations());
        let reference = run(extension, &book(extension, &specs));
        assert!(reference.is_ok(), "{extension}: {reference:?}");
        for order in permutations(&specs) {
            let order: Vec<&str> = order.iter().map(String::as_str).collect();
            let pushed = book(extension, &order);
            assert_eq!(
                plan::canonical(pushed.operations()),
                reference_plan,
                "{order:?}"
            );
            assert_eq!(run(extension, &pushed), reference, "{extension} {order:?}");
        }
    }
}

#[test]
fn the_same_plan_gives_the_same_workbook_every_time() {
    let specs = [
        "delete_rows T!2",
        "delete_rows U!4:5, V!6",
        "delete_columns T!D",
        "delete_columns U!C",
        "delete_sheets V",
    ];
    for extension in EXTENSIONS {
        let first = run(extension, &book(extension, &specs));
        assert!(first.is_ok(), "{extension}: {first:?}");
        for _ in 0..30 {
            assert_eq!(
                run(extension, &book(extension, &specs)),
                first,
                "{extension} varies between runs"
            );
        }
    }
}

#[test]
fn a_position_is_a_row_of_the_workbook_not_of_what_the_deletions_before_it_left() {
    // Two deletions of "T!3" are the row 3, once: they do not become the rows 3 and 4.
    for extension in EXTENSIONS {
        let twice = run(
            extension,
            &book(extension, &["delete_rows T!3", "delete_rows T!3"]),
        );
        let once = run(extension, &book(extension, &["delete_rows T!3"]));
        assert!(twice.is_ok(), "{extension}: {twice:?}");
        assert_eq!(twice, once, "{extension}");
        // And rows 2 and 3 given apart are the rows 2 and 3 given together.
        let apart = run(
            extension,
            &book(extension, &["delete_rows T!2", "delete_rows T!3"]),
        );
        let together = run(extension, &book(extension, &["delete_rows T!2:3"]));
        assert_eq!(apart, together, "{extension}");
    }
}

#[test]
fn deleting_rows_of_a_sheet_that_is_also_deleted_is_not_an_error() {
    // Both are against the workbook as it is: the rows of a sheet that goes with it go too.
    for extension in EXTENSIONS {
        for specs in [
            ["delete_sheets U", "delete_rows U!2"],
            ["delete_rows U!2", "delete_sheets U"],
        ] {
            assert!(
                run(extension, &book(extension, &specs)).is_ok(),
                "{extension} {specs:?}"
            );
        }
    }
}

/// A book of bytes edits like the file it came from.
#[test]
fn a_plan_runs_on_the_bytes_of_a_workbook_as_on_the_file() {
    let specs = ["delete_rows T!4:5", "formula_to_value F", "delete_sheets U"];
    for extension in ["xlsx", "ods"] {
        let mut from_bytes = Book::from_bytes(std::fs::read(asset(extension)).unwrap()).unwrap();
        for spec in specs {
            from_bytes.push(op(spec));
        }
        let dir = tempfile::tempdir().unwrap();
        let output = dir.path().join(format!("out.{extension}"));
        plan::execute(&from_bytes, &output, &engine()).unwrap();
        assert_eq!(
            Ok(snapshot(&output)),
            run(extension, &book(extension, &specs)),
            "{extension}"
        );
    }
}
