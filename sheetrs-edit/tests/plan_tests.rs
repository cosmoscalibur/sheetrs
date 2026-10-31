//! A plan gives the workbook that its schedule gives one operation per pass, whatever order the
//! operations were written in, and gives it every time.
//!
//! The fixture is the workbook of `structure_parity_tests` (sheets `T`, `U`, `V` with numbers and
//! formulas, and sheet `F` with formulas over `T` and over spans of sheets such as `T:V!A2:A4`).

use sheetrs_core::reader::read_workbook;
use sheetrs_edit::{EditPlan, Operation};
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
    let book = read_workbook(path).unwrap();
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

fn plan(specs: &[&str]) -> EditPlan {
    let mut plan = EditPlan::new();
    for spec in specs {
        let (verb, rest) = spec.split_once(' ').unwrap_or((spec, ""));
        match verb {
            "rows" => plan.delete_rows(rest.parse().unwrap()).map(|_| ()).unwrap(),
            "cols" => plan
                .delete_columns(rest.parse().unwrap())
                .map(|_| ())
                .unwrap(),
            "convert" => {
                plan.convert_to_values(rest.parse().unwrap());
            }
            "sheet" => {
                plan.remove_sheet(rest);
            }
            "name" => {
                plan.remove_named_range(rest);
            }
            "recalc" => {
                plan.recalculate_on_open();
            }
            other => panic!("unknown operation {other}"),
        }
    }
    plan
}

fn run(extension: &str, plan: &EditPlan) -> Result<Snapshot, String> {
    let dir = tempfile::tempdir().unwrap();
    let output = dir.path().join(format!("out.{extension}"));
    plan.execute(asset(extension), &output)
        .map_err(|e| format!("{e:#}"))?;
    Ok(snapshot(&output))
}

/// The schedule of the plan, one operation per pass, each on the output of the previous one.
fn run_stepwise(extension: &str, plan: &EditPlan) -> Result<Snapshot, String> {
    let dir = tempfile::tempdir().unwrap();
    let mut current = asset(extension);
    for (i, step) in plan.schedule().into_iter().enumerate() {
        let mut single = EditPlan::new();
        single.push(step).unwrap();
        let output = dir.path().join(format!("step{i}.{extension}"));
        single
            .execute(&current, &output)
            .map_err(|e| format!("{e:#}"))?;
        current = output;
    }
    Ok(snapshot(&current))
}

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

const SCENARIOS: &[(&str, &[&str])] = &[
    ("two bands", &["rows T!4:5", "rows T!9"]),
    (
        "overlapping and touching",
        &["rows T!3:6", "rows T!5:8", "rows T!9"],
    ),
    (
        "rows and columns",
        &["rows T!4:5", "rows T!9", "cols T!B", "cols T!D"],
    ),
    ("convert before a deletion", &["rows T!4:5", "convert F"]),
    (
        "convert and delete a column",
        &["cols T!D", "convert T!C:C"],
    ),
    (
        "convert rows that are deleted",
        &["rows T!2:3", "convert T!3:6"],
    ),
    (
        "a deletion in a sheet that goes",
        &["sheet U", "rows U!2", "convert U"],
    ),
    ("a sheet that F reads goes", &["sheet T", "convert F"]),
    (
        "a sheet and a deletion elsewhere",
        &["rows T!4:5", "sheet U"],
    ),
    (
        "rows of two sheets under a span",
        &["rows T!4:5", "rows U!3"],
    ),
    (
        "rows of three sheets under a span",
        &["rows T!2", "rows U!4:5", "rows V!6"],
    ),
    (
        "rows and columns of two sheets",
        &["rows T!4:5", "cols U!B"],
    ),
    (
        "deletions and a sheet that goes",
        &["rows T!4:5", "rows V!3", "sheet U"],
    ),
    ("two sheets go", &["rows T!4:5", "sheet U", "sheet V"]),
    (
        "everything",
        &[
            "rows T!2",
            "rows U!4:5",
            "cols T!D",
            "convert F",
            "name NameOne",
            "recalc",
        ],
    ),
];

#[test]
fn a_plan_gives_what_its_schedule_gives_one_operation_per_pass() {
    for extension in EXTENSIONS {
        for (name, specs) in SCENARIOS {
            let plan = plan(specs);
            let (all_at_once, stepwise) = (run(extension, &plan), run_stepwise(extension, &plan));
            assert_eq!(
                all_at_once, stepwise,
                "{extension}: '{name}' differs between the one pass and the schedule step by step"
            );
        }
    }
}

#[test]
fn the_order_the_operations_were_written_in_changes_nothing() {
    // Every order of the five operations: 120 plans with one schedule and one result.
    let specs = ["rows T!2", "rows U!4:5", "rows V!6", "convert F", "sheet U"];
    for extension in EXTENSIONS {
        let reference_plan = plan(&specs);
        let reference = run(extension, &reference_plan);
        assert!(reference.is_ok(), "{extension}: {reference:?}");
        for order in permutations(&specs) {
            let order: Vec<&str> = order.iter().map(String::as_str).collect();
            let shuffled = plan(&order);
            assert_eq!(shuffled.schedule(), reference_plan.schedule(), "{order:?}");
            assert_eq!(
                run(extension, &shuffled),
                reference,
                "{extension} {order:?}"
            );
        }
    }
}

#[test]
fn the_same_plan_gives_the_same_workbook_every_time() {
    let plan = plan(&[
        "rows T!2",
        "rows U!4:5",
        "rows V!6",
        "cols T!D",
        "cols U!C",
        "sheet V",
    ]);
    for extension in EXTENSIONS {
        let first = run(extension, &plan);
        assert!(first.is_ok(), "{extension}: {first:?}");
        for _ in 0..30 {
            assert_eq!(
                run(extension, &plan),
                first,
                "{extension} varies between runs"
            );
        }
    }
}

#[test]
fn a_conversion_keeps_the_values_that_a_deletion_would_otherwise_cost() {
    // F!A2 reads T!A4, which the plan deletes. Converted first, F keeps its value; the formulas
    // of the other cells that read the deleted rows are rewritten.
    for extension in EXTENSIONS {
        let converted = run(extension, &plan(&["rows T!4:5", "convert F"])).unwrap();
        let (formula, value) = &converted[&("F".to_string(), 1, 0)];
        assert_eq!(formula, &None, "{extension}");
        assert_eq!(value, "Number(4.0)", "{extension}");

        let not_converted = run(extension, &plan(&["rows T!4:5"])).unwrap();
        assert_eq!(
            not_converted[&("F".to_string(), 1, 0)].0.as_deref(),
            Some("#REF!"),
            "{extension}"
        );
    }
}

#[test]
fn a_schedule_is_made_of_operations_that_run_on_their_own() {
    let plan = plan(&["rows T!4:5", "rows T!9", "cols T!D", "convert F", "sheet U"]);
    for step in plan.schedule() {
        assert!(
            !matches!(step, Operation::RecalculateOnOpen),
            "not part of this plan"
        );
        let mut single = EditPlan::new();
        single.push(step.clone()).unwrap();
        run("xlsx", &single).unwrap_or_else(|e| panic!("{step}: {e}"));
    }
}
