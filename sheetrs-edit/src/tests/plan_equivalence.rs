//! A plan gives the workbook that its operations give one at a time, each on the result of the
//! one before, however it groups them into passes over the file.
//!
//! The fixture is the workbook of `structure_parity_tests`: sheets `T`, `U`, `V` with numbers and
//! formulas, and sheet `F` with formulas over `T`. It holds no reference across sheets
//! (`Sheet1:Sheet3!A1`), which are not supported (issue #75).

use crate::plan::{Operation, Plan};
use crate::writer::{ValueTarget, WorkbookModifications, modify_workbook};
use sheetrs_core::reader::{Source, read_workbook};
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

/// One operation from a spec such as `rows T!4:5`, `cols U!B`, `convert F`, `sheet U`, `name N`;
/// several ranges are separated by a comma and are simultaneous.
fn op(spec: &str) -> Operation {
    let (verb, rest) = spec.split_once(' ').unwrap_or((spec, ""));
    let parts = || rest.split(',').map(str::trim);
    match verb {
        "rows" => Operation::DeleteRows(parts().map(|s| s.parse().unwrap()).collect()),
        "cols" => Operation::DeleteColumns(parts().map(|s| s.parse().unwrap()).collect()),
        "convert" => Operation::ConvertToValues(parts().map(|s| s.parse().unwrap()).collect()),
        "sheet" => Operation::RemoveSheets(parts().map(String::from).collect()),
        "name" => Operation::RemoveNamedRanges(parts().map(String::from).collect()),
        other => panic!("unknown operation {other}"),
    }
}

/// A plan from specs.
fn sequence(specs: &[&str]) -> Plan {
    let mut plan = Plan::new();
    for spec in specs {
        plan.push(op(spec));
    }
    plan
}

/// The plan for operations that are all written against the **input** workbook, in any order,
/// as a command line gives them: they are grouped by kind into the order of a pass, so the
/// plan has one pass whatever the order they came in, and the same operations give the same
/// plan. This is the declarative counterpart of the sequence: the same operations given as a
/// sequence would need their positions translated by whoever wrote them.
fn declared(operations: impl IntoIterator<Item = Operation>) -> Plan {
    fn unique<T: PartialEq>(into: &mut Vec<T>, from: Vec<T>) {
        for item in from {
            if !into.contains(&item) {
                into.push(item);
            }
        }
    }
    let (mut converts, mut rows, mut columns) = (Vec::new(), Vec::new(), Vec::new());
    let (mut names, mut sheets) = (Vec::new(), Vec::new());
    for operation in operations {
        match operation {
            Operation::ConvertToValues(scopes) => unique(&mut converts, scopes),
            Operation::DeleteRows(ranges) => unique(&mut rows, ranges),
            Operation::DeleteColumns(ranges) => unique(&mut columns, ranges),
            Operation::RemoveNamedRanges(n) => names.extend(n),
            Operation::RemoveSheets(n) => sheets.extend(n),
        }
    }
    names.sort();
    names.dedup();
    sheets.sort();
    sheets.dedup();
    converts.sort_by_key(|scope| (scope.sheet.clone(), target_key(scope.target)));
    rows.sort();
    columns.sort();
    let mut plan = Plan::new();
    plan.convert_to_values(converts)
        .delete_rows(rows)
        .delete_columns(columns)
        .remove_named_ranges(names)
        .remove_sheets(sheets);
    plan
}

fn target_key(target: ValueTarget) -> (u8, u32, u32) {
    match target {
        ValueTarget::Sheet => (0, 0, 0),
        ValueTarget::Columns(first, last) => (1, first, last),
        ValueTarget::Rows(first, last) => (2, first, last),
    }
}

fn run(extension: &str, plan: &Plan) -> Result<Snapshot, String> {
    let dir = tempfile::tempdir().unwrap();
    let output = dir.path().join(format!("out.{extension}"));
    plan.execute(&Source::from_file(asset(extension)).unwrap(), &output)
        .map_err(|e| format!("{e:#}"))?;
    Ok(snapshot(&output))
}

/// What one operation asks of the writer.
fn modifications(operation: &Operation) -> WorkbookModifications {
    let mut mods = WorkbookModifications::default();
    match operation {
        Operation::RemoveSheets(names) => {
            mods.remove_sheets = Some(names.iter().cloned().collect());
        }
        Operation::RemoveNamedRanges(names) => {
            mods.remove_named_ranges = Some(names.iter().cloned().collect());
        }
        Operation::ConvertToValues(scopes) => mods.convert_to_values.clone_from(scopes),
        Operation::DeleteRows(rows) => {
            mods.delete_rows = rows.iter().cloned().map(Into::into).collect();
        }
        Operation::DeleteColumns(columns) => {
            mods.delete_columns = columns.iter().cloned().map(Into::into).collect();
        }
    }
    mods
}

/// The plan one operation per pass, straight through the writer: what running them one at a time
/// gives, the definition of its result. The mark to recalculate, which the plan decides, is one
/// more pass at the end. An operation that cannot run stops it.
fn run_one_at_a_time(extension: &str, plan: &Plan) -> Result<Snapshot, String> {
    let dir = tempfile::tempdir().unwrap();
    let mut current = asset(extension);
    let mut step = |mods: &WorkbookModifications, name: String| {
        let output = dir.path().join(format!("{name}.{extension}"));
        modify_workbook(&current, &output, mods).map_err(|e| format!("{e:#}"))?;
        current = output;
        Ok::<(), String>(())
    };
    for (i, operation) in plan.operations().iter().enumerate() {
        step(&modifications(operation), format!("step{i}"))?;
    }
    if plan.marks_recalculation() {
        let marked = WorkbookModifications {
            recalculate_on_open: true,
            ..Default::default()
        };
        step(&marked, "marked".to_string())?;
    }
    Ok(snapshot(&current))
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
    "rows T!2",
    "rows T!4:5",
    "rows T!1",
    "rows T!3, T!9",
    "rows U!3",
    "rows V!2:3",
    "rows F!5",
    "cols T!D",
    "cols U!B",
    "cols V!A",
    "convert F",
    "convert T!C:C",
    "convert U",
    "convert F!3:6",
    "sheet U",
    "sheet V",
    "name NameOne",
    "name Gone",
];

#[test]
fn a_plan_gives_what_its_operations_give_one_at_a_time() {
    // Sequences of up to five operations of the pool, in both formats. Where an operation cannot
    // run (a sheet already removed, a merged area that would lose its first cell), both stop.
    let mut random = Lcg(0x5EED);
    let (mut agreed, mut stopped, mut merged_passes) = (0, 0, 0);
    for _ in 0..120 {
        let length = 1 + random.next(5);
        let specs: Vec<&str> = (0..length).map(|_| POOL[random.next(POOL.len())]).collect();
        let plan = sequence(&specs);
        if plan.passes().is_ok_and(|p| p.len() < length) {
            merged_passes += 1;
        }
        for extension in EXTENSIONS {
            let (together, apart) = (run(extension, &plan), run_one_at_a_time(extension, &plan));
            match (&together, &apart) {
                (Ok(_), Ok(_)) => agreed += 1,
                (Err(_), Err(_)) => stopped += 1,
                _ => {}
            }
            assert_eq!(
                together.is_ok(),
                apart.is_ok(),
                "{extension}: {specs:?}: {:?} / {:?}",
                together.as_ref().err(),
                apart.as_ref().err()
            );
            if together.is_ok() {
                assert_eq!(together, apart, "{extension}: {specs:?}");
            }
        }
    }
    // The test means something only if it ran both kinds of sequence and merged passes.
    assert!(
        agreed > 100 && stopped > 0 && merged_passes > 30,
        "{agreed} {stopped} {merged_passes}"
    );
}

#[test]
fn the_sequences_that_need_more_than_one_pass_are_the_ones_that_convert_after_a_change() {
    let passes = |specs: &[&str]| sequence(specs).passes().unwrap().len();
    assert_eq!(
        passes(&["rows T!2", "cols U!B", "sheet V", "name NameOne"]),
        1
    );
    assert_eq!(passes(&["convert F", "rows T!2", "rows T!3"]), 1);
    // Rows 2 and 3 of the input touch: two deletions to the writer, not one band.
    assert_eq!(passes(&["convert F", "rows T!2", "rows T!2"]), 2);
    assert_eq!(passes(&["rows T!2", "convert F"]), 2);
    assert_eq!(
        passes(&["rows T!2", "convert F", "rows T!1", "convert U"]),
        3
    );
}

#[test]
fn a_conversion_after_a_deletion_sees_what_the_deletion_did() {
    // F!A2 reads T!A4, which row 4 deletion rewrites to `#REF!` and leaves without a result: the
    // conversion that comes after finds no value, where one that comes before keeps it.
    for extension in EXTENSIONS {
        let before = run(extension, &sequence(&["convert F", "rows T!4:5"])).unwrap();
        let (formula, value) = &before[&("F".to_string(), 1, 0)];
        assert_eq!(
            (formula, value.as_str()),
            (&None, "Number(4.0)"),
            "{extension}"
        );

        let after = run(extension, &sequence(&["rows T!4:5", "convert F"])).unwrap();
        assert_ne!(
            after[&("F".to_string(), 1, 0)].1,
            "Number(4.0)",
            "{extension}"
        );
    }
}

#[test]
fn against_input_is_one_pass_equal_to_its_operations_one_at_a_time() {
    let scenarios: &[&[&str]] = &[
        &["rows T!4:5, T!9"],
        &["rows T!3:6, T!5:8, T!9"],
        &["rows T!4:5", "cols T!B, T!D"],
        &["rows T!4:5", "convert F"],
        &["cols T!D", "convert T!C:C"],
        &["rows T!2:3", "convert T!3:6"],
        &["sheet U", "rows U!2", "convert U"],
        &["sheet T", "convert F"],
        &["rows T!4:5", "sheet U"],
        &["rows T!4:5, U!3"],
        &["rows T!2, U!4:5, V!6"],
        &["rows T!4:5, V!3", "sheet U"],
        &["rows T!2, U!4:5", "cols T!D", "convert F", "name NameOne"],
    ];
    for specs in scenarios {
        let declared = declared(specs.iter().map(|spec| op(spec)));
        assert_eq!(declared.passes().unwrap().len(), 1, "{specs:?}");
        for extension in EXTENSIONS {
            assert_eq!(
                run(extension, &declared),
                run_one_at_a_time(extension, &declared),
                "{extension}: {specs:?}"
            );
        }
    }
}

#[test]
fn the_order_the_operations_were_declared_in_changes_nothing() {
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
    // Every order of five operations: 120 plans with one plan and one result.
    let specs = [
        "rows T!2",
        "rows U!4:5, V!6",
        "convert F",
        "sheet U",
        "cols T!D",
    ];
    for extension in EXTENSIONS {
        let reference_plan = declared(specs.iter().map(|s| op(s)));
        let reference = run(extension, &reference_plan);
        assert!(reference.is_ok(), "{extension}: {reference:?}");
        for order in permutations(&specs) {
            let declared = declared(order.iter().map(|s| op(s)));
            assert_eq!(declared, reference_plan, "{order:?}");
            assert_eq!(
                run(extension, &declared),
                reference,
                "{extension} {order:?}"
            );
        }
    }
}

#[test]
fn the_same_plan_gives_the_same_workbook_every_time() {
    let plan = sequence(&[
        "rows T!2",
        "rows U!4:5, V!6",
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

/// The operations on different sheets do not depend on each other, which is why they share a pass
/// and a plan needs no order between sheets.
#[test]
fn the_edits_of_different_sheets_commute() {
    for (a, b) in [
        ("rows T!4:5", "rows U!3"),
        ("rows T!4:5", "cols U!B"),
        ("cols T!D", "cols V!A"),
        ("rows T!2", "rows F!5"),
    ] {
        for extension in EXTENSIONS {
            assert_eq!(
                run_one_at_a_time(extension, &sequence(&[a, b])),
                run_one_at_a_time(extension, &sequence(&[b, a])),
                "{extension}: {a} / {b}"
            );
        }
    }
}

#[test]
fn a_position_is_the_one_the_deletions_before_it_left() {
    // After row 3 goes, row 3 is the old row 4: two deletions of "T!3" are the rows 3 and 4.
    for extension in EXTENSIONS {
        let twice = run(extension, &sequence(&["rows T!3", "rows T!3"]));
        let at_once = run(extension, &sequence(&["rows T!3:4"]));
        assert!(twice.is_ok(), "{extension}: {twice:?}");
        assert_eq!(twice, at_once, "{extension}");
    }
    // The positions are translated through a deletion that is not next to them, too.
    for extension in EXTENSIONS {
        let one_by_one = run(extension, &sequence(&["rows T!2", "rows T!3"]));
        let at_once = run(extension, &sequence(&["rows T!2, T!4"]));
        assert!(one_by_one.is_ok(), "{extension}: {one_by_one:?}");
        assert_eq!(one_by_one, at_once, "{extension}");
    }
}

#[test]
fn an_operation_on_a_sheet_that_was_removed_stops_the_plan_before_it_runs() {
    let plan = sequence(&["sheet U", "rows U!2"]);
    for extension in EXTENSIONS {
        let error = run(extension, &plan).unwrap_err();
        assert!(error.contains("removed"), "{extension}: {error}");
    }
}

/// The formulas of a snapshot, one reference form for both formats: a reference that went reads
/// `T!#REF!` in an XLSX and `#REF!` in an ODS, and `LibreOffice` writes the range of one cell as the
/// cell in an XLSX and as a range in an ODS.
fn formulas_only(snapshot: &Snapshot) -> BTreeMap<(String, u32, u32), String> {
    let one_cell = regex::Regex::new(r"(\$?[A-Z]+\$?[0-9]+):(\$?[A-Z]+\$?[0-9]+)").unwrap();
    snapshot
        .iter()
        .filter_map(|(at, (formula, _))| {
            let formula = formula.as_ref()?.replace("T!#REF!", "#REF!");
            let formula = one_cell.replace_all(&formula, |c: &regex::Captures| {
                if c[1] == c[2] {
                    c[1].to_string()
                } else {
                    c[0].to_string()
                }
            });
            Some((at.clone(), formula.into_owned()))
        })
        .collect()
}

/// The same plan on the XLSX and on the ODS of one workbook gives the same formulas in every cell:
/// the format differences of streaming (shared formulas and `A1` text in one, repeated rows and
/// `[$T.A1:.A3]` notation in the other) do not show in what the formulas become.
#[test]
fn both_formats_give_the_same_formulas_for_the_same_plan() {
    let mut random = Lcg(0xFACE);
    let (mut compared, mut mismatches) = (0, Vec::new());
    for _ in 0..200 {
        let length = 1 + random.next(5);
        let specs: Vec<&str> = (0..length).map(|_| POOL[random.next(POOL.len())]).collect();
        let plan = sequence(&specs);
        match (run("xlsx", &plan), run("ods", &plan)) {
            (Ok(x), Ok(o)) => {
                compared += 1;
                if formulas_only(&x) != formulas_only(&o) {
                    mismatches.push(format!("{specs:?}"));
                }
            }
            (Err(_), Err(_)) => {}
            (x, o) => mismatches.push(format!("{specs:?}: xlsx {:?}, ods {:?}", x.err(), o.err())),
        }
    }
    assert!(
        mismatches.is_empty(),
        "{} of {compared} differ between formats: {mismatches:#?}",
        mismatches.len()
    );
    assert!(compared > 100, "{compared} compared");
}

/// A source of bytes edits like the file it came from.
#[test]
fn a_plan_runs_on_the_bytes_of_a_workbook_as_on_the_file() {
    use sheetrs_core::reader::Format;
    let plan = sequence(&["rows T!4:5", "convert F", "sheet U"]);
    for (extension, format) in [("xlsx", Format::Xlsx), ("ods", Format::Ods)] {
        let bytes = std::fs::read(asset(extension)).unwrap();
        let dir = tempfile::tempdir().unwrap();
        let output = dir.path().join(format!("out.{extension}"));
        plan.execute(&Source::from_bytes(bytes, format), &output)
            .unwrap();
        assert_eq!(Ok(snapshot(&output)), run(extension, &plan), "{extension}");
    }
}
