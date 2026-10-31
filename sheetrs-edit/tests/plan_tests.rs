//! A plan gives the workbook that its operations give one at a time, each on the result of the
//! one before, however it groups them into passes over the file.
//!
//! The fixture is the workbook of `structure_parity_tests`: sheets `T`, `U`, `V` with numbers and
//! formulas, and sheet `F` with formulas over `T`. It holds no reference across sheets
//! (`Sheet1:Sheet3!A1`), which are not supported (issue #75).

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

/// One operation from a spec such as `rows T!4:5`, `cols U!B`, `convert F`, `sheet U`, `name N`,
/// `recalc`; several scopes are separated by a comma and are simultaneous.
fn op(spec: &str) -> Operation {
    let (verb, rest) = spec.split_once(' ').unwrap_or((spec, ""));
    let scopes = || rest.split(',').map(|s| s.trim().parse().unwrap()).collect();
    let names = || rest.split(',').map(|s| s.trim().to_string()).collect();
    match verb {
        "rows" => Operation::DeleteRows(scopes()),
        "cols" => Operation::DeleteColumns(scopes()),
        "convert" => Operation::ConvertToValues(scopes()),
        "sheet" => Operation::RemoveSheets(names()),
        "name" => Operation::RemoveNamedRanges(names()),
        "recalc" => Operation::RecalculateOnOpen,
        other => panic!("unknown operation {other}"),
    }
}

fn sequence(specs: &[&str]) -> EditPlan {
    let mut plan = EditPlan::new();
    for spec in specs {
        plan.push(op(spec)).unwrap();
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

/// The plan one operation per pass: what running them one at a time gives, the definition of its
/// result. An operation that cannot run stops it.
fn run_one_at_a_time(extension: &str, plan: &EditPlan) -> Result<Snapshot, String> {
    let dir = tempfile::tempdir().unwrap();
    let mut current = asset(extension);
    for (i, operation) in plan.operations().iter().enumerate() {
        let mut single = EditPlan::new();
        single.push(operation.clone()).unwrap();
        let output = dir.path().join(format!("step{i}.{extension}"));
        single
            .execute(&current, &output)
            .map_err(|e| format!("{e:#}"))?;
        current = output;
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
    "recalc",
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
        &[
            "rows T!2, U!4:5",
            "cols T!D",
            "convert F",
            "name NameOne",
            "recalc",
        ],
    ];
    for specs in scenarios {
        let declared = EditPlan::against_input(specs.iter().map(|spec| op(spec))).unwrap();
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
        let reference_plan = EditPlan::against_input(specs.iter().map(|s| op(s))).unwrap();
        let reference = run(extension, &reference_plan);
        assert!(reference.is_ok(), "{extension}: {reference:?}");
        for order in permutations(&specs) {
            let declared = EditPlan::against_input(order.iter().map(|s| op(s))).unwrap();
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
/// `T!#REF!` in an XLSX and `#REF!` in an ODS, and LibreOffice writes the range of one cell as the
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

/// Whether the sequence converts formulas after something else changed the workbook (or after the
/// mark to recalculate): a conversion then reads cached results that the earlier change dropped,
/// and the two formats do not agree on what remains (see the ignored test below).
fn converts_after_a_change(specs: &[&str]) -> bool {
    let mut changed = false;
    for spec in specs {
        let verb = spec.split(' ').next().unwrap();
        match verb {
            "convert" if changed => return true,
            "convert" => {}
            _ => changed = true,
        }
    }
    false
}

fn cross_format_mismatches(skip_unsupported: bool) -> (usize, Vec<String>) {
    let mut random = Lcg(0xFACE);
    let (mut compared, mut mismatches) = (0, Vec::new());
    for _ in 0..120 {
        let length = 1 + random.next(4);
        let specs: Vec<&str> = (0..length).map(|_| POOL[random.next(POOL.len())]).collect();
        if skip_unsupported == converts_after_a_change(&specs) {
            continue;
        }
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
    (compared, mismatches)
}

/// The same plan on the XLSX and on the ODS of one workbook gives the same formulas in every cell,
/// for the sequences that do not convert after a change: the format differences of streaming
/// (shared formulas and `A1` text in one, repeated rows and `[$T.A1:.A3]` notation in the other)
/// do not show in what the formulas become.
#[test]
fn both_formats_give_the_same_formulas_for_the_same_plan() {
    let (compared, mismatches) = cross_format_mismatches(true);
    assert!(
        mismatches.is_empty(),
        "{} of {compared} differ between formats: {mismatches:#?}",
        mismatches.len()
    );
    assert!(compared > 40, "{compared} compared");
}

/// KNOWN DIFFERENCE, not fixed: converting formulas after a change that dropped their cached
/// results (a deletion that rewrote them, or `recalc_on_open`, which an ODS applies by dropping
/// every cached result and an XLSX by a flag) gives other workbooks in each format. Run with
/// `--ignored` to see the sequences.
#[test]
#[ignore = "known difference between formats when converting after a change"]
fn both_formats_agree_when_converting_after_a_change() {
    let (compared, mismatches) = cross_format_mismatches(false);
    assert!(
        mismatches.is_empty(),
        "{} of {compared} differ between formats: {mismatches:#?}",
        mismatches.len()
    );
}
