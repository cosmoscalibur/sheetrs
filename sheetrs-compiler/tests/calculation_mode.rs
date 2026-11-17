//! The calculation modes of a build and of a run, as `LibreOffice` has them.
//!
//! By default the values the workbook stores are the constants a build folds and the cache of the
//! formulas it keeps, and a run calculates only what the inputs it is given change (the load of
//! `LibreOffice`, which does not calculate a formula that has a result). `calculate_all` emulates
//! `calculateAll()`: a build calculates the whole workbook first and stores those results; a run ignores
//! the cache. `force_calculate_all` (a build only) stores no cache, so that the program always
//! calculates what it kept.
//!
//! The workbook of the tests: `Data!A1:A3` hold 1, 2 and 3, and `Calc` holds three formulas that cache a
//! result no engine would give (the cached value says which one a program read):
//!
//! | cell | formula | caches | reads the input `Data!A1` |
//! |---|---|---|---|
//! | `B1` | `Data!A1*2` | 999 | yes |
//! | `B2` | `Data!A2*2` | 888 | no |
//! | `B3` | `B1+B2` | 777 | through `B1` |

use serde_json::{Value as Json, json};
use sheetrs_compiler::model::ModelCell;
use sheetrs_compiler::{BindingSpec, Model, Spec, compile};
use sheetrs_engine::addr::CellRef;
use sheetrs_engine::calculate::Overrides;
use sheetrs_engine::ir::Program;
use sheetrs_engine::value::Value;

const CALC: u16 = 1;

fn model() -> Model {
    let mut model = Model::new();
    let data = model.add_sheet("Data");
    model.add_sheet("Calc");
    for (row, value) in ["A1", "A2", "A3"].into_iter().zip([1.0, 2.0, 3.0]) {
        model.set(data, row, Value::Number(value));
    }
    for (row, text, cached) in [
        (0, "Data!A1*2", 999.0),
        (1, "Data!A2*2", 888.0),
        (2, "B1+B2", 777.0),
    ] {
        model.cells.insert(
            CellRef::new(CALC, row, 1),
            ModelCell {
                value: Value::Number(cached),
                formula: Some(text.to_string()),
                array: false,
                recalculated_on_load: false,
            },
        );
    }
    model
}

/// Forty formulas that nothing reads and no output asks for, so that a few formulas with no result are
/// far from a quarter of the workbook (where the build calculates everything and does not look).
fn fillers(model: &mut Model) {
    for row in 0..40 {
        model.cells.insert(
            CellRef::new(CALC, row, 3),
            ModelCell {
                value: Value::Number(3.0),
                formula: Some("Data!A3*1".to_string()),
                array: false,
                recalculated_on_load: false,
            },
        );
    }
}

fn spec() -> Spec {
    Spec {
        inputs: vec![BindingSpec::named("a", "Data!A1")],
        inputs_declared: true,
        outputs: vec![BindingSpec::named("out", "Calc!B1:B3")],
        ..Spec::default()
    }
}

fn build(spec: &Spec) -> Program {
    let mut model = model();
    fillers(&mut model);
    compile(model, spec).unwrap().0
}

/// `Calc!B1:B3` for an input `a` of `input` (or the one the workbook holds), with the given overrides.
fn answer(program: &Program, input: Option<f64>, calculate_all: bool) -> Vec<f64> {
    let request = input.map_or(Json::Null, |a| json!({"a": a}));
    let overrides = Overrides {
        calculate_all,
        ..Overrides::default()
    };
    let out = program.calculate(&request, &overrides).unwrap();
    out["out"]
        .as_array()
        .unwrap()
        .iter()
        .map(|row| row[0].as_f64().unwrap())
        .collect()
}

#[test]
fn by_default_a_run_gives_the_stored_values_when_no_input_changes() {
    let program = build(&spec());
    assert_eq!(answer(&program, None, false), [999.0, 888.0, 777.0]);
    // An input given with the value the workbook holds changes nothing either.
    assert_eq!(answer(&program, Some(1.0), false), [999.0, 888.0, 777.0]);
}

#[test]
fn by_default_a_changed_input_calculates_what_it_reaches_and_nothing_else() {
    let program = build(&spec());
    // `B1` and `B3` are in the cone of the input; `B2` is not, and it is not calculated: it keeps the
    // value the workbook stored (it was folded as it).
    assert_eq!(answer(&program, Some(10.0), false), [20.0, 888.0, 908.0]);
}

#[test]
fn a_formula_no_input_reaches_is_folded_as_the_value_the_workbook_stores() {
    let (program, report) = compile(model(), &spec()).unwrap();
    assert_eq!(report.formulas_cached, 1, "B2 keeps its stored result");
    let b2 = program
        .cells
        .iter()
        .find(|c| c.at == CellRef::new(CALC, 1, 1))
        .unwrap();
    assert_eq!(
        b2.kind,
        sheetrs_engine::ir::CellKind::Const(Value::Number(888.0))
    );
}

#[test]
fn the_program_keeps_the_cached_result_of_the_formulas_an_input_reaches() {
    let program = build(&spec());
    let cached: Vec<&Value> = program.cache.iter().map(|(_, v)| v).collect();
    assert_eq!(cached.len(), 2);
    assert!(cached.contains(&&Value::Number(999.0)));
    assert!(cached.contains(&&Value::Number(777.0)));
}

#[test]
fn a_build_with_calculate_all_stores_the_whole_workbook_calculated() {
    let program = build(&Spec {
        calculate_all: true,
        ..spec()
    });
    // The folded constant is the result of the calculation (4), and so is the cache of what is kept.
    assert_eq!(answer(&program, None, false), [2.0, 4.0, 6.0]);
    assert_eq!(answer(&program, Some(10.0), false), [20.0, 4.0, 24.0]);
    assert!(program.build.options.calculate_all);
    assert!(!program.build.options.force_calculate_all);
}

#[test]
fn force_calculate_all_stores_no_cache_so_the_program_calculates_what_it_kept() {
    let program = build(&Spec {
        force_calculate_all: true,
        ..spec()
    });
    assert!(program.cache.is_empty());
    // `B2` was folded as the stored value; what was kept is calculated, with no input given as well.
    assert_eq!(answer(&program, None, false), [2.0, 888.0, 890.0]);
    assert!(program.build.options.force_calculate_all);
}

#[test]
fn force_calculate_all_makes_a_lighter_program() {
    let with = serde_json::to_vec(&build(&spec())).unwrap().len();
    let without = serde_json::to_vec(&build(&Spec {
        force_calculate_all: true,
        ..spec()
    }))
    .unwrap()
    .len();
    assert!(without < with, "{without} < {with}");
}

#[test]
fn calculate_all_ignores_the_cache_of_unfolded_cells() {
    let program = build(&spec());
    // `B1` and `B3` are calculated; `B2` is a constant of the program (folded), so it is not.
    assert_eq!(answer(&program, None, true), [2.0, 888.0, 890.0]);
    assert_eq!(answer(&program, Some(10.0), true), [20.0, 888.0, 908.0]);
}

#[test]
fn a_run_with_calculate_all_has_no_effect_on_a_program_built_with_force() {
    let program = build(&Spec {
        force_calculate_all: true,
        ..spec()
    });
    assert_eq!(
        answer(&program, Some(10.0), true),
        answer(&program, Some(10.0), false)
    );
}

#[test]
fn a_formula_with_no_cached_result_is_calculated_and_so_is_what_reads_it() {
    let mut model = model();
    model.cells.insert(
        CellRef::new(CALC, 1, 1),
        ModelCell {
            value: Value::Empty,
            formula: Some("Data!A2*2".to_string()),
            array: false,
            recalculated_on_load: false,
        },
    );
    fillers(&mut model);
    let program = compile(model, &spec()).unwrap().0;
    // `B2` has no result: it is calculated (4), and `B3` reads it, so it is too.
    assert_eq!(answer(&program, None, false), [999.0, 4.0, 1003.0]);
}

#[test]
fn a_volatile_formula_keeps_no_result() {
    let mut model = model();
    model.cells.insert(
        CellRef::new(CALC, 3, 1),
        ModelCell {
            value: Value::Number(5.0),
            formula: Some("Data!A1+RAND()*0".to_string()),
            array: false,
            recalculated_on_load: false,
        },
    );
    fillers(&mut model);
    let spec = Spec {
        outputs: vec![BindingSpec::named("out", "Calc!B1:B4")],
        ..spec()
    };
    let program = compile(model, &spec).unwrap().0;
    let slots: Vec<u32> = program.cache.iter().map(|(slot, _)| *slot).collect();
    let volatile = program
        .cells
        .iter()
        .position(|c| c.at == CellRef::new(CALC, 3, 1))
        .unwrap() as u32;
    assert!(!slots.contains(&volatile));
}

// -- the command line: `sheetc calculate`, `sheetc build` and the executable it writes --------------------------

use std::path::{Path, PathBuf};
use std::process::Command;

const SHEETC: &str = env!("CARGO_BIN_EXE_sheetc");

/// `book.xlsx` of the edit tests: `Calc!A1` is `SUM(Data!A1:A6)`, and `Calc!A7` caches `from-excel` where the engine
/// gives `t`; nothing the input `Data!A2` reaches reads `A7`.
fn book(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../sheetrs-edit/tests/assets/cached")
        .join(format!("book.{extension}"))
}

fn values(program: &Path, args: &[&str]) -> (Json, String) {
    let out = Command::new(program).args(args).output().unwrap();
    assert!(
        out.status.success(),
        "{}",
        String::from_utf8_lossy(&out.stderr)
    );
    let text = String::from_utf8_lossy(&out.stdout).into_owned();
    (serde_json::from_str(&text).unwrap_or(Json::Null), text)
}

fn cell(answer: &Json, range: usize) -> Json {
    answer["valueRanges"][range]["values"][0][0].clone()
}

#[test]
fn calculate_has_the_mode_of_the_build_and_calculate_all_calculates_everything() {
    let request = r#"{"inputs": {"Data!A2": 25}}"#;
    for extension in ["xlsx", "ods"] {
        let file = book(extension);
        let file = file.to_str().unwrap();
        let ask = |flag: &[&str]| {
            let mut args = vec![
                "calculate",
                "-i",
                file,
                "--request",
                request,
                "--range",
                "Calc!A1",
                "--range",
                "Calc!A7",
            ];
            args.extend_from_slice(flag);
            values(Path::new(SHEETC), &args).0
        };
        let by_default = ask(&[]);
        assert_eq!(
            cell(&by_default, 0),
            json!(215),
            "{extension}: reached by the input"
        );
        assert_eq!(
            cell(&by_default, 1),
            json!("from-excel"),
            "{extension}: kept"
        );
        let all = ask(&["--calculate-all"]);
        assert_eq!(cell(&all, 0), json!(215), "{extension}");
        assert_eq!(cell(&all, 1), json!("t"), "{extension}: calculated");
    }
}

#[test]
fn calculate_all_ignores_the_cache_of_an_executable_unless_built_with_force() {
    let dir = tempfile::tempdir().unwrap();
    let settings = dir.path().join("declarations.json");
    std::fs::write(
        &settings,
        r#"{"inputs": [{"name": "a2", "range": "Data!A2"}],
            "outputs": [{"name": "first", "range": "Calc!A1"}, {"name": "kept", "range": "Calc!A7"}]}"#,
    )
    .unwrap();
    let build = |name: &str, flag: &[&str]| {
        let exe = dir.path().join(name);
        let file = book("xlsx");
        let mut args = vec![
            "build",
            "-i",
            file.to_str().unwrap(),
            "--settings-file",
            settings.to_str().unwrap(),
            "-o",
            exe.to_str().unwrap(),
        ];
        args.extend_from_slice(flag);
        let out = Command::new(SHEETC).args(&args).output().unwrap();
        assert!(
            out.status.success(),
            "{}",
            String::from_utf8_lossy(&out.stderr)
        );
        exe
    };
    let answer = |exe: &Path, flag: &[&str]| {
        let mut args = vec!["--request", r#"{"inputs": {"a2": 25}}"#];
        args.extend_from_slice(flag);
        values(exe, &args).0
    };
    let cache = build("cache", &[]);
    let forced = build("forced", &["--force-calculate-all"]);
    // With the cache, the cell the input does not reach is the one the workbook stores...
    let a = answer(&cache, &[]);
    assert_eq!(a["valueRanges"][0]["values"][0][0], json!(215));
    assert_eq!(a["valueRanges"][1]["values"][0][0], json!("from-excel"));
    // ... and `--calculate-all` of the executable calculates what the program kept (the formula `A7` reads `Data!B1`,
    // which is not an input: the program kept it, and calculates it).
    let all = answer(&cache, &["--calculate-all"]);
    assert_eq!(all["valueRanges"][0]["values"][0][0], json!(215));
    // A program built with `--force-calculate-all` has no cache: the run is the same with and without the flag.
    assert_eq!(answer(&forced, &[]), answer(&forced, &["--calculate-all"]));
    assert!(std::fs::metadata(&forced).unwrap().len() < std::fs::metadata(&cache).unwrap().len());
}

#[test]
fn request_inputs_are_the_build_inputs_unless_the_workbook_declares_them() {
    use sheetrs_compiler::memory::{BookCompile, Options};
    use sheetrs_engine::ir::Mode;
    let book = sheetrs_io::Book::from_file(book("xlsx")).unwrap();
    let asked = Options {
        request_inputs: vec!["Data!A2".into()],
        ..Options::default()
    };
    // What the input does not reach is folded as the result the workbook stores...
    let (_, report) = book.compile(&asked, Mode::Embedded).unwrap();
    assert!(report.formulas_cached > 0, "{report:?}");
    // ... and with no input to count on, every formula reads a constant that may be one: nothing is folded.
    let (_, report) = book.compile(&Options::default(), Mode::Embedded).unwrap();
    assert_eq!(report.formulas_cached, 0, "{report:?}");
}
