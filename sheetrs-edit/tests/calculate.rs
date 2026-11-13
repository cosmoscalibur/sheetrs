//! `calculate` is a step of the plan: the engine writes the
//! inputs and calculates every formula, and the workbook is saved with the results and without the
//! mark to recalculate it when it is opened. Where the engine cannot calculate the workbook the
//! mark is set and the run warns; the outputs asked for in that case are an error.

use serde_json::{Map, Value as Json, json};
use sheetrs_edit::{BookEdit, Calculation, Error, RECALC_NOT_SUPPORTED};
use sheetrs_io::Book;
use sheetrs_io::reader::CellValue;
use sheetrs_protocol::MajorDimension;
use std::io::Read;
use std::path::{Path, PathBuf};

const EXTENSIONS: [&str; 2] = ["xlsx", "ods"];

fn structure(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/structure")
        .join(format!("book.{extension}"))
}

/// The editor workbook: it has an external reference the engine cannot calculate.
fn external(extension: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/assets/book")
        .join(format!("book.{extension}"))
}

fn editor(path: &Path) -> Book {
    Book::from_file(path).unwrap()
}

fn value(path: &Path, sheet: &str, row: u32, column: u32) -> CellValue {
    let workbook = Book::from_file(path).unwrap().workbook().unwrap();
    let sheet = workbook.sheets.iter().find(|s| s.name == sheet).unwrap();
    sheet
        .cells
        .iter()
        .find(|(at, _)| (at.0, at.1) == (row, column))
        .map_or(CellValue::Empty, |(_, cell)| cell.value.clone())
}

fn number(path: &Path, sheet: &str, row: u32, column: u32) -> f64 {
    match value(path, sheet, row, column) {
        CellValue::Number(n) => n,
        other => panic!("{sheet} {row},{column} is {other:?}, not a number"),
    }
}

/// Whether the workbook is marked to be recalculated when it is opened: an XLSX says it in
/// `calcPr`, an ODS has no mark and holds no cached result in its formula cells.
fn marked(path: &Path) -> bool {
    if path.extension().is_some_and(|e| e == "xlsx") {
        let mut archive = zip::ZipArchive::new(std::fs::File::open(path).unwrap()).unwrap();
        let mut xml = String::new();
        archive
            .by_name("xl/workbook.xml")
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        xml.contains("fullCalcOnLoad=\"1\"") || xml.contains("fullCalcOnLoad=\"true\"")
    } else {
        let workbook = Book::from_file(path).unwrap().workbook().unwrap();
        // Some formula that holds no cached result (a formula that caches an error of its own, as
        // an external reference does, keeps it).
        workbook
            .sheets
            .iter()
            .flat_map(|s| s.cells.iter())
            .any(|(_, cell)| cell.formula.is_some() && matches!(cell.value, CellValue::Empty))
    }
}

#[test]
fn a_deletion_runs_after_the_calculation_whatever_the_order_and_what_it_leaves_stale_is_calculated()
{
    // Everything is written against the workbook as it is: the calculation is of the workbook, and
    // the deletion comes after it and leaves the results of the formulas that read the row stale, so
    // the engine calculates them again as the last thing: no mark, fresh results.
    for extension in EXTENSIONS {
        let outputs: Vec<_> = [true, false]
            .into_iter()
            .map(|calculation_first| {
                let dir = tempfile::tempdir().unwrap();
                let out = dir.path().join(format!("out.{extension}"));
                let mut editor = editor(&structure(extension));
                let calculation = Calculation::default();
                if calculation_first {
                    editor.calculate(calculation);
                    editor.delete_rows(["T!2".parse().unwrap()]);
                } else {
                    editor.delete_rows(["T!2".parse().unwrap()]);
                    editor.calculate(calculation);
                }
                let report = editor.write(&out).unwrap();
                assert!(
                    report.diagnostics.is_empty(),
                    "{extension}: {:?}",
                    report.diagnostics
                );
                assert!(!marked(&out), "{extension}: the workbook is marked");
                // F!A5 is SUM(T!A2:A3): the row 2 went, it adds the one row that is left.
                assert!((number(&out, "F", 4, 0) - 3.0).abs() < 1e-9, "{extension}");
                std::fs::read(&out).unwrap()
            })
            .collect();
        assert_eq!(
            outputs[0], outputs[1],
            "{extension}: the order changed the result"
        );
    }
}

#[test]
fn a_deletion_alone_has_its_stale_results_calculated_and_no_mark() {
    for extension in EXTENSIONS {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut editor = editor(&structure(extension));
        editor.delete_rows(["T!2".parse().unwrap()]);
        let report = editor.write(&out).unwrap();
        assert!(
            report.diagnostics.is_empty(),
            "{extension}: {:?}",
            report.diagnostics
        );
        assert!(!marked(&out), "{extension}");
        assert!((number(&out, "F", 4, 0) - 3.0).abs() < 1e-9, "{extension}");
    }
}

#[test]
fn the_mark_is_only_where_the_engine_fails_and_the_report_says_why() {
    // The `book` workbook has an external reference the engine cannot calculate: the deletion leaves
    // results stale, the engine fails at calculating them again, and then (and only then) the
    // workbook is marked to be recalculated when it is opened.
    for extension in EXTENSIONS {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut editor = editor(&external(extension));
        editor.delete_rows(["Data!10:11".parse().unwrap()]);
        let report = editor.write(&out).unwrap();
        assert!(
            matches!(report.diagnostics.as_slice(), [d] if d.code == RECALC_NOT_SUPPORTED && d.context.get("reason").is_some_and(|r| !r.is_empty())),
            "{extension}: {:?}",
            report.diagnostics
        );
        assert!(marked(&out), "{extension}");
    }
}

#[test]
fn the_inputs_are_written_and_the_formulas_calculated() {
    for extension in EXTENSIONS {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut inputs = Map::new();
        inputs.insert("T!A1".into(), json!(10));
        let mut editor = editor(&structure(extension));
        editor.calculate(Calculation {
            inputs,
            ranges: None,
            major: MajorDimension::default(),
            answer: false,
            ..Calculation::default()
        });
        editor.write(&out).unwrap();
        // T!C1 is SUM(A1).
        assert!((number(&out, "T", 0, 2) - 10.0).abs() < 1e-9, "{extension}");
        assert!(!marked(&out), "{extension}");
    }
}

#[test]
fn a_workbook_the_engine_cannot_calculate_is_written_marked_and_the_run_warns() {
    for extension in EXTENSIONS {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut editor = editor(&external(extension));
        editor
            .delete_rows(["Data!10:11".parse().unwrap()])
            .calculate(Calculation {
                inputs: Map::new(),
                ranges: None,
                major: MajorDimension::default(),
                answer: false,
                ..Calculation::default()
            });
        let report = editor.write(&out).unwrap();
        assert!(
            matches!(report.diagnostics.as_slice(), [d] if d.code == RECALC_NOT_SUPPORTED),
            "{extension}: {:?}",
            report.diagnostics
        );
        assert!(
            report.diagnostics[0]
                .message
                .contains("recalculation it needs will be done by the spreadsheet program"),
            "{}",
            report.diagnostics[0]
        );
        assert!(marked(&out), "{extension}: the workbook is not marked");
    }
}

#[test]
fn a_calculation_as_the_only_step_that_cannot_run_still_marks_the_workbook() {
    for extension in EXTENSIONS {
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        let mut editor = editor(&external(extension));
        editor.calculate(Calculation {
            inputs: Map::new(),
            ranges: None,
            major: MajorDimension::default(),
            answer: false,
            ..Calculation::default()
        });
        let report = editor.write(&out).unwrap();
        assert_eq!(report.diagnostics.len(), 1, "{extension}");
        assert!(marked(&out), "{extension}");
    }
}

#[test]
fn the_outputs_asked_for_of_a_workbook_the_engine_cannot_calculate_are_an_error() {
    let dir = tempfile::tempdir().unwrap();
    let mut editor = editor(&external("xlsx"));
    editor.calculate(Calculation {
        answer: true,
        ..Calculation::default()
    });
    let error = editor.write(dir.path().join("out.xlsx")).unwrap_err();
    assert!(
        format!("{error:#}").contains("could not be calculated"),
        "{error:#}"
    );
    assert!(!dir.path().join("out.xlsx").exists());
}

#[test]
fn a_run_that_did_not_ask_for_outputs_has_none_to_write() {
    let dir = tempfile::tempdir().unwrap();
    let mut editor = editor(&structure("xlsx"));
    editor.delete_rows(["T!2".parse().unwrap()]);
    let report = editor.write(dir.path().join("out.xlsx")).unwrap();
    assert!(report.value_ranges.is_none());
    let error = report.write_json(Vec::new()).unwrap_err();
    assert!(matches!(error, Error::NoCalculation), "{error:#}");
}

#[test]
fn the_outputs_are_kept_in_the_report_and_written_as_json() {
    for extension in EXTENSIONS {
        let mut editor = editor(&structure(extension));
        editor
            .delete_rows(["T!2".parse().unwrap()])
            .calculate(Calculation {
                ranges: Some(vec!["F!A5:A6".to_string()]),
                answer: true,
                ..Calculation::default()
            });
        // The outputs without writing the workbook anywhere.
        let report = editor.write_to(std::io::sink()).unwrap();
        let mut json = Vec::new();
        report.write_json(&mut json).unwrap();
        let answer: Json = serde_json::from_slice(&json).unwrap();
        let values = &answer["valueRanges"][0]["values"];
        // F!A5 is SUM(T!A2:A3) of the workbook as it is: the row 2 is deleted after.
        assert!(
            (values[0][0].as_f64().unwrap() - 5.0).abs() < 1e-9,
            "{extension}: {answer}"
        );
    }
}

#[test]
fn the_major_of_a_calculation_lays_the_answer_out_as_rows_or_as_columns() {
    let answer = |major| {
        let mut editor = editor(&structure("xlsx"));
        editor.calculate(Calculation {
            ranges: Some(vec!["F!A5:A6".to_string()]),
            major,
            answer: true,
            ..Calculation::default()
        });
        let report = editor.write_to(std::io::sink()).unwrap();
        let mut json = Vec::new();
        report.write_json(&mut json).unwrap();
        let answer: Json = serde_json::from_slice(&json).unwrap();
        answer["valueRanges"][0]["values"].clone()
    };
    let rows = answer(MajorDimension::Rows);
    let columns = answer(MajorDimension::Columns);
    assert_eq!(rows.as_array().unwrap().len(), 2, "{rows}");
    assert_eq!(columns.as_array().unwrap().len(), 1, "{columns}");
    assert_eq!(columns[0][0], rows[0][0]);
    assert_eq!(columns[0][1], rows[1][0]);
}

#[test]
fn without_the_outputs_kept_the_report_holds_none() {
    let mut editor = editor(&structure("xlsx"));
    editor.calculate(Calculation {
        inputs: Map::new(),
        ranges: None,
        major: MajorDimension::default(),
        answer: false,
        ..Calculation::default()
    });
    let report = editor.write_to(std::io::sink()).unwrap();
    assert!(report.value_ranges.is_none());
    assert!(matches!(
        report.write_json(Vec::new()),
        Err(Error::NoCalculation)
    ));
}

#[test]
fn explain_says_the_calculation_and_what_happens_if_it_cannot_run() {
    let mut editor = editor(&structure("xlsx"));
    editor
        .delete_rows(["T!2".parse().unwrap()])
        .calculate(Calculation {
            inputs: Map::new(),
            ranges: None,
            major: MajorDimension::default(),
            answer: false,
            ..Calculation::default()
        });
    let text = editor.explain();
    assert!(
        text.contains("calculate every formula, with 0 inputs"),
        "{text}"
    );
    // The calculation is the first pass whatever the order they were pushed in, the deletion after it
    // leaves the results stale, and the engine calculates them again as the last pass.
    assert!(text.contains("pass 3 of 3"), "{text}");
    assert!(text.find("calculate every").unwrap() < text.find("delete rows").unwrap());
    assert!(
        text.contains("mark it to be recalculated when opened"),
        "{text}"
    );
}

#[test]
fn a_conversion_converts_the_calculated_results_in_memory_too() {
    for extension in EXTENSIONS {
        let mut editor = editor(&structure(extension));
        editor
            .delete_rows(["T!2".parse().unwrap()])
            .calculate(Calculation {
                inputs: Map::new(),
                ranges: None,
                major: MajorDimension::default(),
                answer: false,
                ..Calculation::default()
            })
            .formula_to_value(["F".parse().unwrap()]);
        let mut bytes = Vec::new();
        editor.write_to(&mut bytes).unwrap();
        let dir = tempfile::tempdir().unwrap();
        let out = dir.path().join(format!("out.{extension}"));
        std::fs::write(&out, &bytes).unwrap();
        // The conversion is of the calculated workbook as it is: F!A5 is SUM(T!A2:A3), 5.
        assert!((number(&out, "F", 4, 0) - 5.0).abs() < 1e-9, "{extension}");
    }
}

#[test]
fn a_calculation_on_a_workbook_that_comes_marked_leaves_it_unmarked() {
    use sheetrs_io::write::{WorkbookModifications, modify_workbook};
    for extension in EXTENSIONS {
        let dir = tempfile::tempdir().unwrap();
        let marked_file = dir.path().join(format!("marked.{extension}"));
        let mods = WorkbookModifications {
            recalculate_on_open: true,
            ..WorkbookModifications::default()
        };
        modify_workbook(&structure(extension), &marked_file, &mods).unwrap();
        assert!(marked(&marked_file), "{extension}");
        let out = dir.path().join(format!("out.{extension}"));
        let mut second = editor(&marked_file);
        second.calculate(Calculation {
            inputs: Map::new(),
            ranges: None,
            major: MajorDimension::default(),
            answer: false,
            ..Calculation::default()
        });
        second.write(&out).unwrap();
        assert!(!marked(&out), "{extension}");
        assert!((number(&out, "F", 4, 0) - 5.0).abs() < 1e-9, "{extension}");
    }
}

#[test]
fn the_outputs_of_a_plan_are_given_without_writing_the_workbook() {
    for extension in EXTENSIONS {
        let mut editor = editor(&structure(extension));
        editor
            .delete_rows(["T!2".parse().unwrap()])
            .calculate(Calculation {
                ranges: Some(vec!["F!A5:A6".to_string()]),
                ..Calculation::default()
            });
        let value_ranges = editor.value_ranges().unwrap();
        let entries = value_ranges.to_entries(Some(&["F!A5:A6".to_string()]));
        assert_eq!(entries.len(), 1, "{extension}");
        assert!(
            (entries[0].1[0][0].as_f64().unwrap() - 5.0).abs() < 1e-9,
            "{extension}: {entries:?}"
        );
        // The plan is not consumed: the same call again gives the same value ranges.
        assert_eq!(editor.value_ranges().unwrap(), value_ranges, "{extension}");
    }
}

#[test]
fn a_plan_without_a_calculation_has_no_value_ranges_to_give() {
    let mut editor = editor(&structure("xlsx"));
    editor.delete_rows(["T!2".parse().unwrap()]);
    assert!(matches!(editor.value_ranges(), Err(Error::NoCalculation)));
}
