use std::collections::BTreeMap;

use sheetrs_testkit::case::{Case, CellInput, CellValue, Observed, Op, Read, SheetSpec};
use sheetrs_testkit::oracle;

fn sheet(name: &str, cells: &[(&str, CellInput)]) -> SheetSpec {
    SheetSpec {
        name: name.to_owned(),
        cells: cells
            .iter()
            .map(|(address, content)| ((*address).to_owned(), content.clone()))
            .collect::<BTreeMap<_, _>>(),
    }
}

fn read(sheet: &str, range: &str) -> Read {
    Read {
        sheet: sheet.to_owned(),
        range: range.to_owned(),
    }
}

fn formula(text: &str) -> CellInput {
    CellInput::Text(text.to_owned())
}

fn number(v: f64, formula: Option<&str>) -> CellValue {
    CellValue::Number {
        v,
        formula: formula.map(str::to_owned),
    }
}

fn run_one(case: Case) -> Result<Observed, oracle::Error> {
    let mut observed = oracle::run(&[case])?.into_observed()?;
    Ok(observed.remove(0))
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn deleting_a_row_shifts_cells_and_rewrites_references() -> Result<(), oracle::Error> {
    let observed = run_one(Case {
        name: "delete_row_shifts_reference".to_owned(),
        sheets: vec![sheet(
            "S1",
            &[
                ("A1", CellInput::Number(1.0)),
                ("A2", CellInput::Number(2.0)),
                ("A3", CellInput::Number(3.0)),
                ("B1", formula("=SUM(A1:A3)")),
            ],
        )],
        ops: vec![Op::DeleteRows {
            sheet: "S1".to_owned(),
            rows: "2".to_owned(),
        }],
        reads: vec![read("S1", "A1:B2")],
    })?;
    assert_eq!(
        observed.reads[0].cells,
        vec![
            vec![
                CellValue::Number {
                    v: 1.0,
                    formula: None
                },
                number(4.0, Some("=SUM(A1:A2)"))
            ],
            vec![
                CellValue::Number {
                    v: 3.0,
                    formula: None
                },
                CellValue::Empty
            ],
        ]
    );
    Ok(())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn deleting_a_referenced_sheet_leaves_a_reference_error() -> Result<(), oracle::Error> {
    let observed = run_one(Case {
        name: "delete_sheet_breaks_reference".to_owned(),
        sheets: vec![
            sheet("S1", &[("A1", formula("=$S2.A1"))]),
            sheet("S2", &[("A1", CellInput::Number(5.0))]),
        ],
        ops: vec![Op::DeleteSheet {
            sheet: "S2".to_owned(),
        }],
        reads: vec![read("S1", "A1")],
    })?;
    assert!(
        matches!(
            observed.reads[0].cells[0][0],
            CellValue::Error { code: 524, .. }
        ),
        "{:?}",
        observed.reads[0].cells
    );
    Ok(())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn converting_a_formula_to_a_value_drops_the_formula() -> Result<(), oracle::Error> {
    let observed = run_one(Case {
        name: "to_values_drops_formula".to_owned(),
        sheets: vec![sheet(
            "S1",
            &[("A1", CellInput::Number(2.0)), ("B1", formula("=A1*2"))],
        )],
        ops: vec![
            Op::ToValues {
                sheet: "S1".to_owned(),
                range: "B1".to_owned(),
            },
            Op::RecalculateHard,
        ],
        reads: vec![read("S1", "B1")],
    })?;
    assert_eq!(observed.reads[0].cells, vec![vec![number(4.0, None)]]);
    Ok(())
}

/// Converting a formula that evaluates to an error (`#DIV/0!`, `#N/A`, ...) yields an empty text cell.
#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn converting_an_error_formula_to_a_value_leaves_an_empty_string() -> Result<(), oracle::Error> {
    let observed = run_one(Case {
        name: "to_values_error_cell".to_owned(),
        sheets: vec![sheet(
            "S1",
            &[("A1", formula("=1/0")), ("A2", formula("=A1"))],
        )],
        ops: vec![Op::ToValues {
            sheet: "S1".to_owned(),
            range: "A1:A2".to_owned(),
        }],
        reads: vec![read("S1", "A1:A2")],
    })?;
    let empty_text = || CellValue::String {
        v: String::new(),
        formula: None,
    };
    assert_eq!(
        observed.reads[0].cells,
        vec![vec![empty_text()], vec![empty_text()]]
    );
    Ok(())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn deleting_a_column_shifts_cells_and_rewrites_references() -> Result<(), oracle::Error> {
    let observed = run_one(Case {
        name: "delete_column_shifts_reference".to_owned(),
        sheets: vec![sheet(
            "S1",
            &[
                ("A1", CellInput::Number(1.0)),
                ("B1", CellInput::Number(2.0)),
                ("C1", CellInput::Number(3.0)),
                ("D1", formula("=SUM(A1:C1)")),
            ],
        )],
        ops: vec![Op::DeleteColumns {
            sheet: "S1".to_owned(),
            columns: "B".to_owned(),
        }],
        reads: vec![read("S1", "A1:C1")],
    })?;
    assert_eq!(
        observed.reads[0].cells,
        vec![vec![
            number(1.0, None),
            number(3.0, None),
            number(4.0, Some("=SUM(A1:B1)")),
        ]]
    );
    Ok(())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn text_cells_and_error_text_are_observed() -> Result<(), oracle::Error> {
    let observed = run_one(Case {
        name: "text_and_error_cells".to_owned(),
        sheets: vec![sheet(
            "S1",
            &[
                ("A1", CellInput::Text("hello".to_owned())),
                ("B1", formula("=1/0")),
            ],
        )],
        ops: vec![Op::Recalculate],
        reads: vec![read("S1", "A1:B1")],
    })?;
    assert_eq!(
        observed.reads[0].cells,
        vec![vec![
            CellValue::String {
                v: "hello".to_owned(),
                formula: None
            },
            CellValue::Error {
                code: 532,
                text: "#DIV/0!".to_owned(),
                formula: Some("=1/0".to_owned())
            },
        ]]
    );
    Ok(())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn numbers_keep_every_bit_through_the_oracle() -> Result<(), oracle::Error> {
    let observed = run_one(Case {
        name: "number_bits".to_owned(),
        sheets: vec![sheet(
            "S1",
            &[
                ("A1", formula("=0.1+0.2")),
                ("A2", CellInput::Number(6.934_698_245_112_694e276)),
            ],
        )],
        ops: vec![],
        reads: vec![read("S1", "A1:A2")],
    })?;
    let bits = |row: usize| match observed.reads[0].cells[row][0] {
        CellValue::Number { v, .. } => v.to_bits(),
        ref other => panic!("expected a number, found {other:?}"),
    };
    assert_eq!(bits(0), (0.1_f64 + 0.2).to_bits());
    assert_eq!(bits(1), 6.934_698_245_112_694e276_f64.to_bits());
    Ok(())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn spans_that_are_not_plain_rows_or_columns_fail_their_case() -> Result<(), oracle::Error> {
    let broken = |name: &str, op: Op| Case {
        name: name.to_owned(),
        sheets: vec![sheet("S1", &[("A1", CellInput::Number(1.0))])],
        ops: vec![op],
        reads: vec![read("S1", "A1")],
    };
    let rows = |spec: &str| Op::DeleteRows {
        sheet: "S1".to_owned(),
        rows: spec.to_owned(),
    };
    let columns = |spec: &str| Op::DeleteColumns {
        sheet: "S1".to_owned(),
        columns: spec.to_owned(),
    };
    let cases = vec![
        broken("underscore", rows("1_0")),
        broken("open_end", rows("2:")),
        broken("sign", rows("+3")),
        broken("space", rows(" 4")),
        broken("column_open_end", columns("B:")),
        broken("column_digit", columns("1")),
    ];
    let report = oracle::run(&cases)?;
    for outcome in &report.results {
        assert!(
            matches!(outcome, oracle::Outcome::Failed { .. }),
            "{outcome:?}"
        );
    }
    Ok(())
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn a_malformed_row_span_fails_only_its_case() -> Result<(), oracle::Error> {
    let broken = Case {
        name: "reversed_row_span".to_owned(),
        sheets: vec![sheet("S1", &[("A1", CellInput::Number(1.0))])],
        ops: vec![Op::DeleteRows {
            sheet: "S1".to_owned(),
            rows: "3:2".to_owned(),
        }],
        reads: vec![read("S1", "A1")],
    };
    let healthy = Case {
        name: "healthy".to_owned(),
        sheets: vec![sheet("S1", &[("A1", CellInput::Number(1.0))])],
        ops: vec![],
        reads: vec![read("S1", "A1")],
    };
    let report = oracle::run(&[broken, healthy])?;
    assert!(
        matches!(
            &report.results[0],
            oracle::Outcome::Failed { error, .. } if error.contains("3:2")
        ),
        "{:?}",
        report.results[0]
    );
    assert!(matches!(report.results[1], oracle::Outcome::Observed(_)));
    Ok(())
}
