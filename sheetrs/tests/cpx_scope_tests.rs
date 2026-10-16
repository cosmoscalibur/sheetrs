use sheetrs::reader::{Cell, CellValue, Sheet, Workbook};
use sheetrs::violation::CellReference;
use sheetrs::{Linter, RuleId, ViolationScope};

fn formula_cell(row: u32, col: u32, formula: &str) -> Cell {
    Cell {
        row,
        col,
        value: CellValue::Empty,
        formula: Some(<Box<str>>::from(formula)),
        is_array: false,
        num_fmt: None,
    }
}

fn sheet_with(name: &str, index: u16, cells: Vec<Cell>) -> Sheet {
    let mut sheet = Sheet::new(name.to_string(), index);
    for cell in cells {
        sheet.cells.insert((cell.row, cell.col), cell);
    }
    sheet
}

/// Formula over `max_formula_length` that triggers only CPX509.
fn long_formula() -> String {
    format!("=\"{}\"", "x".repeat(300))
}

fn lint(workbook: &Workbook) -> Vec<(RuleId, ViolationScope)> {
    Linter::new()
        .lint_workbook(workbook)
        .unwrap()
        .into_iter()
        .map(|v| (v.rule_id, v.scope))
        .collect()
}

/// CPX504-CPX509 violations are ordered by sheet, then numeric row/column,
/// whatever the `HashMap` iteration order of the walker.
#[test]
fn test_cpx_violations_sorted_by_sheet_then_cell() {
    let long = long_formula();
    let positions = [(9, 0), (0, 1), (1, 0), (2, 0), (9, 5)];
    let workbook = Workbook {
        sheets: vec![
            sheet_with(
                "S0",
                0,
                positions
                    .iter()
                    .map(|&(r, c)| formula_cell(r, c, &long))
                    .collect(),
            ),
            sheet_with("S1", 1, vec![formula_cell(1, 0, &long)]),
        ],
        ..Default::default()
    };

    for _ in 0..20 {
        let got: Vec<_> = lint(&workbook)
            .into_iter()
            .filter(|(id, _)| *id == RuleId::Cpx509)
            .map(|(_, scope)| scope)
            .collect();
        let expected: Vec<_> = [(0, 1), (1, 0), (2, 0), (9, 0), (9, 5)]
            .iter()
            .map(|&(r, c)| ViolationScope::Cell(0, CellReference::new(r, c)))
            .chain([ViolationScope::Cell(1, CellReference::new(1, 0))])
            .collect();
        assert_eq!(got, expected);
    }
}

/// Every rule of the CPX5xx cascade and CPX507 reports the formula cell.
#[test]
fn test_cpx504_to_cpx509_report_cell_scope() {
    let cases: [(RuleId, String); 6] = [
        (
            RuleId::Cpx504,
            "=IF(A1,IF(B1,IF(C1,IF(D1,IF(E1,IF(F1,1,0),0),0),0),0),0)".to_string(),
        ),
        (
            RuleId::Cpx505,
            "=ABS(MIN(MAX(ROUND(ABS(MIN(A1,2)),2),0),100))".to_string(),
        ),
        (RuleId::Cpx506, "=1+2+3+4+5+6+7+8+9+10".to_string()),
        (RuleId::Cpx507, "=Sheet1!A1+Sheet2!B2+Sheet3!C3".to_string()),
        (
            RuleId::Cpx508,
            "=SUM(A1,B1,C1,D1,E1,F1,G1,H1,I1,J1,K1)".to_string(),
        ),
        (RuleId::Cpx509, long_formula()),
    ];
    let workbook = Workbook {
        sheets: vec![
            Sheet::new("S0".to_string(), 0),
            sheet_with(
                "S1",
                1,
                cases
                    .iter()
                    .enumerate()
                    .map(|(i, (_, f))| formula_cell(i as u32, 3, f))
                    .collect(),
            ),
        ],
        ..Default::default()
    };

    let violations = lint(&workbook);
    for (i, (rule, formula)) in cases.iter().enumerate() {
        let expected = ViolationScope::Cell(1, CellReference::new(i as u32, 3));
        assert!(
            violations.contains(&(*rule, expected)),
            "{rule} not reported at cell scope for {formula}"
        );
    }
    assert!(
        !violations
            .iter()
            .any(|(id, scope)| (RuleId::Cpx504..=RuleId::Cpx509).contains(id)
                && matches!(scope, ViolationScope::Sheet(_))),
        "CPX504-CPX509 must not use sheet scope"
    );
}
