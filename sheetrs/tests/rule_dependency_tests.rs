//! Disabling a rule only removes its own violations, even when other enabled
//! rules depend on it (issue #53).

use sheetrs::reader::{Cell, CellValue, Sheet, Workbook, read_workbook};
use sheetrs::violation::ExcelError;
use sheetrs::{FormatContext, Linter, LinterConfig, RuleId};

const TEST_XLSX: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/minimal_test.xlsx");
const TEST_ODS: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/../tests/minimal_test.ods");

/// Each provider with the rules that depend on it.
const DEPENDENCIES: &[(RuleId, &[RuleId])] = &[
    (
        RuleId::Calc202,
        &[
            RuleId::Err102,
            RuleId::Err103,
            RuleId::Ref306,
            RuleId::Ref309,
            RuleId::Ref310,
            RuleId::Data706,
        ],
    ),
    (RuleId::Err102, &[RuleId::Err103]),
    (
        RuleId::Cpx504,
        &[
            RuleId::Cpx505,
            RuleId::Cpx506,
            RuleId::Cpx508,
            RuleId::Cpx509,
        ],
    ),
];

/// Rules that depend on a provider without being one.
fn dependents_only() -> Vec<RuleId> {
    let mut ids: Vec<RuleId> = DEPENDENCIES
        .iter()
        .flat_map(|(_, dependents)| dependents.iter().copied())
        .filter(|id| !DEPENDENCIES.iter().any(|(provider, _)| provider == id))
        .collect();
    ids.sort();
    ids.dedup();
    ids
}

/// Lints `workbook` with `config` and returns the formatted violations.
fn lint(workbook: &Workbook, config: LinterConfig) -> Vec<(RuleId, String)> {
    let ctx = FormatContext { workbook };
    Linter::with_config(config)
        .lint_workbook(workbook)
        .unwrap()
        .iter()
        .map(|v| {
            let line = format!("{:?} {}", v.scope, v.format_message(&ctx));
            (v.rule_id, line)
        })
        .collect()
}

fn cell(row: u32, col: u32, value: CellValue, formula: Option<&str>) -> Cell {
    Cell {
        row,
        col,
        value,
        formula: formula.map(Box::from),
        ..Default::default()
    }
}

/// Builds a workbook where every rule in [`DEPENDENCIES`] reports a violation.
fn dependency_workbook() -> Workbook {
    let long_formula = format!("LEN(\"{}\")", "x".repeat(260));
    let mut cells: Vec<Cell> = (0..6)
        .map(|row| cell(row, 0, CellValue::Number(f64::from(row + 1)), None))
        .collect();
    cells.extend([
        // REF310: A6 continues the summed range
        cell(0, 1, CellValue::Empty, Some("SUM(A1:A5)")),
        // DATA706
        cell(0, 2, CellValue::Text("42".into()), None),
        cell(0, 3, CellValue::Empty, Some("C1+1")),
        // ERR103 from E1; ERR102 on G1
        cell(0, 4, CellValue::Error(ExcelError::DivZero), None),
        cell(0, 5, CellValue::Empty, Some("E1*2")),
        cell(0, 6, CellValue::Error(ExcelError::NA), None),
        // REF309: J1 is empty
        cell(0, 7, CellValue::Empty, Some("J1+A1")),
        // CALC202
        cell(1, 1, CellValue::Empty, Some("B3")),
        cell(2, 1, CellValue::Empty, Some("B2")),
        // CPX505, CPX506, CPX508, CPX509
        cell(1, 4, CellValue::Empty, Some("((((((A1))))))")),
        cell(
            2,
            4,
            CellValue::Empty,
            Some("A1+A2+A3+A4+A5+A6+A1+A2+A3+A4"),
        ),
        cell(
            3,
            4,
            CellValue::Empty,
            Some("SUM(A1,A2,A3,A4,A5,A6,A1,A2,A3,A4,A5)"),
        ),
        cell(4, 4, CellValue::Empty, Some(&long_formula)),
        // CPX504
        cell(
            5,
            4,
            CellValue::Empty,
            Some("IF(A1,IF(A2,IF(A3,IF(A4,IF(A5,IF(A6,1))))))"),
        ),
    ]);
    let mut data = Sheet::new("Data".to_string(), 0);
    data.cells = cells.into_iter().map(|c| ((c.row, c.col), c)).collect();
    // REF306: no formula refers to this sheet
    let mut unused = Sheet::new("Unused".to_string(), 1);
    unused
        .cells
        .insert((0, 0), cell(0, 0, CellValue::Number(1.0), None));
    Workbook {
        sheets: vec![data, unused],
        ..Default::default()
    }
}

fn assert_disabling_provider_keeps_dependents(workbook: &Workbook, label: &str) {
    let baseline = lint(workbook, LinterConfig::default());
    for &(provider, _) in DEPENDENCIES {
        let mut config = LinterConfig::default();
        config.global.disabled_rules.insert(provider.to_string());
        let without_provider = lint(workbook, config);
        let expected: Vec<_> = baseline.iter().filter(|(id, _)| *id != provider).collect();
        let actual: Vec<_> = without_provider.iter().collect();
        assert_eq!(
            actual, expected,
            "disabling {provider} changed other rules' output in {label}"
        );
    }
}

fn assert_enabling_dependents_alone_keeps_them(workbook: &Workbook, label: &str) {
    let dependents = dependents_only();
    let baseline = lint(workbook, LinterConfig::default());
    let mut config = LinterConfig::default();
    config
        .global
        .enabled_rules
        .extend(dependents.iter().map(RuleId::to_string));
    let expected: Vec<_> = baseline
        .iter()
        .filter(|(id, _)| dependents.contains(id))
        .collect();
    let actual = lint(workbook, config);
    assert_eq!(
        actual.iter().collect::<Vec<_>>(),
        expected,
        "enabling only {dependents:?} changed their output in {label}"
    );
}

#[test]
fn test_dependency_workbook_reports_every_rule() {
    let baseline = lint(&dependency_workbook(), LinterConfig::default());
    for &(provider, dependents) in DEPENDENCIES {
        for id in dependents.iter().chain([&provider]) {
            assert!(
                baseline.iter().any(|(reported, _)| reported == id),
                "no {id} violation in the dependency workbook"
            );
        }
    }
}

#[test]
fn test_disabling_provider_keeps_dependents() {
    assert_disabling_provider_keeps_dependents(&dependency_workbook(), "dependency workbook");
}

#[test]
fn test_enabling_dependents_alone_keeps_them() {
    assert_enabling_dependents_alone_keeps_them(&dependency_workbook(), "dependency workbook");
}

#[test]
fn test_disabling_provider_keeps_dependents_xlsx() {
    assert_disabling_provider_keeps_dependents(&read_workbook(TEST_XLSX).unwrap(), TEST_XLSX);
}

#[test]
fn test_disabling_provider_keeps_dependents_ods() {
    assert_disabling_provider_keeps_dependents(&read_workbook(TEST_ODS).unwrap(), TEST_ODS);
}

#[test]
fn test_enabling_dependents_alone_keeps_them_xlsx() {
    assert_enabling_dependents_alone_keeps_them(&read_workbook(TEST_XLSX).unwrap(), TEST_XLSX);
}

#[test]
fn test_enabling_dependents_alone_keeps_them_ods() {
    assert_enabling_dependents_alone_keeps_them(&read_workbook(TEST_ODS).unwrap(), TEST_ODS);
}
