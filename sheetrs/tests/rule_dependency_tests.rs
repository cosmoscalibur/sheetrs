//! Disabling a rule only removes its own violations, even when other enabled
//! rules depend on it (issue #53).

use sheetrs::reader::read_workbook;
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

/// Lints `path` with the default rules, minus `disabled`, and returns the formatted violations.
fn lint(path: &str, disabled: Option<RuleId>) -> Vec<(RuleId, String)> {
    let mut config = LinterConfig::default();
    if let Some(id) = disabled {
        config.global.disabled_rules.insert(id.to_string());
    }
    let workbook = read_workbook(path).unwrap();
    let ctx = FormatContext {
        workbook: &workbook,
    };
    Linter::with_config(config)
        .lint_workbook(&workbook)
        .unwrap()
        .iter()
        .map(|v| {
            let line = format!("{:?} {}", v.scope, v.format_message(&ctx));
            (v.rule_id, line)
        })
        .collect()
}

fn assert_disabling_provider_keeps_dependents(path: &str) {
    let baseline = lint(path, None);
    for &(provider, dependents) in DEPENDENCIES {
        let without_provider = lint(path, Some(provider));
        assert!(
            !without_provider.iter().any(|(id, _)| *id == provider),
            "{provider} is disabled but still reported in {path}"
        );
        let expected: Vec<_> = baseline.iter().filter(|(id, _)| *id != provider).collect();
        let actual: Vec<_> = without_provider.iter().collect();
        assert_eq!(
            actual, expected,
            "disabling {provider} changed other rules' output in {path}"
        );
        assert!(
            baseline.iter().any(|(id, _)| dependents.contains(id)),
            "no violation of the rules depending on {provider} in {path}"
        );
    }
}

#[test]
fn test_disabling_provider_keeps_dependents_xlsx() {
    assert_disabling_provider_keeps_dependents(TEST_XLSX);
}

#[test]
fn test_disabling_provider_keeps_dependents_ods() {
    assert_disabling_provider_keeps_dependents(TEST_ODS);
}
