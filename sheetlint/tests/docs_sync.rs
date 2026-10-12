//! Keeps the rule reference in `README.md` and the `sheetlint.toml` example in
//! sync with the rule registry.

use sheetrs::LinterConfig;
use sheetrs::rules::registry::get_all_rule_metadata;
use std::collections::BTreeSet;

const README: &str = include_str!("../README.md");
const EXAMPLE_CONFIG: &str = include_str!("../../sheetlint.toml");

/// Rows of the rule reference tables, as `(rule ID, Params cell)`.
fn rule_rows() -> Vec<(&'static str, &'static str)> {
    README
        .lines()
        .filter_map(|line| {
            let line = line.strip_prefix("| **")?;
            let (id, rest) = line.split_once("** |")?;
            let params = rest.trim_end_matches('|').rsplit('|').next()?.trim();
            Some((id, params))
        })
        .collect()
}

#[test]
fn readme_lists_each_registered_rule_once() {
    let documented: Vec<&str> = rule_rows().into_iter().map(|(id, _)| id).collect();
    let documented_set: BTreeSet<&str> = documented.iter().copied().collect();
    assert_eq!(
        documented.len(),
        documented_set.len(),
        "README lists a rule ID more than once"
    );

    let config = LinterConfig::default();
    let registered: BTreeSet<&str> = get_all_rule_metadata(&config)
        .iter()
        .map(|m| m.id.as_str())
        .collect();
    assert_eq!(documented_set, registered);
}

#[test]
fn example_config_lists_every_documented_param() {
    let config = LinterConfig::from_toml(EXAMPLE_CONFIG).expect("sheetlint.toml must parse");
    for (id, params) in rule_rows() {
        for param in params.split('`').skip(1).step_by(2) {
            assert!(
                config.global.params.contains_key(param),
                "{id} parameter `{param}` missing from sheetlint.toml"
            );
        }
    }
}
