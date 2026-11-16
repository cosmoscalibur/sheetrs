//! Configuration system for linter rules

use crate::error::{Error, Result};
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::Path;

/// Main linter configuration.
///
/// A `LinterConfig` is always valid: the only ways to get one are [`Default`] and parsing
/// ([`from_toml`](Self::from_toml), [`from_file`](Self::from_file), or deserializing it), and
/// parsing rejects what the linter cannot honour: a rule or category it does not know (a
/// misspelled rule would otherwise be ignored without a word), `ALL` among the disabled rules, a
/// parameter outside its range, an unknown feature. Nothing downstream, no executable and no
/// binding, validates it again.
#[derive(Debug, Clone, Serialize, Default)]
pub struct LinterConfig {
    pub(crate) global: GlobalConfig,
    pub(crate) sheets: HashMap<String, SheetConfig>,
}

/// The shape of a configuration before it is known to be valid; only [`LinterConfig`]'s
/// `Deserialize` sees it.
#[derive(Deserialize)]
struct RawLinterConfig {
    #[serde(default)]
    global: GlobalConfig,
    #[serde(default)]
    sheets: HashMap<String, SheetConfig>,
}

impl<'de> Deserialize<'de> for LinterConfig {
    fn deserialize<D: serde::Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        Self::from_raw(RawLinterConfig::deserialize(deserializer)?)
            .map_err(serde::de::Error::custom)
    }
}

impl LinterConfig {
    /// Load and validate the configuration in a TOML file.
    ///
    /// # Errors
    ///
    /// When the file cannot be read, or its content is not a valid configuration (see
    /// [`from_toml`](Self::from_toml)).
    pub fn from_file(path: impl AsRef<Path>) -> Result<Self> {
        let path = path.as_ref();
        let content = fs::read_to_string(path).map_err(|source| Error::File {
            path: path.to_path_buf(),
            source,
        })?;
        Self::from_toml(&content)
    }

    /// Parse and validate a configuration in TOML.
    ///
    /// # Errors
    ///
    /// When the text is not valid TOML of the configuration, names a rule or category the linter
    /// does not know, puts `ALL` among the disabled rules, or gives a parameter a value outside
    /// its range.
    pub fn from_toml(content: &str) -> Result<Self> {
        Self::from_raw(toml::from_str(content)?)
    }

    fn from_raw(RawLinterConfig { global, sheets }: RawLinterConfig) -> Result<Self> {
        let config = Self { global, sheets };
        config.validate_rules(&crate::rules::registry::valid_tokens())?;
        Ok(config)
    }

    /// Check if a rule is enabled globally
    #[must_use]
    pub fn is_rule_enabled(&self, rule_id: &str) -> bool {
        // If explicitly disabled, return false
        if self
            .global
            .disabled_rules
            .iter()
            .any(|selector| matches_rule_selector(selector, rule_id))
        {
            return false;
        }

        // Check enabled rules
        if self.global.enabled_rules.is_empty() {
            // Default behavior if nothing enabled: all enabled (subject to default_active)
            return true;
        }

        // If enabled_rules has entries, check if matches
        self.global
            .enabled_rules
            .iter()
            .any(|selector| matches_rule_selector(selector, rule_id))
    }

    /// Check if a rule is enabled for a specific sheet
    #[must_use]
    pub fn is_rule_enabled_for_sheet(&self, rule_id: &str, sheet_name: &str) -> bool {
        // First check global setting
        if !self.is_rule_enabled(rule_id) {
            return false;
        }

        // Check sheet-specific override
        if let Some(sheet_config) = self.sheets.get(sheet_name)
            && sheet_config
                .disabled_rules
                .iter()
                .any(|selector| matches_rule_selector(selector, rule_id))
        {
            return false;
        }

        true
    }

    /// Validate the configuration against a set of valid tokens (parsing does it with the rules
    /// the linter has)
    pub(crate) fn validate_rules(&self, valid_tokens: &HashSet<String>) -> Result<()> {
        // Validate global disabled rules (NO "ALL" allowed)
        for rule in &self.global.disabled_rules {
            if rule == "ALL" {
                return Err(Error::AllDisabled);
            }
            if !valid_tokens.contains(rule) {
                return Err(Error::UnknownRule {
                    token: rule.clone(),
                    place: "global disabled_rules".into(),
                });
            }
        }

        // Validate global enabled rules
        for rule in &self.global.enabled_rules {
            if !valid_tokens.contains(rule) {
                return Err(Error::UnknownRule {
                    token: rule.clone(),
                    place: "global enabled_rules".into(),
                });
            }
        }

        // Validate sheet-specific disabled rules
        for (sheet_name, sheet_config) in &self.sheets {
            for rule in &sheet_config.disabled_rules {
                if !valid_tokens.contains(rule) {
                    return Err(Error::UnknownRule {
                        token: rule.clone(),
                        place: format!("sheet '{sheet_name}' disabled_rules"),
                    });
                }
            }
        }

        self.validate_params()
    }

    /// Parameters that must be a whole number of at least 1: with 0 or less the rule that reads
    /// them would report every case (REF310) or none.
    const POSITIVE_INT_PARAMS: &'static [&'static str] = &["ref310_min_adjacent_data"];

    /// Validate the values of the parameters the rules read.
    fn validate_params(&self) -> Result<()> {
        let scopes = std::iter::once(("global".to_string(), &self.global.params)).chain(
            self.sheets
                .iter()
                .map(|(name, s)| (format!("sheet '{name}'"), &s.params)),
        );
        for (scope, params) in scopes {
            for key in Self::POSITIVE_INT_PARAMS {
                if let Some(value) = params.get(*key)
                    && value.as_integer().is_none_or(|n| n < 1)
                {
                    return Err(Error::NotPositive {
                        key,
                        scope,
                        value: value.to_string(),
                    });
                }
            }
        }
        if let Some(value) = self.global.params.get(Self::LOCALE_KEY)
            && value
                .as_str()
                .is_none_or(|tag| tag.parse::<sheetrs_io::locale::LocaleTag>().is_err())
        {
            return Err(Error::Locale {
                value: value.to_string(),
            });
        }
        for (key, expected, valid) in [
            (
                "now",
                "a number (the serial number of NOW())",
                (|v: &toml::Value| v.is_integer() || v.is_float()) as fn(&toml::Value) -> bool,
            ),
            ("seed", "a non-negative integer", |v: &toml::Value| {
                v.as_integer().is_some_and(|n| n >= 0)
            }),
        ] {
            if let Some(value) = self.global.params.get(key)
                && !valid(value)
            {
                return Err(Error::Volatile {
                    key,
                    expected,
                    value: value.to_string(),
                });
            }
        }
        for key in Self::FEATURE_KEYS {
            let Some(value) = self.global.params.get(*key) else {
                continue;
            };
            let names: Option<Vec<&str>> = value
                .as_array()
                .and_then(|a| a.iter().map(|v| v.as_str()).collect());
            let Some(names) = names else {
                return Err(Error::FeatureList {
                    key,
                    known: Self::FEATURES.join(", "),
                });
            };
            for name in names {
                if !Self::FEATURES
                    .iter()
                    .any(|f| f.eq_ignore_ascii_case(name.trim()))
                {
                    return Err(Error::UnknownFeature {
                        name: name.to_string(),
                        key,
                        known: Self::FEATURES.join(", "),
                    });
                }
            }
        }
        Ok(())
    }

    /// The key of the locale of the workbook (read by `sheetrs query` and `sheetrs edit`; the rules do
    /// not use it).
    const LOCALE_KEY: &'static str = "locale";

    /// Compatibility features a rule may honour, by name (`enable_features`).
    pub const FEATURES: &'static [&'static str] = &["LONGTEXT"];

    /// Keys under which the feature list can be written.
    const FEATURE_KEYS: &'static [&'static str] = &["enable_features", "enable-features"];

    /// Whether the global parameter `key` is set.
    #[must_use]
    pub fn has_param(&self, key: &str) -> bool {
        self.global.params.contains_key(key)
    }

    /// Whether the feature `name` is listed in `enable_features` (any case). Every feature is off
    /// unless listed.
    #[must_use]
    pub fn feature_enabled(&self, name: &str) -> bool {
        Self::FEATURE_KEYS.iter().any(|key| {
            self.param_array(key, None)
                .is_some_and(|list| list.iter().any(|n| n.trim().eq_ignore_ascii_case(name)))
        })
    }

    /// Get a parameter value with fallback chain: sheet -> global
    #[must_use]
    pub fn param_int(&self, key: &str, sheet_name: Option<&str>) -> Option<i64> {
        // Try sheet-specific first
        if let Some(sheet) = sheet_name.and_then(|name| self.sheets.get(name))
            && let Some(value) = sheet.params.get(key).and_then(toml::Value::as_integer)
        {
            return Some(value);
        }

        // Try global
        self.global
            .params
            .get(key)
            .and_then(toml::Value::as_integer)
    }

    /// Get a parameter value as string with fallback chain: sheet -> global
    #[must_use]
    pub fn param_str<'a>(&'a self, key: &str, sheet_name: Option<&str>) -> Option<&'a str> {
        // Try sheet-specific first
        if let Some(sheet) = sheet_name.and_then(|name| self.sheets.get(name))
            && let Some(value) = sheet.params.get(key).and_then(|v| v.as_str())
        {
            return Some(value);
        }

        // Try global
        self.global.params.get(key).and_then(|v| v.as_str())
    }

    /// Get a parameter value as array with fallback chain: sheet -> global
    #[must_use]
    pub fn param_array(&self, key: &str, sheet_name: Option<&str>) -> Option<Vec<String>> {
        // Try sheet-specific first
        if let Some(sheet) = sheet_name.and_then(|name| self.sheets.get(name))
            && let Some(arr) = sheet.params.get(key).and_then(|v| {
                v.as_array().map(|arr| {
                    arr.iter()
                        .filter_map(|item| item.as_str().map(std::string::ToString::to_string))
                        .collect()
                })
            })
        {
            return Some(arr);
        }

        // Try global
        self.global.params.get(key).and_then(|v| {
            v.as_array().map(|arr| {
                arr.iter()
                    .filter_map(|item| item.as_str().map(std::string::ToString::to_string))
                    .collect()
            })
        })
    }

    /// Get a parameter value as float array with fallback chain: sheet -> global
    #[must_use]
    pub fn param_float_array(&self, key: &str, sheet_name: Option<&str>) -> Option<Vec<f64>> {
        // Try sheet-specific first
        if let Some(sheet) = sheet_name.and_then(|name| self.sheets.get(name))
            && let Some(arr) = sheet.params.get(key).and_then(|v| {
                v.as_array().map(|arr| {
                    arr.iter()
                        .filter_map(|item| item.as_float().or(item.as_integer().map(|i| i as f64)))
                        .collect()
                })
            })
        {
            return Some(arr);
        }

        // Try global
        self.global.params.get(key).and_then(|v| {
            v.as_array().map(|arr| {
                arr.iter()
                    .filter_map(|item| item.as_float().or(item.as_integer().map(|i| i as f64)))
                    .collect()
            })
        })
    }

    /// Get a parameter value as boolean with fallback chain: sheet -> global
    #[must_use]
    pub fn param_bool(&self, key: &str, sheet_name: Option<&str>) -> Option<bool> {
        // Try sheet-specific first
        if let Some(sheet) = sheet_name.and_then(|name| self.sheets.get(name))
            && let Some(value) = sheet.params.get(key).and_then(toml::Value::as_bool)
        {
            return Some(value);
        }

        // Try global
        self.global.params.get(key).and_then(toml::Value::as_bool)
    }
}

/// Global configuration
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct GlobalConfig {
    /// List of enabled rules (empty means all enabled)
    #[serde(default)]
    pub enabled_rules: HashSet<String>,
    /// List of disabled rules
    #[serde(default)]
    pub disabled_rules: HashSet<String>,
    #[serde(flatten)]
    pub params: HashMap<String, toml::Value>,
}

fn matches_rule_selector(selector: &str, rule_id: &str) -> bool {
    if selector == "ALL" {
        return true;
    }
    rule_id == selector || rule_id.starts_with(selector)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_rule_matching() {
        assert!(matches_rule_selector("ALL", "ERR001"));
        assert!(matches_rule_selector("ERR", "ERR001"));
        assert!(matches_rule_selector("ERR001", "ERR001"));
        assert!(!matches_rule_selector("ERR", "SEC001"));
        assert!(!matches_rule_selector("SEC001", "ERR001"));
    }

    #[test]
    fn test_config_activation() {
        let mut config = LinterConfig::default();

        // Default: everything matches if enabled_rules empty
        assert!(config.is_rule_enabled("ERR001"));

        // Explicit disable
        config.global.disabled_rules.insert("ERR001".to_string());
        assert!(!config.is_rule_enabled("ERR001"));
        assert!(config.is_rule_enabled("SEC001"));

        // Prefix disable
        config.global.disabled_rules.clear();
        config.global.disabled_rules.insert("SEC".to_string());
        assert!(!config.is_rule_enabled("SEC001"));
        assert!(config.is_rule_enabled("ERR001"));

        // Specific enable
        config.global.disabled_rules.clear();
        config.global.enabled_rules.insert("ERR".to_string());
        assert!(config.is_rule_enabled("ERR001"));
        assert!(!config.is_rule_enabled("SEC001"));
    }

    #[test]
    fn test_validation() {
        let config = LinterConfig::default();
        let mut tokens = HashSet::new();
        tokens.insert("ALL".to_string());
        tokens.insert("ERR".to_string());
        tokens.insert("ERR001".to_string());

        // Valid
        assert!(config.validate_rules(&tokens).is_ok());

        // Invalid Global Disable ALL
        let mut bad_config = config.clone();
        bad_config.global.disabled_rules.insert("ALL".to_string());
        assert!(bad_config.validate_rules(&tokens).is_err());

        // Invalid Token
        let mut bad_config = config.clone();
        bad_config.global.enabled_rules.insert("XYZ".to_string());
        assert!(bad_config.validate_rules(&tokens).is_err());

        // Invalid Sheet Disable Token
        let mut bad_config = config.clone();
        let mut sheet_config = SheetConfig::default();
        sheet_config.disabled_rules.insert("ABC".to_string());
        bad_config.sheets.insert("Sheet1".to_string(), sheet_config);
        assert!(bad_config.validate_rules(&tokens).is_err());
    }

    #[test]
    fn parsing_rejects_what_the_linter_cannot_honour() {
        for bad in [
            "[global]\ndisabled_rules = [\"ERR1O2\"]\n",
            "[global]\nenabled_rules = [\"NOPE\"]\n",
            "[global]\ndisabled_rules = [\"ALL\"]\n",
            "[sheets.Data]\ndisabled_rules = [\"ERR1O2\"]\n",
        ] {
            let error = LinterConfig::from_toml(bad).expect_err(bad);
            assert!(error.to_string().contains("Configuration error"), "{error}");
            // The same through any serde route: there is no way around the check.
            assert!(toml::from_str::<LinterConfig>(bad).is_err(), "{bad}");
        }
        assert!(
            LinterConfig::from_toml("[global]\ndisabled_rules = [\"ERR102\", \"CPX\"]\n").is_ok()
        );
        assert!(LinterConfig::from_toml("").is_ok());
    }

    #[test]
    fn a_configuration_error_says_what_is_wrong_with_a_type() {
        use crate::Error;
        let error =
            LinterConfig::from_toml("[global]\ndisabled_rules = [\"ERR1O2\"]\n").unwrap_err();
        assert!(
            matches!(&error, Error::UnknownRule { token, place } if token == "ERR1O2" && place.contains("disabled")),
            "{error:?}"
        );
        assert!(matches!(
            LinterConfig::from_toml("[global]\ndisabled_rules = [\"ALL\"]\n"),
            Err(Error::AllDisabled)
        ));
        assert!(matches!(
            LinterConfig::from_toml("[global]\nref310_min_adjacent_data = 0\n"),
            Err(Error::NotPositive { .. })
        ));
        assert!(matches!(
            LinterConfig::from_toml("[global]\nenable_features = [\"NOPE\"]\n"),
            Err(Error::UnknownFeature { .. })
        ));
        for bad in ["locale = 3", "locale = \"es CO\"", "locale = \"\""] {
            assert!(
                matches!(
                    LinterConfig::from_toml(&format!("[global]\n{bad}\n")),
                    Err(Error::Locale { .. })
                ),
                "{bad}"
            );
        }
        assert!(LinterConfig::from_toml("[global]\nlocale = \"es-CO\"\n").is_ok());
        for bad in ["now = \"today\"", "seed = -1", "seed = 1.5", "seed = \"7\""] {
            assert!(
                matches!(
                    LinterConfig::from_toml(&format!("[global]\n{bad}\n")),
                    Err(Error::Volatile { .. })
                ),
                "{bad}"
            );
        }
        assert!(LinterConfig::from_toml("[global]\nnow = 45366.75\nseed = 7\n").is_ok());
        assert!(matches!(
            LinterConfig::from_toml("global = 3"),
            Err(Error::Syntax(_))
        ));
        assert!(matches!(
            LinterConfig::from_file("does/not/exist.toml"),
            Err(Error::File { .. })
        ));
    }
}

/// Sheet-specific configuration
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SheetConfig {
    /// Rules disabled for this sheet
    #[serde(default)]
    pub disabled_rules: HashSet<String>,
    #[serde(flatten)]
    pub params: HashMap<String, toml::Value>,
}
