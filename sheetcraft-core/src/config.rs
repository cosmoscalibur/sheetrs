//! Configuration system for linter rules

use anyhow::Result;
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};
use std::fs;
use std::path::Path;

/// Main linter configuration
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LinterConfig {
    #[serde(default)]
    pub global: GlobalConfig,
    #[serde(default)]
    pub sheets: HashMap<String, SheetConfig>,
}

impl LinterConfig {
    /// Load configuration from a TOML file
    pub fn from_file<P: AsRef<Path>>(path: P) -> Result<Self> {
        let content = fs::read_to_string(path)?;
        let config: LinterConfig = toml::from_str(&content)?;
        Ok(config)
    }

    /// Check if a rule is enabled globally
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
    pub fn is_rule_enabled_for_sheet(&self, rule_id: &str, sheet_name: &str) -> bool {
        // First check global setting
        if !self.is_rule_enabled(rule_id) {
            return false;
        }

        // Check sheet-specific override
        if let Some(sheet_config) = self.sheets.get(sheet_name) {
            if sheet_config
                .disabled_rules
                .iter()
                .any(|selector| matches_rule_selector(selector, rule_id))
            {
                return false;
            }
        }

        true
    }

    /// Validate the configuration against a set of valid tokens
    pub fn validate_rules(&self, valid_tokens: &HashSet<String>) -> Result<()> {
        // Validate global disabled rules (NO "ALL" allowed)
        for rule in &self.global.disabled_rules {
            if rule == "ALL" {
                anyhow::bail!("Configuration error: 'ALL' is not allowed in global disabled_rules");
            }
            if !valid_tokens.contains(rule) {
                anyhow::bail!(
                    "Configuration error: Unknown rule or category '{}' in global disabled_rules",
                    rule
                );
            }
        }

        // Validate global enabled rules
        for rule in &self.global.enabled_rules {
            if !valid_tokens.contains(rule) {
                anyhow::bail!(
                    "Configuration error: Unknown rule or category '{}' in global enabled_rules",
                    rule
                );
            }
        }

        // Validate sheet-specific disabled rules
        for (sheet_name, sheet_config) in &self.sheets {
            for rule in &sheet_config.disabled_rules {
                if !valid_tokens.contains(rule) {
                    anyhow::bail!(
                        "Configuration error: Unknown rule or category '{}' in sheet '{}' disabled_rules",
                        rule,
                        sheet_name
                    );
                }
            }
        }

        Ok(())
    }

    /// Get a parameter value with fallback chain: sheet -> global
    pub fn get_param_int(&self, key: &str, sheet_name: Option<&str>) -> Option<i64> {
        // Try sheet-specific first
        if let Some(sheet) = sheet_name.and_then(|name| self.sheets.get(name)) {
            if let Some(value) = sheet.params.get(key).and_then(|v| v.as_integer()) {
                return Some(value);
            }
        }

        // Try global
        self.global.params.get(key).and_then(|v| v.as_integer())
    }

    /// Get a parameter value as string with fallback chain: sheet -> global
    pub fn get_param_str<'a>(&'a self, key: &str, sheet_name: Option<&str>) -> Option<&'a str> {
        // Try sheet-specific first
        if let Some(sheet) = sheet_name.and_then(|name| self.sheets.get(name)) {
            if let Some(value) = sheet.params.get(key).and_then(|v| v.as_str()) {
                return Some(value);
            }
        }

        // Try global
        self.global.params.get(key).and_then(|v| v.as_str())
    }

    /// Get a parameter value as array with fallback chain: sheet -> global
    pub fn get_param_array(&self, key: &str, sheet_name: Option<&str>) -> Option<Vec<String>> {
        // Try sheet-specific first
        if let Some(sheet) = sheet_name.and_then(|name| self.sheets.get(name)) {
            if let Some(arr) = sheet.params.get(key).and_then(|v| {
                v.as_array().map(|arr| {
                    arr.iter()
                        .filter_map(|item| item.as_str().map(|s| s.to_string()))
                        .collect()
                })
            }) {
                return Some(arr);
            }
        }

        // Try global
        self.global.params.get(key).and_then(|v| {
            v.as_array().map(|arr| {
                arr.iter()
                    .filter_map(|item| item.as_str().map(|s| s.to_string()))
                    .collect()
            })
        })
    }

    /// Get a parameter value as float array with fallback chain: sheet -> global
    pub fn get_param_float_array(&self, key: &str, sheet_name: Option<&str>) -> Option<Vec<f64>> {
        // Try sheet-specific first
        if let Some(sheet) = sheet_name.and_then(|name| self.sheets.get(name)) {
            if let Some(arr) = sheet.params.get(key).and_then(|v| {
                v.as_array().map(|arr| {
                    arr.iter()
                        .filter_map(|item| item.as_float().or(item.as_integer().map(|i| i as f64)))
                        .collect()
                })
            }) {
                return Some(arr);
            }
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
    pub fn get_param_bool(&self, key: &str, sheet_name: Option<&str>) -> Option<bool> {
        // Try sheet-specific first
        if let Some(sheet) = sheet_name.and_then(|name| self.sheets.get(name)) {
            if let Some(value) = sheet.params.get(key).and_then(|v| v.as_bool()) {
                return Some(value);
            }
        }

        // Try global
        self.global.params.get(key).and_then(|v| v.as_bool())
    }
}

impl Default for LinterConfig {
    fn default() -> Self {
        Self {
            global: GlobalConfig::default(),
            sheets: HashMap::new(),
        }
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
