//! Build configuration file (`sheetc.toml`), in the manner of rustc's: every
//! key has a command-line counterpart that takes precedence.
//!
//! ```toml
//! settings = "settings"   # sheet that declares inputs, outputs and metadata
//! outputs = "declared"    # declared | named | any
//! fold = true
//!
//! [locale]                # the user's regional settings, fixed at build time
//! tag = "es-CO"
//! decimal = ","
//! group = "."
//! two_digit_year_start = 1930
//! ```
//!
//! The file only holds what the workbook cannot say itself: the date system,
//! iteration, search settings and dialect come from the document.

use crate::compile::{BindingSpec, CompileError, QueryLevel, Spec};
use crate::locale::{self, LocaleOverride};
use crate::model::Model;
use crate::settings;
use serde::Deserialize;
use std::path::{Path, PathBuf};

pub const FILE_NAME: &str = "sheetc.toml";

#[derive(Debug, Clone, Default, Deserialize)]
#[serde(deny_unknown_fields)]
pub struct Config {
    /// Name of the settings sheet (default `settings`, optional when absent).
    pub settings: Option<String>,
    /// `declared`, `named` or `any`.
    pub outputs: Option<String>,
    /// Evaluate input-independent formulas at build time (default true).
    pub fold: Option<bool>,
    /// Fixed value of `NOW()`, as a serial number.
    pub now: Option<f64>,
    pub seed: Option<u64>,
    pub locale: Option<LocaleOverride>,
}

impl LocaleOverride {
    /// Fields set on `self` win over `base`.
    fn over(self, base: Self) -> Self {
        Self {
            tag: self.tag.or(base.tag),
            decimal: self.decimal.or(base.decimal),
            group: self.group.or(base.group),
            true_text: self.true_text.or(base.true_text),
            false_text: self.false_text.or(base.false_text),
            currency: self.currency.or(base.currency),
            two_digit_year_start: self.two_digit_year_start.or(base.two_digit_year_start),
            date_order: self.date_order.or(base.date_order),
        }
    }
}

impl Config {
    /// Settings of `self` (typically the command line) take precedence over
    /// `base` (the file).
    #[must_use]
    pub fn over(self, base: Self) -> Self {
        Self {
            settings: self.settings.or(base.settings),
            outputs: self.outputs.or(base.outputs),
            fold: self.fold.or(base.fold),
            now: self.now.or(base.now),
            seed: self.seed.or(base.seed),
            locale: match (self.locale, base.locale) {
                (Some(a), Some(b)) => Some(a.over(b)),
                (a, b) => a.or(b),
            },
        }
    }

    /// The compiler's [`Spec`] for `model`: the settings sheet supplies inputs,
    /// outputs and metadata; this configuration the rest.
    pub fn spec(&self, model: &Model) -> Result<Spec, CompileError> {
        let query = match self.outputs.as_deref() {
            None => QueryLevel::Declared,
            Some(s) => QueryLevel::parse(s).ok_or_else(|| {
                CompileError(format!("outputs must be declared, named or any, got `{s}`"))
            })?,
        };
        let mut spec = Spec {
            fold: self.fold.unwrap_or(true),
            now: self.now,
            rand_seed: self.seed,
            query,
            ..Spec::default()
        };
        let loc = self.locale.clone().unwrap_or_default();
        let (l, notes) =
            locale::from_tag(loc.tag.as_deref().unwrap_or("en"), &loc).map_err(CompileError)?;
        spec.locale = Some(l);
        spec.notes = notes;

        let name = self.settings.as_deref().unwrap_or(settings::DEFAULT_SHEET);
        if let Some(s) = settings::read(model, name, self.settings.is_some())? {
            if s.outputs.is_empty() && query == QueryLevel::Declared {
                return Err(CompileError(format!(
                    "settings sheet `{name}` declares no output; add `output` rows or build with outputs = \"any\""
                )));
            }
            spec.inputs = s.inputs.iter().map(|t| BindingSpec::parse(t)).collect();
            spec.outputs = s.outputs.iter().map(|t| BindingSpec::parse(t)).collect();
            spec.meta = s.meta;
            spec.skip_sheet = Some(s.sheet);
        }
        Ok(spec)
    }
}

pub fn parse(text: &str, origin: &str) -> Result<Config, String> {
    toml::from_str(text).map_err(|e| format!("{origin}: {e}"))
}

pub fn load(path: &Path) -> Result<Config, String> {
    let text = std::fs::read_to_string(path).map_err(|e| format!("{}: {e}", path.display()))?;
    parse(&text, &path.display().to_string())
}

/// The configuration next to the workbook, when there is one. Only that
/// directory is searched, so a build does not depend on where it is started.
pub fn discover(workbook: &Path) -> Option<PathBuf> {
    let dir = workbook
        .parent()
        .filter(|d| !d.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    let p = dir.join(FILE_NAME);
    p.is_file().then_some(p)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reads_every_section() {
        let c = parse("settings = \"cfg\"\nouts = 1\n", "t");
        assert!(c.is_err(), "unknown keys are rejected");
        let c = parse(
            "settings = \"cfg\"\noutputs = \"any\"\nfold = false\n[locale]\ntag = \"es-CO\"\ndecimal = \",\"\n",
            "t",
        )
        .unwrap();
        assert_eq!(c.settings.as_deref(), Some("cfg"));
        assert_eq!(c.outputs.as_deref(), Some("any"));
        let l = c.locale.unwrap();
        assert_eq!((l.tag.as_deref(), l.decimal), (Some("es-CO"), Some(',')));
        assert!(parse("[locale]\nbogus = 1\n", "t").is_err());
    }
}
