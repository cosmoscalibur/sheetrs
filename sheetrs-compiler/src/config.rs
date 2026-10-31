//! Build configuration file (`sheetc.toml`), in the manner of rustc's: every
//! key has a command-line counterpart that takes precedence.
//!
//! ```toml
//! settings = "settings"   # sheet that declares inputs, outputs and metadata
//! outputs = "declared"    # declared | named | any
//! fold = true
//! enable_features = ["LONGTEXT"]   # compatibility features, all off unless listed
//! strict = true           # refuse workbooks with macros (false: warn)
//!
//! [locale]                # the user's regional settings, fixed at build time
//! tag = "es-CO"
//! decimal = ","
//! group = "."
//! date_order = "dmy"
//! two_digit_year_start = 1930
//! ```
//!
//! The file only holds what the workbook cannot say itself. Whatever the document does not
//! specify (how `TRUE()` is stored, what an error of a function's domain is, text
//! collation, ...) is LibreOffice's behavior. The date system, iteration and search
//! settings come from the document.

use crate::compile::{BindingSpec, CompileError, Features, QueryLevel, Spec};
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
    /// Refuse what the compiler cannot reproduce (macros); default true. When
    /// false they are reported and the build goes on.
    pub strict: Option<bool>,
    /// Fixed value of `NOW()`, as a serial number.
    /// Engine whose behavior is reproduced: `libreoffice` (default) or `excel`.
    pub now: Option<f64>,
    pub seed: Option<u64>,
    /// What `INFO` (`system`, `osversion`, `release`) and `CELL("filename")` (`directory`,
    /// `filename`) say; absent, the machine the program runs on.
    pub host: Option<sheetrs_engine::options::Host>,
    /// Compatibility features on, by name (`LONGTEXT`). Every feature is off unless listed.
    #[serde(alias = "enable-features")]
    pub enable_features: Option<Vec<String>>,
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
            am: self.am.or(base.am),
            pm: self.pm.or(base.pm),
            currency: self.currency.or(base.currency),
            currency_after: self.currency_after.or(base.currency_after),
            currency_parens: self.currency_parens.or(base.currency_parens),
            two_digit_year_start: self.two_digit_year_start.or(base.two_digit_year_start),
            date_order: self.date_order.or(base.date_order),
            date_patterns: self.date_patterns.or(base.date_patterns),
            months: self.months.or(base.months),
            months_abbr: self.months_abbr.or(base.months_abbr),
            days: self.days.or(base.days),
            days_abbr: self.days_abbr.or(base.days_abbr),
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
            strict: self.strict.or(base.strict),
            now: self.now.or(base.now),
            seed: self.seed.or(base.seed),
            host: match (self.host, base.host) {
                (Some(a), Some(b)) => Some(a.over(b)),
                (a, b) => a.or(b),
            },
            enable_features: self.enable_features.or(base.enable_features),
            locale: match (self.locale, base.locale) {
                (Some(a), Some(b)) => Some(a.over(b)),
                (a, b) => a.or(b),
            },
        }
    }

    /// The compiler's [`Spec`] for `model`: the settings sheet supplies inputs,
    /// outputs and metadata; this configuration the rest.
    pub fn spec(&self, model: &Model) -> Result<Spec, CompileError> {
        let mut spec = self.base_spec(model)?;
        let name = self.settings.as_deref().unwrap_or(settings::DEFAULT_SHEET);
        if let Some(s) = settings::read(model, name, self.settings.is_some())? {
            if s.outputs.is_empty() && spec.query == QueryLevel::Declared {
                return Err(CompileError(format!(
                    "settings sheet `{name}` declares no output; add `output` rows or build with outputs = \"any\""
                )));
            }
            spec.inputs = s
                .inputs
                .iter()
                .map(|(n, t)| BindingSpec::named(n, t))
                .collect();
            spec.outputs = s
                .outputs
                .iter()
                .map(|(n, t)| BindingSpec::named(n, t))
                .collect();
            spec.meta = s.meta;
            spec.skip_sheet = Some(s.sheet);
        }
        Ok(spec)
    }

    /// Everything of the spec that does not come from the settings sheet: the
    /// level, folding, clock, seed, locale and the macro policy.
    pub fn base_spec(&self, model: &Model) -> Result<Spec, CompileError> {
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
            host: self.host.clone().unwrap_or_default(),
            features: match &self.enable_features {
                None => Features::default(),
                Some(names) => Features::parse(names).map_err(CompileError)?,
            },
            query,
            ..Spec::default()
        };
        let loc = self.locale.clone().unwrap_or_default();
        let (l, notes) =
            locale::from_tag(loc.tag.as_deref().unwrap_or("en"), &loc).map_err(CompileError)?;
        spec.locale = Some(l);
        spec.notes = notes;
        if model.macros {
            if self.strict.unwrap_or(true) {
                return Err(CompileError(
                    "the workbook contains macros, which are never run, so what depends on them would differ; \
                     build with --no-strict (strict = false) to compile anyway"
                        .into(),
                ));
            }
            spec.notes.push(
                "warning: the workbook contains macros; they are ignored (strict mode is off)"
                    .into(),
            );
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

#[cfg(test)]
mod feature_tests {
    use super::*;

    #[test]
    fn enable_features_are_off_unless_listed() {
        let model = Model::default();
        let spec = |toml: &str| parse(toml, "t").unwrap().base_spec(&model);
        assert!(!spec("").unwrap().features.longtext);
        assert!(
            spec("enable_features = [\"longtext\"]")
                .unwrap()
                .features
                .longtext
        );
        assert!(
            spec("enable-features = [\"LONGTEXT\"]")
                .unwrap()
                .features
                .longtext
        );
        assert!(!spec("enable_features = []").unwrap().features.longtext);
        let err = spec("enable_features = [\"NOPE\"]").unwrap_err();
        assert!(err.0.contains("unknown feature `NOPE`"), "{err}");
    }
}
