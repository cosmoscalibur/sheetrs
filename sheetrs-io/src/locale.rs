//! The locale of a workbook: the regional settings (decimal and group separators, the names of months,
//! the order of a date) that give a text to a number, and that read a number from a text.
//!
//! The locale is a property of the [`Book`](crate::Book). It is the first of these that is set:
//!
//! 1. the one given to the book ([`Book::with_locale`](crate::Book::with_locale); the command lines,
//!    Python and WebAssembly give it with the same name, `locale`), which an option of a single call
//!    still overrides;
//! 2. the configuration file (`locale` in the `[global]` table of `sheetrs.toml`, see
//!    [`from_config`]; the tag of the `[locale]` table of `sheetc.toml`): the command lines read it and
//!    give it to the book as the one above, so a flag wins over it;
//! 3. the default of the format ([`Book::default_locale`](crate::Book::default_locale)): an ODS says
//!    it (the language and country of the default cell style), an XLSX does not, so it is the
//!    locale of the system ([`system`]); `en-US` when the system has none, and in WebAssembly, which has
//!    no system.
//!
//! This module only knows what a tag is. What a tag means (the separators, the names) is the table of
//! the engine, which each tool consults when it needs the locale.

use crate::error::{Error, Result};
use std::fmt;
use std::str::FromStr;

/// A locale tag: a language and, optionally, a region (`es`, `es-CO`). A subset of BCP 47, the one
/// `LibreOffice` calls a locale: two or three letters of language and two letters (or three digits)
/// of region; a script subtag (`sr-Latn-RS`) is skipped. The language is lower case and the region
/// upper case, whatever was written, and `_` and `-` both separate them (`es_CO`).
///
/// ```
/// use sheetrs_io::locale::LocaleTag;
///
/// let tag: LocaleTag = "es_co".parse().unwrap();
/// assert_eq!(tag.to_string(), "es-CO");
/// assert_eq!((tag.language(), tag.region()), ("es", Some("CO")));
/// assert!("not a locale".parse::<LocaleTag>().is_err());
/// ```
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct LocaleTag {
    language: String,
    region: Option<String>,
}

impl LocaleTag {
    /// The language, in lower case (`es`).
    #[must_use]
    pub fn language(&self) -> &str {
        &self.language
    }

    /// The region, in upper case (`CO`), when the tag has one.
    #[must_use]
    pub fn region(&self) -> Option<&str> {
        self.region.as_deref()
    }
}

impl FromStr for LocaleTag {
    type Err = Error;

    fn from_str(text: &str) -> Result<Self> {
        let bad = || Error::InvalidLocale(text.to_string());
        let mut parts = text.trim().split(['-', '_']);
        let language = parts
            .next()
            .filter(|l| (2..=3).contains(&l.len()) && l.chars().all(|c| c.is_ascii_alphabetic()))
            .ok_or_else(bad)?
            .to_ascii_lowercase();
        let mut region = None;
        for part in parts {
            let is_script = part.len() == 4 && part.chars().all(|c| c.is_ascii_alphabetic());
            let is_region = (part.len() == 2 && part.chars().all(|c| c.is_ascii_alphabetic()))
                || (part.len() == 3 && part.chars().all(|c| c.is_ascii_digit()));
            if region.is_none() && is_script {
                continue;
            }
            if region.is_some() || !is_region {
                return Err(bad());
            }
            region = Some(part.to_ascii_uppercase());
        }
        Ok(Self { language, region })
    }
}

impl fmt::Display for LocaleTag {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.region {
            Some(region) => write!(f, "{}-{region}", self.language),
            None => f.write_str(&self.language),
        }
    }
}

/// The locale a process environment asks for, as `LibreOffice` takes it on Linux: the first of
/// `LC_ALL`, `LC_CTYPE` and `LANG` that is set and not empty (`LibreOffice` 26.8 ignores
/// `LC_MESSAGES`, `LC_NUMERIC`, `LC_TIME`, `LC_MONETARY` and `LANGUAGE`: they only change its
/// interface). The value is read as POSIX writes it (`es_CO.UTF-8`, `de_DE@euro`); `C` and `POSIX`
/// ask for no locale, which is `None`, and so does a value that is not a locale name.
///
/// `get` reads a variable: [`system`] gives it the real environment.
///
/// ```
/// use sheetrs_io::locale::from_environment;
///
/// let env = |name: &str| match name {
///     "LANG" => Some("es_CO.UTF-8".to_string()),
///     "LC_ALL" => Some(String::new()),
///     _ => None,
/// };
/// assert_eq!(from_environment(env).unwrap().to_string(), "es-CO");
/// ```
#[must_use]
pub fn from_environment(get: impl Fn(&str) -> Option<String>) -> Option<LocaleTag> {
    let value = ["LC_ALL", "LC_CTYPE", "LANG"]
        .iter()
        .find_map(|name| get(name).filter(|v| !v.is_empty()))?;
    // `language[_TERRITORY][.codeset][@modifier]`
    let name = value.split(['.', '@']).next().unwrap_or_default();
    name.parse().ok()
}

/// The locale of the system the program runs on (see [`from_environment`]); `None` when it has none
/// and always in WebAssembly, which has no system.
#[must_use]
pub fn system() -> Option<LocaleTag> {
    #[cfg(not(target_arch = "wasm32"))]
    {
        from_environment(|name| std::env::var(name).ok())
    }
    #[cfg(target_arch = "wasm32")]
    {
        None
    }
}

/// The `locale` of the `[global]` table of a configuration file (`sheetrs.toml`), `None` when it has
/// none. The other keys of the file belong to the tools that read them and are not looked at.
///
/// # Errors
///
/// [`Error::Config`] when the text is not TOML or `locale` is not a text, and
/// [`Error::InvalidLocale`] when it is not a locale tag: a misspelled locale is not ignored.
///
/// ```
/// use sheetrs_io::locale::from_config;
///
/// let tag = from_config("[global]\nlocale = \"es-CO\"\n").unwrap().unwrap();
/// assert_eq!(tag.to_string(), "es-CO");
/// assert!(from_config("[global]\nlocale = \"es CO\"\n").is_err());
/// assert!(from_config("[global]\n").unwrap().is_none());
/// ```
#[cfg(feature = "config")]
pub fn from_config(text: &str) -> Result<Option<LocaleTag>> {
    #[derive(serde::Deserialize, Default)]
    struct Global {
        locale: Option<String>,
    }
    #[derive(serde::Deserialize, Default)]
    struct File {
        #[serde(default)]
        global: Global,
    }
    let file: File = toml::from_str(text).map_err(|e| Error::Config(e.to_string()))?;
    file.global.locale.map(|tag| tag.parse()).transpose()
}

/// [`from_config`] on a file.
///
/// # Errors
///
/// [`Error::Open`] when the file cannot be read, and as [`from_config`].
#[cfg(feature = "config")]
pub fn from_config_file(path: &std::path::Path) -> Result<Option<LocaleTag>> {
    let text = std::fs::read_to_string(path).map_err(|source| Error::Open {
        path: path.to_path_buf(),
        source,
    })?;
    from_config(&text).map_err(|error| Error::Context {
        message: format!("{}", path.display()),
        source: Box::new(error),
    })
}

/// The locale the configuration of a command line gives: the file `config` when the command names
/// one (it must exist), else `sheetrs.toml` in the working directory when there is one; `None` when
/// there is no file or it states no locale.
///
/// # Errors
///
/// As [`from_config_file`].
#[cfg(feature = "config")]
pub fn configured(config: Option<&std::path::Path>) -> Result<Option<LocaleTag>> {
    if let Some(path) = config {
        return from_config_file(path);
    }
    let default = std::path::Path::new(CONFIG_FILE);
    if default.is_file() {
        from_config_file(default)
    } else {
        Ok(None)
    }
}

/// The locale of the book of a command line: the `--locale` flag when it was given, else
/// [`configured`].
///
/// # Errors
///
/// [`Error::InvalidLocale`] when the flag is not a locale tag, and as [`configured`].
#[cfg(feature = "config")]
pub fn for_command(
    flag: Option<&str>,
    config: Option<&std::path::Path>,
) -> Result<Option<LocaleTag>> {
    match flag {
        Some(tag) => tag.parse().map(Some),
        None => configured(config),
    }
}

/// The configuration file a command line reads from the working directory when it is not told which.
pub const CONFIG_FILE: &str = "sheetrs.toml";

#[cfg(test)]
mod tests {
    use super::*;

    fn env(pairs: &'static [(&str, &str)]) -> impl Fn(&str) -> Option<String> {
        move |name| {
            pairs
                .iter()
                .find(|(k, _)| *k == name)
                .map(|(_, v)| (*v).to_string())
        }
    }

    #[test]
    fn a_tag_is_normalized_and_a_script_is_skipped() {
        let tag = |t: &str| t.parse::<LocaleTag>().unwrap().to_string();
        assert_eq!(tag("es_co"), "es-CO");
        assert_eq!(tag("EN"), "en");
        assert_eq!(tag("sr-Latn-RS"), "sr-RS");
        assert_eq!(tag("es-419"), "es-419");
        assert_eq!(tag(" de-DE "), "de-DE");
    }

    #[test]
    fn what_is_not_a_tag_is_refused() {
        for bad in [
            "", "e", "english", "es-", "es-CO-XX", "es CO", "12", "es-C", "C.UTF-8",
        ] {
            assert!(bad.parse::<LocaleTag>().is_err(), "{bad:?}");
        }
        let error = "nope!".parse::<LocaleTag>().unwrap_err().to_string();
        assert!(error.contains("nope!"), "{error}");
    }

    #[test]
    fn the_system_locale_is_the_one_libreoffice_reads_on_linux() {
        let tag = |pairs| from_environment(env(pairs)).map(|t| t.to_string());
        assert_eq!(tag(&[("LANG", "es_CO.UTF-8")]).as_deref(), Some("es-CO"));
        assert_eq!(tag(&[("LC_CTYPE", "de_DE")]).as_deref(), Some("de-DE"));
        // LC_ALL beats the rest, an empty variable is not set, and the others do not count.
        assert_eq!(
            tag(&[("LC_ALL", "de_DE.UTF-8"), ("LANG", "es_CO")]).as_deref(),
            Some("de-DE")
        );
        assert_eq!(
            tag(&[("LC_ALL", ""), ("LANG", "es_ES@euro")]).as_deref(),
            Some("es-ES")
        );
        assert_eq!(tag(&[("LC_MESSAGES", "de_DE")]), None);
        assert_eq!(
            tag(&[("LANG", "es_CO"), ("LC_MESSAGES", "de_DE")]).as_deref(),
            Some("es-CO")
        );
        assert_eq!(tag(&[("LANGUAGE", "de")]), None);
        // The C locale asks for none, and it is the first set that counts.
        assert_eq!(tag(&[("LANG", "C")]), None);
        assert_eq!(tag(&[("LANG", "POSIX")]), None);
        assert_eq!(tag(&[("LANG", "es_CO"), ("LC_ALL", "C.UTF-8")]), None);
        assert_eq!(tag(&[]), None);
    }

    #[cfg(feature = "config")]
    #[test]
    fn the_configuration_names_the_locale_in_global_and_refuses_a_bad_one() {
        let tag = |t: &str| from_config(t).map(|o| o.map(|t| t.to_string()));
        assert_eq!(
            tag("[global]\nlocale = \"es_CO\"\n").unwrap().as_deref(),
            Some("es-CO")
        );
        // Another key of the linter's file is not this reader's business.
        assert_eq!(
            tag("[global]\nmax_sheets = 3\n[sheets.A]\ndisabled_rules = []\n").unwrap(),
            None
        );
        assert_eq!(tag("").unwrap(), None);
        assert!(matches!(
            tag("[global]\nlocale = 3\n"),
            Err(Error::Config(_))
        ));
        assert!(matches!(
            tag("[global]\nlocale = \"es CO\"\n"),
            Err(Error::InvalidLocale(_))
        ));
        assert!(matches!(tag("[global"), Err(Error::Config(_))));
    }
}
