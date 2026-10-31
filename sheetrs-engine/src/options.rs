//! Execution parameters of the engine. Date system, iteration, search settings and the
//! syntax of reference text are defined by the document; the locale is the user's.

use crate::dates;
use crate::locale_data::{DEFAULT_REGION, LOCALES};
use serde::{Deserialize, Serialize};

/// The null date: the date of serial 0, on the real calendar (no 1900-02-29, negative serials
/// allowed). LibreOffice stores it per document and any date is allowed; the usual ones are
/// 1899-12-30 (its default, and what an XLSX gets unless it says `date1904`), 1900-01-01 and
/// 1904-01-01.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct DateSystem {
    /// Days from 1970-01-01 to the null date.
    pub null_date: i64,
}

impl DateSystem {
    pub const NULL_1899: Self = Self { null_date: -25_569 };
    pub const NULL_1900: Self = Self { null_date: -25_567 };
    pub const NULL_1904: Self = Self { null_date: -24_107 };

    /// The system whose serial 0 is the given calendar date.
    pub fn from_ymd(y: i64, m: i64, d: i64) -> Option<Self> {
        let ok = (1..=12).contains(&m) && (1..=dates::days_in_month(y, m)).contains(&d);
        ok.then(|| Self {
            null_date: dates::days_from_civil(y, m, d),
        })
    }

    /// The null date as `(year, month, day)`.
    pub fn ymd(self) -> (i64, i64, i64) {
        dates::civil_from_days(self.null_date)
    }
}

impl Default for DateSystem {
    fn default() -> Self {
        Self::NULL_1899
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub struct Iteration {
    pub max_iterations: u32,
    pub max_change: f64,
}

impl Default for Iteration {
    fn default() -> Self {
        // Excel defaults.
        Self {
            max_iterations: 100,
            max_change: 0.001,
        }
    }
}

/// Order of the numeric components of a text date (`3/1/2024`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum DateOrder {
    Mdy,
    Dmy,
    Ymd,
}

/// How text criteria and patterns are interpreted by `COUNTIF`, `MATCH`,
/// `VLOOKUP`, `SEARCH`, … (a document setting; an XLSX import
/// is `Wildcards`, whole cell, case-insensitive).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SearchMode {
    /// `*`, `?` and `~` escapes.
    Wildcards,
    /// Regular expressions.
    Regex,
    /// The text as written.
    Literal,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct Search {
    pub mode: SearchMode,
    /// Criteria must match the whole cell, not a part of it.
    pub whole_cell: bool,
    /// Text comparison distinguishes case (comparison operators, criteria and
    /// lookups).
    pub case_sensitive: bool,
}

impl Search {
    /// What LibreOffice sets when it imports an XLSX: wildcards, whole cell, case-insensitive.
    pub const XLSX_IMPORT: Self = Self {
        mode: SearchMode::Wildcards,
        whole_cell: true,
        case_sensitive: false,
    };
}

impl Default for Search {
    fn default() -> Self {
        Self::XLSX_IMPORT
    }
}

/// A locale of the generated table (`locale_data.rs`), as LibreOffice reports it.
pub struct LocaleRow {
    pub tag: &'static str,
    /// The language id of `[$-409]` in a format code.
    pub lcid: u16,
    /// `mdy`, `dmy` or `ymd`: the order of the short date.
    pub order: &'static str,
    pub decimal: char,
    pub group: char,
    pub date_sep: char,
    pub true_text: &'static str,
    pub false_text: &'static str,
    pub am: &'static str,
    pub pm: &'static str,
    pub currency: &'static str,
    /// `DOLLAR`'s positive and negative amounts: `¤` is the symbol, `#` the number.
    pub currency_pos: &'static str,
    pub currency_neg: &'static str,
    pub months: [&'static str; 12],
    pub months_abbr: [&'static str; 12],
    pub days: [&'static str; 7],
    pub days_abbr: [&'static str; 7],
    /// What the input scanner accepts as a date of digits: `D/M/Y`, `M/D`… (any other character
    /// is a separator that must be written as it is).
    pub date_patterns: &'static [&'static str],
    /// How the scanner reads a date with a month name (see `locale_data.rs`).
    pub named: &'static str,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Locale {
    pub tag: String,
    /// The language id of `[$-409]` in a format code (0: none).
    pub lcid: u16,
    pub date_order: DateOrder,
    /// First year of the 100-year window for two-digit years in text dates and
    /// (LibreOffice) `DATE`: 1930 makes 30–99 → 19xx and 00–29 → 20xx.
    pub two_digit_year_start: u16,
    pub decimal: char,
    pub group: char,
    /// The separator of the short date.
    pub date_sep: char,
    pub true_text: String,
    pub false_text: String,
    /// Texts of `AM/PM` in a time format.
    pub am: String,
    pub pm: String,
    pub currency: String,
    /// `DOLLAR`'s positive and negative amounts: `¤` stands for the symbol, `#` for the number.
    pub currency_pos: String,
    pub currency_neg: String,
    pub months: [String; 12],
    pub months_abbr: [String; 12],
    /// Sunday first.
    pub days: [String; 7],
    pub days_abbr: [String; 7],
    /// The patterns of a date of digits the locale accepts (see [`LocaleRow`]).
    pub date_patterns: Vec<String>,
    /// How a date with a month name is read, one letter for each form (see `datetext`).
    pub named: String,
}

fn owned<const N: usize>(a: [&str; N]) -> [String; N] {
    a.map(String::from)
}

impl Locale {
    pub fn from_row(r: &LocaleRow) -> Self {
        Self {
            tag: r.tag.into(),
            lcid: r.lcid,
            date_order: match r.order {
                "mdy" => DateOrder::Mdy,
                "ymd" => DateOrder::Ymd,
                _ => DateOrder::Dmy,
            },
            two_digit_year_start: 1930,
            decimal: r.decimal,
            group: r.group,
            date_sep: r.date_sep,
            true_text: r.true_text.into(),
            false_text: r.false_text.into(),
            am: r.am.into(),
            pm: r.pm.into(),
            currency: r.currency.into(),
            currency_pos: r.currency_pos.into(),
            currency_neg: r.currency_neg.into(),
            months: owned(r.months),
            months_abbr: owned(r.months_abbr),
            days: owned(r.days),
            days_abbr: owned(r.days_abbr),
            date_patterns: r.date_patterns.iter().map(|p| (*p).into()).collect(),
            named: r.named.into(),
        }
    }

    /// The row of the table for a tag: `es-CO` itself, or a bare language (`es`) as its
    /// default region.
    pub fn row(tag: &str) -> Option<&'static LocaleRow> {
        let norm = tag.replace('_', "-");
        let (lang, region) = match norm.split_once('-') {
            Some((l, r)) => (l.to_ascii_lowercase(), Some(r.to_ascii_uppercase())),
            None => (norm.to_ascii_lowercase(), None),
        };
        let wanted = match region {
            Some(r) => format!("{lang}-{r}"),
            None => DEFAULT_REGION
                .iter()
                .find(|(l, _)| *l == lang)?
                .1
                .to_string(),
        };
        LOCALES.iter().find(|r| r.tag == wanted)
    }

    /// The row whose language id is `lcid` (`[$-407]`): the locale a format code names.
    pub fn row_by_lcid(lcid: u16) -> Option<&'static LocaleRow> {
        LOCALES.iter().find(|r| r.lcid == lcid)
    }

    pub fn en_us() -> Self {
        Self::by_tag("en-US").unwrap_or_else(|| unreachable!("en-US is in the table"))
    }

    pub fn es_es() -> Self {
        Self::by_tag("es-ES").unwrap_or_else(|| unreachable!("es-ES is in the table"))
    }

    /// Locale by tag (`en`, `es-CO`, `es_ES`): any of the table.
    pub fn by_tag(tag: &str) -> Option<Self> {
        Self::row(tag).map(Self::from_row)
    }
}

impl Default for Locale {
    fn default() -> Self {
        Self::en_us()
    }
}

/// What `INFO` and `CELL("filename")` say of the machine and the file. `None` is the machine the
/// program runs on and the executable's own directory, as a program computes them; a build (or a
/// run) can fix each of them, as it can the clock, so that the results do not change from one
/// machine to another.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct Host {
    /// `INFO("system")`: `LINUX`, `WNT` or `MACOSX`.
    pub system: Option<String>,
    /// `INFO("osversion")`.
    pub osversion: Option<String>,
    /// `INFO("release")`: the version of the program that calculates.
    pub release: Option<String>,
    /// The directory `CELL("filename")` puts before the file name, as given (a path).
    pub directory: Option<String>,
    /// The file name of the workbook, with its extension.
    pub filename: Option<String>,
}

impl Host {
    /// The fields set on `self` win over `base`.
    #[must_use]
    pub fn over(self, base: Self) -> Self {
        Self {
            system: self.system.or(base.system),
            osversion: self.osversion.or(base.osversion),
            release: self.release.or(base.release),
            directory: self.directory.or(base.directory),
            filename: self.filename.or(base.filename),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct EngineOptions {
    pub date_system: DateSystem,
    /// `None`: circular references evaluate to `#CIRC!`.
    pub iteration: Option<Iteration>,
    pub locale: Locale,
    /// Text criteria and pattern settings.
    pub search: Search,
    /// `INDIRECT` reads `Sheet!A1` (Excel A1); false: `Sheet.A1` (Calc A1, a document
    /// created in LibreOffice).
    pub string_ref_excel: bool,
    /// Fixed value (serial number) for `NOW()`/`TODAY()`; the system clock
    /// when `None`.
    pub now: Option<f64>,
    /// Seed of the `RAND` generator; derived from the clock when `None`.
    pub rand_seed: Option<u64>,
    #[serde(default)]
    pub host: Host,
}

impl Default for EngineOptions {
    fn default() -> Self {
        Self {
            date_system: DateSystem::default(),
            iteration: None,
            locale: Locale::default(),
            search: Search::XLSX_IMPORT,
            string_ref_excel: true,
            now: None,
            rand_seed: None,
            host: Host::default(),
        }
    }
}

impl EngineOptions {
    /// What reading a text as a date needs: the null date and the clock.
    pub fn date_ctx(&self) -> crate::datetext::DateCtx {
        crate::datetext::DateCtx::new(self.date_system, self.now)
    }

    /// Current date-time as a serial number in the configured date system.
    pub fn now_serial(&self) -> f64 {
        self.now.unwrap_or_else(|| {
            let secs = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map_or(0.0, |d| d.as_secs_f64());
            dates::unix_to_serial(secs, self.date_system)
        })
    }
}
