//! What a cell says as text.
//!
//! Two texts: the one a person **sees** ([`TextContext::display`]: the value written with its
//! format, in a locale) and the one a search **sees** ([`TextContext::searched`]), which is not
//! the same. `LibreOffice` searches the text of the input line of a cell, not its display:
//! a number is looked for as the digits it was typed with (`1234.5` for a cell that shows
//! `1,234.50`), a percentage as `25.6%`, a date as the date the input line writes (`03/15/2023`, or
//! the ISO date when the format of the cell is an ISO one), and the formula of a formula cell in the
//! notation of `LibreOffice` (`=SUM(A1,B1)`, `=$'My Sheet'.A1`, written with the separators of the
//! locale) — never in the notation of Excel. These are the observations of 26.8 for the `en-US`
//! locale that [`TextContext`] reproduces.

use crate::error::{Error, Result};
use sheetrs_engine::dates::serial_to_ymd_quiet;
use sheetrs_engine::options::{DateOrder, DateSystem, Locale};
use sheetrs_formula::calc_read::{Notation, native_to_calc};
use sheetrs_io::Book;
use sheetrs_io::reader::{Cell, CellValue, ExternalUrl, Format};
use sheetrs_protocol::SearchIn;
use std::fmt::Write as _;

/// What a text becomes when it is typed into a cell, as `LibreOffice` reads it: the replacement of a
/// Find & Replace is entered in the cell as if it were typed (observed in 26.8: replacing the
/// text of a cell by `12` makes a number, by `=1+1` a formula, by `TRUE` a boolean, by `''12` a text,
/// and by nothing an empty cell).
#[derive(Debug, Clone, PartialEq)]
pub enum Typed {
    /// An empty cell.
    Empty,
    Text(String),
    Number(f64),
    Boolean(bool),
    /// A formula in the notation of `LibreOffice`, without the `=`.
    Formula(String),
}

pub use sheetrs_io::EntryFormat;

/// What decides how a number is written: the locale and the date system of the workbook.
#[derive(Debug, Clone)]
pub struct TextContext {
    locale: Locale,
    sys: DateSystem,
    /// How the file writes its formulas.
    notation: Notation,
    /// The workbook, to tell where the files its formulas refer to are.
    source: Option<Book>,
    /// The external workbooks of the file, read when a formula needs them.
    externals: std::cell::OnceCell<Vec<ExternalUrl>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Kind {
    General,
    Percent,
    Date { iso: bool, time: bool },
    Time { total_hours: bool },
}

/// What a format code is about, without the quoted text and the bracketed parts that are not
/// elapsed time (`[Red]`, `[$EUR]`).
pub(crate) fn kind(format: Option<&str>) -> Kind {
    let Some(format) = format else {
        return Kind::General;
    };
    let mut letters = String::new();
    let mut percent = false;
    let mut total_hours = false;
    let mut chars = format.chars();
    while let Some(c) = chars.next() {
        match c {
            '"' => {
                for q in chars.by_ref() {
                    if q == '"' {
                        break;
                    }
                }
            }
            '\\' | '_' | '*' => {
                chars.next();
            }
            '[' => {
                let mut inner = String::new();
                for q in chars.by_ref() {
                    if q == ']' {
                        break;
                    }
                    inner.push(q);
                }
                let inner = inner.to_ascii_lowercase();
                if inner.chars().all(|c| c == 'h') && !inner.is_empty() {
                    total_hours = true;
                    letters.push('h');
                } else if !inner.is_empty() && inner.chars().all(|c| matches!(c, 'm' | 's')) {
                    letters.push(inner.chars().next().unwrap_or('m'));
                }
            }
            '%' => percent = true,
            c if c.is_ascii_alphabetic() => letters.push(c.to_ascii_lowercase()),
            _ => {}
        }
    }
    // `E+00` is scientific, `General` and `AM/PM` carry letters of their own.
    let letters = letters.replace("general", "").replace("am/pm", "");
    let has = |c: char| letters.contains(c);
    let date = has('y') || has('d') || (has('m') && !(has('h') || has('s')));
    let time = has('h') || has('s');
    if date {
        let order: String = letters
            .chars()
            .filter(|c| matches!(c, 'y' | 'm' | 'd'))
            .collect();
        let iso = order.starts_with('y') && format.contains('-');
        Kind::Date { iso, time }
    } else if time {
        Kind::Time { total_hours }
    } else if percent {
        Kind::Percent
    } else {
        Kind::General
    }
}

/// Whether the format is the text one (`@`), with or without the language mark a file may put
/// before it (`[$-409]@`).
#[must_use]
pub fn is_text(format: Option<&str>) -> bool {
    format.is_some_and(|f| {
        let mut rest = f;
        while let Some(after) = rest.strip_prefix("[$-")
            && let Some((_, tail)) = after.split_once(']')
        {
            rest = tail;
        }
        rest == "@"
    })
}

/// Whether the format is `General`, with or without the language mark a file may put before it
/// (`[$-409]General`).
#[must_use]
pub fn is_general(format: Option<&str>) -> bool {
    format.is_none_or(|f| {
        let mut rest = f;
        while let Some(after) = rest.strip_prefix("[$-")
            && let Some((_, tail)) = after.split_once(']')
        {
            rest = tail;
        }
        rest.eq_ignore_ascii_case("general")
    })
}

/// `[1]Sheet1!B3` and `'[1]my sheet'!B3:C4` as `'url'#$Sheet1.B3` and `'url'#$'my sheet'.B3:C4`.
fn xlsx_externals(text: &str, urls: &[ExternalUrl]) -> String {
    let mut out = String::new();
    let mut rest = text;
    let mut in_string = false;
    while let Some(c) = rest.chars().next() {
        if c == '"' {
            in_string = !in_string;
        }
        let quoted = c == '\'' && rest[1..].starts_with('[');
        if !in_string
            && (c == '[' || quoted)
            && let Some((url, sheet, reference, used)) = external_at(rest, quoted, urls)
        {
            let name = if sheet.chars().all(|c| c.is_alphanumeric() || c == '_') {
                sheet
            } else {
                format!("'{}'", sheet.replace('\'', "''"))
            };
            let _ = write!(out, "'{url}'#${name}.{reference}");
            rest = &rest[used..];
            continue;
        }
        out.push(c);
        rest = &rest[c.len_utf8()..];
    }
    out
}

/// The external reference that starts `text`: the URL, the sheet, the cells and the length.
fn external_at(
    text: &str,
    quoted: bool,
    urls: &[ExternalUrl],
) -> Option<(String, String, String, usize)> {
    let start = usize::from(quoted);
    let after = &text[start + 1..];
    let (number, tail) = after.split_once(']')?;
    let index: usize = number.parse().ok()?;
    let url = urls.get(index.checked_sub(1)?)?.url.clone();
    let (sheet, tail) = if quoted {
        let end = tail.find('\'')?;
        (&tail[..end], &tail[end + 1..])
    } else {
        let end = tail.find('!')?;
        (&tail[..end], &tail[end..])
    };
    let tail = tail.strip_prefix('!')?;
    let length = tail
        .find(|c: char| !(c.is_ascii_alphanumeric() || c == '$' || c == ':'))
        .unwrap_or(tail.len());
    if length == 0 || sheet.is_empty() || sheet.contains(['[', ']', '!']) {
        return None;
    }
    let reference = &tail[..length];
    let used = text.len() - tail.len() + length;
    Some((url, sheet.to_string(), reference.to_string(), used))
}

/// `'file:///the/author/test.xlsx'#$Sheet1.B3` with the URL `LibreOffice` resolves.
fn ods_externals(text: &str, urls: &[ExternalUrl]) -> String {
    let mut out = String::new();
    let mut rest = text;
    while let Some(at) = rest.find("'file:") {
        out.push_str(&rest[..at]);
        let tail = &rest[at + 1..];
        if let Some((address, _)) = tail.split_once("'#") {
            let name = address.rsplit('/').next().unwrap_or(address);
            let url = urls
                .iter()
                .find(|w| w.basename == name)
                .map_or(address, |w| w.url.as_str());
            out.push('\'');
            out.push_str(url);
            rest = &tail[address.len()..];
        } else {
            out.push('\'');
            rest = tail;
        }
    }
    out.push_str(rest);
    out
}

/// A typed formula with its external references replaced by names (see
/// [`TextContext::protect_externals`]).
#[derive(Debug, Clone, Default)]
pub struct Protected {
    /// The formula with a name where each external reference was.
    pub text: String,
    /// The name and what the file writes in its place.
    names: Vec<(String, String)>,
}

impl Protected {
    /// Whether `name` is one of the names that stand for an external reference.
    #[must_use]
    pub fn is_name(&self, name: &str) -> bool {
        self.names.iter().any(|(n, _)| n.eq_ignore_ascii_case(name))
    }

    /// The compiled formula with the references back.
    #[must_use]
    pub fn restore(&self, formula: &str) -> String {
        self.names
            .iter()
            .fold(formula.to_string(), |text, (name, native)| {
                text.replace(name.as_str(), native)
            })
    }
}

/// `1.5E3`, `-1e-3`, `.5E2`: digits with an exponent.
fn is_scientific(t: &str) -> bool {
    let t = t.trim_start_matches(['+', '-']);
    let Some(at) = t.find(['e', 'E']) else {
        return false;
    };
    let (mantissa, exponent) = (&t[..at], t[at + 1..].trim_start_matches(['+', '-']));
    !mantissa.is_empty()
        && mantissa.chars().all(|c| c.is_ascii_digit() || c == '.')
        && mantissa.chars().any(|c| c.is_ascii_digit())
        && !exponent.is_empty()
        && exponent.chars().all(|c| c.is_ascii_digit())
}

/// A time of day of 24 hours or more (`25:00`, `48:30:15`), which the input line reads as a
/// duration; the engine reads times of a day only.
fn total_hours(t: &str) -> Option<f64> {
    let parts: Vec<&str> = t.split(':').collect();
    if !(2..=3).contains(&parts.len())
        || !parts
            .iter()
            .all(|p| !p.is_empty() && p.chars().all(|c| c.is_ascii_digit()))
    {
        return None;
    }
    let [h, m] = [parts[0], parts[1]].map(|p| p.parse::<f64>().unwrap_or(f64::NAN));
    let sec = parts
        .get(2)
        .map_or(0.0, |p| p.parse::<f64>().unwrap_or(f64::NAN));
    (h >= 24.0 && m < 60.0 && sec < 60.0).then(|| (h * 3600.0 + m * 60.0 + sec) / 86400.0)
}

/// `2024-03-15`: three numbers apart with `-`, year first.
fn is_iso_date(t: &str) -> bool {
    let parts: Vec<&str> = t.split('-').collect();
    parts.len() == 3
        && parts
            .iter()
            .all(|p| !p.is_empty() && p.chars().all(|c| c.is_ascii_digit()))
}

/// `1 1/2`, `-1 1/2`, `0 3/4`: a whole number and a fraction, which the input line reads as a
/// number (a fraction alone is a date).
fn mixed_fraction(t: &str) -> Option<f64> {
    let (whole, fraction) = t.split_once(' ')?;
    let (numerator, denominator) = fraction.split_once('/')?;
    let digits = |p: &str| !p.is_empty() && p.chars().all(|c| c.is_ascii_digit());
    let sign = if whole.starts_with('-') { -1.0 } else { 1.0 };
    let whole = whole.trim_start_matches(['+', '-']);
    if !(digits(whole) && digits(numerator) && digits(denominator)) {
        return None;
    }
    let denominator: f64 = denominator.parse().ok()?;
    if denominator == 0.0 {
        return None;
    }
    Some(sign * (whole.parse::<f64>().ok()? + numerator.parse::<f64>().ok()? / denominator))
}

impl TextContext {
    /// # Errors
    ///
    /// When `locale` is not one the engine has.
    pub fn new(locale: Option<&str>, date_system: DateSystem) -> Result<Self> {
        let locale = match locale {
            None => Locale::en_us(),
            Some(tag) => Locale::by_tag(tag)
                .ok_or_else(|| Error::Request(format!("unknown locale `{tag}`")))?,
        };
        Ok(Self {
            locale,
            sys: date_system,
            notation: Notation::Excel,
            source: None,
            externals: std::cell::OnceCell::new(),
        })
    }

    /// The date system of the workbook, as the file states it ([`Book::date_system`]), instead of the
    /// one the context was made with.
    #[must_use]
    pub fn with_date_system(mut self, date_system: DateSystem) -> Self {
        self.sys = date_system;
        self
    }

    /// The date system the dates of the workbook are numbered in.
    #[must_use]
    pub fn date_system(&self) -> DateSystem {
        self.sys
    }

    /// The workbook the text is read from: its formulas show the external workbooks they refer to
    /// with the URL `LibreOffice` shows (`'file:///folder/test.xlsx'#$Sheet1.B3`), resolved against
    /// the folder of the document.
    #[must_use]
    pub fn for_book(mut self, source: &Book) -> Self {
        self.source = Some(source.clone());
        self
    }

    /// The formulas the cells carry are written as the files of this format write them (an XLSX
    /// in Excel's notation, an ODS in `OpenFormula`): what a search sees of a formula depends on it.
    #[must_use]
    pub fn for_format(mut self, format: Format) -> Self {
        self.notation = match format {
            Format::Ods => Notation::OpenFormula,
            _ => Notation::Excel,
        };
        self
    }

    /// What `input` becomes when it is typed into a cell; `text_format` is whether the cell has the
    /// text format (`@`), which keeps everything as text.
    #[must_use]
    pub fn typed(&self, input: &str, text_format: bool) -> Typed {
        self.typed_in(input, text_format, None)
    }

    /// [`typed`](Self::typed) into a cell whose number `format` is known: the literal text that the
    /// format writes around the number (`"x"0`, `0" kg"`) is accepted in the input (`x9` is `9`).
    #[must_use]
    pub fn typed_in(&self, input: &str, text_format: bool, format: Option<&str>) -> Typed {
        let typed = self.typed_plain(input, text_format);
        if let (Typed::Text(_), Some(format)) = (&typed, format)
            && !text_format
            && let Some(stripped) = strip_literals_of(input, format)
        {
            let again = self.typed_plain(&stripped, false);
            if matches!(again, Typed::Number(_)) {
                return again;
            }
        }
        typed
    }

    fn typed_plain(&self, input: &str, text_format: bool) -> Typed {
        if input.is_empty() {
            return Typed::Empty;
        }
        if text_format {
            return Typed::Text(input.to_string());
        }
        if let Some(rest) = input.strip_prefix('\'') {
            return Typed::Text(rest.to_string());
        }
        if let Some(rest) = input.strip_prefix('=')
            && !rest.is_empty()
        {
            return Typed::Formula(rest.to_string());
        }
        let trimmed = input.trim();
        if trimmed.eq_ignore_ascii_case(&self.locale.true_text) {
            return Typed::Boolean(true);
        }
        if trimmed.eq_ignore_ascii_case(&self.locale.false_text) {
            return Typed::Boolean(false);
        }
        let dates = sheetrs_engine::datetext::DateCtx::new(self.sys, None);
        match sheetrs_engine::coerce::parse_number(trimmed, &self.locale, dates)
            .or_else(|| self.currency_amount(trimmed, dates))
            .or_else(|| mixed_fraction(trimmed))
            .or_else(|| total_hours(trimmed))
            .or_else(|| self.negative_in_parentheses(trimmed, dates))
        {
            Some(x) => Typed::Number(x),
            None => Typed::Text(input.to_string()),
        }
    }

    /// `(5)` and `($5)`: a number, or an amount, in parentheses is negative.
    fn negative_in_parentheses(
        &self,
        text: &str,
        dates: sheetrs_engine::datetext::DateCtx,
    ) -> Option<f64> {
        let inner = text.strip_prefix('(')?.strip_suffix(')')?.trim();
        if inner.starts_with(['-', '+', '(']) {
            return None;
        }
        let x = sheetrs_engine::coerce::parse_number(inner, &self.locale, dates)
            .filter(|_| inner.chars().all(|c| !c.is_alphabetic() && c != '/'))
            .or_else(|| self.currency_amount(inner, dates))?;
        Some(-x)
    }

    /// The number format `LibreOffice` gives a cell in `General` that is typed `input`, when it is
    /// not `General` (`None`: a text, a formula, or a plain number).
    #[must_use]
    pub fn entry_format(&self, input: &str) -> Option<EntryFormat> {
        match self.typed_plain(input, false) {
            Typed::Boolean(_) => return Some(EntryFormat::Boolean),
            Typed::Number(_) => {}
            _ => return None,
        }
        let t = input.trim();
        let dates = sheetrs_engine::datetext::DateCtx::new(self.sys, None);
        if sheetrs_engine::coerce::parse_number(t, &self.locale, dates).is_none() {
            let inner = t
                .strip_prefix('(')
                .and_then(|r| r.strip_suffix(')'))
                .unwrap_or(t);
            return if self.currency_amount(inner, dates).is_some() {
                Some(EntryFormat::Currency)
            } else if mixed_fraction(t).is_some() {
                Some(EntryFormat::Fraction)
            } else if total_hours(t).is_some() {
                Some(EntryFormat::TimeTotal)
            } else {
                None
            };
        }
        if t.ends_with('%') {
            return Some(EntryFormat::Percent);
        }
        if is_scientific(t) {
            return Some(EntryFormat::Scientific);
        }
        if let Some(time) = self.time_format(t) {
            return Some(time);
        }
        // A date, with a time of day after it or not: the rightmost split whose tail is a time.
        for (i, ch) in t.char_indices().rev() {
            if (ch == ' ' || ch == 'T')
                && i > 0
                && self.time_format(t[i + ch.len_utf8()..].trim()).is_some()
            {
                return Some(match (is_iso_date(t[..i].trim()), ch) {
                    (true, 'T') => EntryFormat::DateTimeIsoT,
                    (true, _) => EntryFormat::DateTimeIso,
                    _ => EntryFormat::DateTime,
                });
            }
        }
        if is_iso_date(t) {
            Some(EntryFormat::DateIso)
        } else if t.contains('/') || t.contains(char::is_alphabetic) || t[1..].contains('-') {
            Some(EntryFormat::Date)
        } else {
            None
        }
    }

    /// The format of a time of day typed alone (`12:30`, `1:30 PM`, `25:00`, `12:30:45.5`).
    fn time_format(&self, t: &str) -> Option<EntryFormat> {
        let lower = t.replace('\u{a0}', " ").to_lowercase();
        let (body, half) = [&self.locale.pm, &self.locale.am]
            .iter()
            .find_map(|m| {
                lower
                    .strip_suffix(m.to_lowercase().as_str())
                    .map(|rest| (rest.trim().to_string(), true))
            })
            .unwrap_or((lower.clone(), false));
        let parts: Vec<&str> = body.split(':').collect();
        let digits = |p: &str| !p.is_empty() && p.chars().all(|c| c.is_ascii_digit());
        if !(2..=3).contains(&parts.len()) || !parts[..2].iter().all(|p| digits(p)) {
            return None;
        }
        let seconds = parts.get(2);
        let fraction = seconds.is_some_and(|s| s.contains('.'));
        if let Some(s) = seconds
            && !s.split('.').all(digits)
        {
            return None;
        }
        if fraction {
            Some(EntryFormat::TimeTotalFraction)
        } else if !half && parts[0].parse::<u32>().ok()? > 23 {
            Some(EntryFormat::TimeTotal)
        } else {
            Some(EntryFormat::Time)
        }
    }

    /// A number written with the currency symbol of the locale at one end (`$20.00`, `-$5`,
    /// `($5)`, `5$`, `$ 5`), which the input line of `LibreOffice` takes as a number. With a `%` it
    /// stays text.
    fn currency_amount(&self, text: &str, dates: sheetrs_engine::datetext::DateCtx) -> Option<f64> {
        let symbol = self.locale.currency.as_str();
        if symbol.is_empty() || !text.contains(symbol) || text.contains('%') {
            return None;
        }
        let at = text.find(symbol)?;
        let before = text[..at].trim_start_matches(['-', '(', ' ']);
        let after = text[at + symbol.len()..].trim_start_matches(' ');
        // The symbol is at the start of the number or at its end, nothing else in front of it.
        if !before.is_empty() && !after.is_empty() {
            return None;
        }
        let number = text.replacen(symbol, "", 1).replace(' ', "");
        // An exponent after a currency symbol is not a number: `$1e3` stays text.
        if number.contains(['e', 'E']) {
            return None;
        }
        sheetrs_engine::coerce::parse_number(&number, &self.locale, dates)
    }

    /// The number as `General` writes it.
    fn general(&self, x: f64) -> String {
        sheetrs_engine::coerce::fmt_number_display(x, &self.locale)
    }

    fn boolean(&self, b: bool) -> String {
        if b {
            self.locale.true_text.clone()
        } else {
            self.locale.false_text.clone()
        }
    }

    /// The cell as a person sees it: the value written with the format of the cell.
    #[must_use]
    pub fn display(&self, cell: &Cell) -> String {
        let format = cell.num_fmt.as_deref();
        match &cell.value {
            CellValue::Empty => String::new(),
            CellValue::Text(t) => format
                .and_then(|f| sheetrs_engine::format::format_text(t, f, &self.locale))
                .unwrap_or_else(|| t.to_string()),
            CellValue::Boolean(b) => self.boolean(*b),
            CellValue::Error(e) => e.to_string(),
            CellValue::Number(x) => match format {
                Some(f) if !is_general(Some(f)) => {
                    sheetrs_engine::format::format_number_shown(*x, f, &self.locale, self.sys)
                        .unwrap_or_else(|| self.general(*x))
                }
                _ => self.general(*x),
            },
        }
    }

    /// The standard date format of the locale (`MM/DD/YY` for `en-US`), which `LibreOffice` gives
    /// to the result of a date calculation.
    #[must_use]
    pub fn default_date_format(&self) -> String {
        let sep = self.locale.date_sep;
        match self.locale.date_order {
            DateOrder::Mdy => format!("MM{sep}DD{sep}YY"),
            DateOrder::Dmy => format!("DD{sep}MM{sep}YY"),
            DateOrder::Ymd => format!("YY{sep}MM{sep}DD"),
        }
    }

    /// A number as the input line of the cell writes it.
    #[must_use]
    pub fn input_line(&self, x: f64, format: Option<&str>) -> String {
        match kind(format) {
            Kind::General => self.general(x),
            Kind::Percent => format!("{}%", self.general(x * 100.0)),
            Kind::Date { iso, time } => {
                let Some((y, m, d)) = serial_to_ymd_quiet(x, self.sys) else {
                    return self.general(x);
                };
                let sep = self.locale.date_sep;
                let date = if iso {
                    format!("{y:04}-{m:02}-{d:02}")
                } else {
                    match self.locale.date_order {
                        DateOrder::Mdy => format!("{m:02}{sep}{d:02}{sep}{y:04}"),
                        DateOrder::Dmy => format!("{d:02}{sep}{m:02}{sep}{y:04}"),
                        DateOrder::Ymd => format!("{y:04}{sep}{m:02}{sep}{d:02}"),
                    }
                };
                if time {
                    format!("{date} {}", clock(x, false))
                } else {
                    date
                }
            }
            // A time of day is searched in twelve hours, `06:00:00 PM` (observed in 26.8 for
            // `hh:mm`, `hh:mm:ss`, `h:mm AM/PM`, `mm:ss`); elapsed time (`[hh]:mm:ss`) in the clock.
            Kind::Time { total_hours: true } => clock(x, true),
            Kind::Time { total_hours: false } => {
                let seconds = (x.rem_euclid(1.0) * 86_400.0).round() as i64;
                let hour = seconds / 3600;
                let (am_pm, twelve) = if hour >= 12 {
                    (&self.locale.pm, if hour == 12 { 12 } else { hour - 12 })
                } else {
                    (&self.locale.am, if hour == 0 { 12 } else { hour })
                };
                format!(
                    "{twelve:02}:{:02}:{:02} {am_pm}",
                    seconds / 60 % 60,
                    seconds % 60
                )
            }
        }
    }

    /// The formula (as the file writes it, see [`for_format`](Self::for_format)) in the notation of
    /// `LibreOffice`, with the `=`.
    #[must_use]
    pub fn formula_text(&self, native: &str, _sheet: &str) -> String {
        let text = native_to_calc(
            native,
            self.notation,
            &sheetrs_engine::functions::in_catalogue,
        );
        format!("={}", self.with_external_urls(text))
    }

    /// A formula typed in the notation of `LibreOffice` with the external references set apart
    /// (`'file:///folder/test.xlsx'#$Sheet1.B3`): the compiler of what is typed knows nothing of
    /// them, so each is a name while it works and [`Protected::restore`] writes it back as the file
    /// writes it (`[1]Sheet1!B3`, `['file:///author/test.xlsx'#$Sheet1.B3]`). One that names a
    /// workbook the file does not have stays as it is.
    #[must_use]
    pub fn protect_externals(&self, calc: &str) -> Protected {
        let mut protected = Protected::default();
        let Some(source) = self.source.as_ref().filter(|_| calc.contains("'#")) else {
            protected.text = calc.to_string();
            return protected;
        };
        let urls = self
            .externals
            .get_or_init(|| source.external_urls().unwrap_or_default());
        let xlsx = matches!(self.notation, Notation::Excel);
        let mut rest = calc;
        while let Some(at) = rest.find('\'') {
            protected.text.push_str(&rest[..at]);
            rest = &rest[at..];
            let found = rest[1..].split_once("'#").and_then(|(url, tail)| {
                let index = urls.iter().position(|w| w.url == url)?;
                let tail = tail.strip_prefix('$').unwrap_or(tail);
                let (sheet, tail) = if let Some(quoted) = tail.strip_prefix('\'') {
                    let end = quoted.find('\'')?;
                    (quoted[..end].to_string(), &quoted[end + 1..])
                } else {
                    let end = tail.find('.')?;
                    (tail[..end].to_string(), &tail[end..])
                };
                let tail = tail.strip_prefix('.')?;
                let length = tail
                    .find(|c: char| !(c.is_ascii_alphanumeric() || c == '$' || c == ':'))
                    .unwrap_or(tail.len());
                (length > 0).then(|| {
                    let reference = &tail[..length];
                    let simple = sheet.chars().all(|c| c.is_alphanumeric() || c == '_');
                    let native = if xlsx && simple {
                        format!("[{}]{sheet}!{reference}", index + 1)
                    } else if xlsx {
                        format!("'[{}]{sheet}'!{reference}", index + 1)
                    } else if simple {
                        format!("['{}'#${sheet}.{reference}]", urls[index].address)
                    } else {
                        format!("['{}'#$'{sheet}'.{reference}]", urls[index].address)
                    };
                    (native, rest.len() - tail.len() + length)
                })
            });
            if let Some((native, used)) = found {
                let token = format!("SHEETRSEXT{}X", protected.names.len());
                protected.text.push_str(&token);
                protected.names.push((token, native));
                rest = &rest[used..];
            } else {
                protected.text.push('\'');
                rest = &rest[1..];
            }
        }
        protected.text.push_str(rest);
        protected
    }

    /// The external workbooks of a formula written as `LibreOffice` writes them: `[1]Sheet1!B3` of
    /// an XLSX and the address of an ODS become `'file:///folder/test.xlsx'#$Sheet1.B3`.
    fn with_external_urls(&self, text: String) -> String {
        let xlsx = matches!(self.notation, Notation::Excel);
        let wanted = if xlsx {
            text.contains('[') && text.chars().any(|c| c.is_ascii_digit())
        } else {
            text.contains("'file:")
        };
        let Some(source) = self.source.as_ref().filter(|_| wanted) else {
            return text;
        };
        let urls = self
            .externals
            .get_or_init(|| source.external_urls().unwrap_or_default());
        if urls.is_empty() {
            return text;
        }
        if xlsx {
            xlsx_externals(&text, urls)
        } else {
            ods_externals(&text, urls)
        }
    }

    /// The text of the cell that a search of `search_in` looks at, or `None` when the search does
    /// not look at the cell: an empty one, an error, and (unless `include_formulas`) the result of
    /// a formula.
    #[must_use]
    pub fn searched(
        &self,
        cell: &Cell,
        sheet: &str,
        search_in: SearchIn,
        include_formulas: bool,
    ) -> Option<String> {
        self.searched_as(cell, sheet, search_in, include_formulas, false)
    }

    /// [`searched`](Self::searched), and with `formatted` the text a cell shows (see
    /// `Find::formatted`) in place of its input line.
    #[must_use]
    pub fn searched_as(
        &self,
        cell: &Cell,
        sheet: &str,
        search_in: SearchIn,
        include_formulas: bool,
        formatted: bool,
    ) -> Option<String> {
        if let Some(formula) = cell.formula.as_deref() {
            return match search_in {
                SearchIn::Formula => {
                    let text = self.formula_text(formula, sheet);
                    // Every cell of an array formula shows it between braces.
                    Some(if cell.is_array {
                        format!("{{{text}}}")
                    } else {
                        text
                    })
                }
                SearchIn::Value if include_formulas => self.constant_text(cell, formatted),
                SearchIn::Value | SearchIn::Comment => None,
            };
        }
        // The notes are not in the cell: the readers give them apart.
        if search_in == SearchIn::Comment {
            return None;
        }
        self.constant_text(cell, formatted)
    }

    fn constant_text(&self, cell: &Cell, formatted: bool) -> Option<String> {
        if formatted && !matches!(cell.value, CellValue::Empty | CellValue::Error(_)) {
            return Some(self.display(cell));
        }
        match &cell.value {
            CellValue::Empty | CellValue::Error(_) => None,
            CellValue::Text(t) => Some(t.to_string()),
            CellValue::Boolean(b) => Some(self.boolean(*b)),
            CellValue::Number(x) => Some(self.input_line(*x, cell.num_fmt.as_deref())),
        }
    }
}

/// The input without the literal text that `format` writes before and after the number, when the
/// input has it (`x9` for `"x"0`); `None` when the format has none or the input does not.
fn strip_literals_of(input: &str, format: &str) -> Option<String> {
    let section = format.split(';').next().unwrap_or(format);
    let mut prefix = String::new();
    let mut suffix = String::new();
    let mut seen_digit = false;
    let mut chars = section.chars().peekable();
    while let Some(c) = chars.next() {
        let literal = match c {
            '"' => Some(chars.by_ref().take_while(|&q| q != '"').collect::<String>()),
            '\\' => chars.next().map(String::from),
            '[' => {
                chars.by_ref().take_while(|&q| q != ']').for_each(drop);
                None
            }
            '_' | '*' => {
                chars.next();
                None
            }
            '0' | '#' | '?' | '.' | ',' | '%' => {
                if c != '%' {
                    seen_digit = true;
                }
                None
            }
            c if c.is_alphabetic() || c == ' ' => Some(c.to_string()),
            _ => None,
        };
        if let Some(l) = literal {
            if seen_digit {
                suffix.push_str(&l);
            } else {
                prefix.push_str(&l);
            }
        }
    }
    let mut text = input.trim();
    let mut changed = false;
    let (prefix, suffix) = (prefix.trim(), suffix.trim());
    if !prefix.is_empty()
        && let Some(rest) = text.strip_prefix(prefix)
    {
        text = rest;
        changed = true;
    }
    if !suffix.is_empty()
        && let Some(rest) = text.strip_suffix(suffix)
    {
        text = rest;
        changed = true;
    }
    changed.then(|| text.trim().to_string())
}

/// `HH:MM:SS` of the fraction of a day (the hours of the whole value when it is elapsed time).
fn clock(x: f64, total_hours: bool) -> String {
    let seconds = (x.rem_euclid(1.0) * 86_400.0).round() as i64;
    let hours = if total_hours {
        (x * 24.0).floor() as i64
    } else {
        seconds / 3600
    };
    format!("{hours:02}:{:02}:{:02}", seconds / 60 % 60, seconds % 60)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ctx() -> TextContext {
        TextContext::new(None, DateSystem::default()).unwrap()
    }

    // Observed in `LibreOffice` 26.8 (en-US): the text a search finds a cell by.
    #[test]
    fn a_number_is_searched_as_typed_and_not_as_shown() {
        let c = ctx();
        assert_eq!(c.input_line(1234.5, Some("#,##0.00")), "1234.5");
        assert_eq!(c.input_line(1234.5, Some("$#,##0.00")), "1234.5");
        assert_eq!(c.input_line(1_234_567.891, Some("0.00E+00")), "1234567.891");
        assert_eq!(c.input_line(0.1 + 0.2, None), "0.3");
        assert_eq!(
            c.input_line(-1234.5, Some("#,##0.00;[Red]-#,##0.00")),
            "-1234.5"
        );
    }

    #[test]
    fn a_percentage_is_searched_with_its_sign() {
        assert_eq!(ctx().input_line(0.256, Some("0.0%")), "25.6%");
        assert_eq!(ctx().input_line(0.5, Some("0%")), "50%");
    }

    #[test]
    fn a_date_is_searched_as_the_input_line_writes_it() {
        let c = ctx();
        assert_eq!(c.input_line(45000.0, Some("d-mmm-yy")), "03/15/2023");
        assert_eq!(c.input_line(45000.0, Some("mmmm d, yyyy")), "03/15/2023");
        assert_eq!(c.input_line(45000.0, Some("yyyy")), "03/15/2023");
        assert_eq!(c.input_line(45000.0, Some("yyyy-mm-dd")), "2023-03-15");
        assert_eq!(
            c.input_line(45000.0, Some("yyyy-mm-dd hh:mm:ss")),
            "2023-03-15 00:00:00"
        );
        assert_eq!(
            c.input_line(45000.75, Some("dd/mm/yyyy hh:mm")),
            "03/15/2023 18:00:00"
        );
        assert_eq!(c.input_line(0.75, Some("[hh]:mm:ss")), "18:00:00");
        // A time of day is searched in twelve hours (`LibreOffice` 26.8).
        assert_eq!(c.input_line(0.75, Some("hh:mm")), "06:00:00 PM");
        assert_eq!(c.input_line(0.0, Some("h:mm AM/PM")), "12:00:00 AM");
    }

    // Observed in `LibreOffice` 26.8: the replacement is entered as if it were typed.
    #[test]
    fn a_replacement_is_typed_into_the_cell() {
        let c = ctx();
        assert_eq!(c.typed("", false), Typed::Empty);
        assert_eq!(c.typed("12", false), Typed::Number(12.0));
        assert_eq!(c.typed("  12 ", false), Typed::Number(12.0));
        assert_eq!(c.typed("50%", false), Typed::Number(0.5));
        assert_eq!(c.typed("1e3", false), Typed::Number(1000.0));
        assert_eq!(c.typed("03/15/2024", false), Typed::Number(45366.0));
        assert_eq!(c.typed("true", false), Typed::Boolean(true));
        assert_eq!(c.typed("=1+1", false), Typed::Formula("1+1".into()));
        assert_eq!(c.typed("=", false), Typed::Text("=".into()));
        assert_eq!(c.typed("'12", false), Typed::Text("12".into()));
        assert_eq!(c.typed("abc", false), Typed::Text("abc".into()));
        assert_eq!(c.typed("   b", false), Typed::Text("   b".into()));
        assert_eq!(
            c.typed("12", true),
            Typed::Text("12".into()),
            "a text format keeps text"
        );
    }

    #[test]
    fn a_formula_is_searched_in_the_notation_of_libreoffice() {
        let c = ctx();
        assert_eq!(c.formula_text("SUM(A1,A1)", "T"), "=SUM(A1,A1)");
        assert_eq!(c.formula_text("'U Sheet'!A1+1", "T"), "=$'U Sheet'.A1+1");
        assert_eq!(c.formula_text("A1+$A$1", "T"), "=A1+$A$1");
    }

    #[test]
    fn the_display_is_the_value_with_its_format() {
        let c = ctx();
        let cell = Cell {
            value: CellValue::Number(1234.5),
            num_fmt: Some("#,##0.00".into()),
            ..Cell::default()
        };
        assert_eq!(c.display(&cell), "1,234.50");
        let general = Cell {
            value: CellValue::Number(0.5),
            ..Cell::default()
        };
        assert_eq!(c.display(&general), "0.5");
    }

    #[test]
    fn a_search_skips_errors_and_the_results_of_formulas_unless_asked() {
        let c = ctx();
        let formula = Cell {
            value: CellValue::Number(2469.0),
            formula: Some("A1*2".into()),
            ..Cell::default()
        };
        assert_eq!(c.searched(&formula, "T", SearchIn::Value, false), None);
        assert_eq!(
            c.searched(&formula, "T", SearchIn::Value, true).as_deref(),
            Some("2469")
        );
        assert_eq!(
            c.searched(&formula, "T", SearchIn::Formula, false)
                .as_deref(),
            Some("=A1*2")
        );
    }
}
