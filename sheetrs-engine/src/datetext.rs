//! Dates and times written as text, as LibreOffice's input scanner reads them under a locale.
//!
//! A date of digits is accepted when it has the shape of one of the locale's patterns
//! (`D/M/Y`, `M/D`: the separators must be written as in the pattern), or the ISO shape `Y-M-D`
//! (`a-b-c` read as year, month, day, after the patterns). A pattern without a year gives the current year. A date with a month name is read
//! by the form it is written in: the position of the month, the separator, and the digits of the
//! numbers. What LibreOffice does with each form depends on the locale and is data
//! (`Locale::named`, measured by `scripts/gen_locale_table.py`).

use crate::coerce::expand_year;
use crate::dates;
use crate::options::{DateSystem, Locale};

/// What reading a text as a date needs besides the locale: the null date and the clock (a date
/// written without a year is of the current one).
#[derive(Debug, Clone, Copy)]
pub struct DateCtx {
    pub sys: DateSystem,
    /// The fixed `now` of the build or the run; the system clock when `None`.
    pub now: Option<f64>,
}

impl DateCtx {
    pub fn new(sys: DateSystem, now: Option<f64>) -> Self {
        Self { sys, now }
    }

    /// The current year, read from the clock only when a date asks for it.
    pub fn year(&self) -> i64 {
        let serial = self.now.unwrap_or_else(|| {
            let secs = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map_or(0.0, |d| d.as_secs_f64());
            dates::unix_to_serial(secs, self.sys)
        });
        dates::serial_to_ymd_quiet(serial, self.sys).map_or(2000, |d| d.0)
    }
}

/// A text date or date-time, as typed into a cell, optionally followed by a time. Also a time
/// alone (`13:30`, `1:30 PM`), which is a fraction of a day.
pub fn parse_date_text(s: &str, loc: &Locale, dc: DateCtx) -> Option<f64> {
    let t = s.trim();
    if let Some(tm) = parse_time(t, loc) {
        return Some(tm);
    }
    // Rightmost split whose tail is a time.
    for (i, ch) in t.char_indices().rev() {
        if (ch == ' ' || ch == 'T')
            && i > 0
            && let Some(tm) = parse_time(t[i + ch.len_utf8()..].trim(), loc)
        {
            return parse_date_only(t[..i].trim(), loc, dc).map(|d| d + tm);
        }
    }
    parse_date_only(t, loc, dc)
}

/// A time of day: `13:30`, `1:30:15`, `1:30 PM` (with the locale's text for AM and PM).
fn parse_time(s: &str, loc: &Locale) -> Option<f64> {
    let t = s.trim();
    let half = |text: &str| -> Option<&str> {
        let suffix = text.replace('\u{a0}', " ").to_lowercase();
        let lower = t.replace('\u{a0}', " ").to_lowercase();
        lower
            .strip_suffix(suffix.as_str())
            .map(|rest| &t[..rest.len()])
    };
    let (body, pm) = if let Some(b) = half(&loc.pm) {
        (b.trim(), Some(true))
    } else if let Some(b) = half(&loc.am) {
        (b.trim(), Some(false))
    } else {
        (t, None)
    };
    let parts: Vec<&str> = body.split(':').collect();
    if !(2..=3).contains(&parts.len()) {
        return None;
    }
    let num = |p: &str| -> Option<f64> {
        (!p.is_empty() && p.chars().all(|c| c.is_ascii_digit() || c == '.'))
            .then(|| p.parse().ok())?
    };
    let mut h = num(parts[0])?;
    let m = num(parts[1])?;
    let sec = parts.get(2).map_or(Some(0.0), |p| num(p))?;
    if let Some(pm) = pm {
        if !(1.0..=12.0).contains(&h) {
            return None;
        }
        h = h % 12.0 + if pm { 12.0 } else { 0.0 };
    }
    if h > 23.0 && pm.is_none() || m >= 60.0 || sec >= 60.0 {
        return None;
    }
    Some((h * 3600.0 + m * 60.0 + sec) / 86400.0)
}

fn parse_date_only(d: &str, loc: &Locale, dc: DateCtx) -> Option<f64> {
    let (y, m, day) = match numeric_date(d, loc, dc) {
        Some(r) => r?,
        None => named_date(d, loc, dc)?,
    };
    if !(1..=12).contains(&m) || day < 1 || day > dates::days_in_month(y, m) {
        return None;
    }
    dates::ymd_to_serial_quiet(y, m, day, dc.sys)
}

// ----- dates of digits -----

/// The numbers of a text and what stands between them: `3/1/2024` is `[3, 1, 2024]` and
/// `["/", "/"]`; a separator after the last number is kept as one more.
fn split_digits(d: &str) -> Option<(Vec<&str>, Vec<&str>)> {
    let mut nums = Vec::new();
    let mut seps = Vec::new();
    let mut rest = d;
    if !rest.starts_with(|c: char| c.is_ascii_digit()) {
        return None;
    }
    while !rest.is_empty() {
        let end = rest
            .find(|c: char| !c.is_ascii_digit())
            .unwrap_or(rest.len());
        nums.push(&rest[..end]);
        rest = &rest[end..];
        let end = rest
            .find(|c: char| c.is_ascii_digit())
            .unwrap_or(rest.len());
        if end > 0 {
            seps.push(&rest[..end]);
        }
        rest = &rest[end..];
    }
    Some((nums, seps))
}

/// A pattern's letters (`D`, `M`, `Y`) and the separators between and after them.
fn split_pattern(p: &str) -> (Vec<char>, Vec<&str>) {
    let mut roles = Vec::new();
    let mut seps = Vec::new();
    let mut start = None;
    for (i, c) in p.char_indices() {
        if matches!(c, 'D' | 'M' | 'Y') {
            if let Some(s) = start.take() {
                seps.push(&p[s..i]);
            }
            roles.push(c);
        } else if start.is_none() {
            start = Some(i);
        }
    }
    if let Some(s) = start {
        seps.push(&p[s..]);
    }
    (roles, seps)
}

/// The shape `a-b-c` that is read as `Y-M-D` whatever the locale.
fn iso_shaped(nums: &[&str], seps: &[&str]) -> bool {
    nums.len() == 3 && seps == ["-", "-"]
}

/// Whether a text has the shape of a date of digits in the locale.
pub fn looks_like_numeric_date(s: &str, loc: &Locale) -> bool {
    let Some((nums, seps)) = split_digits(s) else {
        return false;
    };
    iso_shaped(&nums, &seps)
        || loc.date_patterns.iter().any(|p| {
            let (roles, ps) = split_pattern(p);
            roles.len() == nums.len() && ps == seps
        })
}

/// `None`: not the shape of a date of digits; `Some(None)`: it is, but it is not a date.
fn numeric_date(d: &str, loc: &Locale, dc: DateCtx) -> Option<Option<(i64, i64, i64)>> {
    let (nums, seps) = split_digits(d)?;
    let iso = iso_shaped(&nums, &seps);
    let read = |roles: &[char]| -> Option<(i64, i64, i64)> {
        let (mut y, mut m, mut day) = (None::<i64>, 0, 0);
        for (role, n) in roles.iter().zip(&nums) {
            let v: i64 = n.parse().ok()?;
            match role {
                'D' if n.len() <= 2 => day = v,
                'M' if n.len() <= 2 => m = v,
                'Y' if n.len() <= 5 => {
                    y = Some(if n.len() <= 2 {
                        expand_year(v, loc.two_digit_year_start)
                    } else {
                        v
                    });
                }
                _ => return None,
            }
        }
        Some((y.unwrap_or_else(|| dc.year()), m, day))
    };
    let valid = |d: &(i64, i64, i64)| {
        (1..=12).contains(&d.1) && (1..=dates::days_in_month(d.0, d.1)).contains(&d.2)
    };
    let mut matched = false;
    for p in &loc.date_patterns {
        let (roles, ps) = split_pattern(p);
        if roles.len() == nums.len() && ps == seps {
            matched = true;
            if let Some(date) = read(&roles).filter(valid) {
                return Some(Some(date));
            }
        }
    }
    if iso {
        return Some(read(&['Y', 'M', 'D']));
    }
    matched.then_some(None)
}

// ----- dates with a month name -----

/// The month a word names (full or abbreviated, with or without the dot of an abbreviation) and
/// whether the dot was added to an abbreviation that has none.
fn month_of(word: &str, loc: &Locale) -> Option<(i64, bool)> {
    let w = word.to_lowercase();
    for (i, month) in (1..=12_i64).enumerate() {
        let full = loc.months[i].to_lowercase();
        let abbr = loc.months_abbr[i].to_lowercase();
        // The dot belongs to an abbreviation that is shorter than the name (`mar.`, not `mars.`).
        let added_dot =
            abbr != full && !abbr.ends_with('.') && w.strip_suffix('.') == Some(abbr.as_str());
        let by_abbr = w == abbr || added_dot;
        // `Sept` is read as September wherever the abbreviation is `Sep`.
        let sept = i == 8 && abbr == "sep" && w == "sept";
        if w == full || by_abbr || sept {
            return Some((month, added_dot));
        }
    }
    None
}

fn is_weekday(word: &str, loc: &Locale) -> bool {
    let w = word.to_lowercase();
    (0..7).any(|i| {
        let abbr = loc.days_abbr[i].to_lowercase();
        w == loc.days[i].to_lowercase() || w == abbr || w.strip_suffix('.') == Some(abbr.as_str())
    })
}

#[derive(Debug, PartialEq)]
enum Item<'a> {
    Month(i64),
    Num(&'a str),
}

/// The form a date with a month name is written in: the items, the separator between them (an
/// index of `[' ', '-', '/']`), a dot after the day, a comma after the second item.
struct Written<'a> {
    items: Vec<Item<'a>>,
    sep: usize,
    day_dot: bool,
    comma: bool,
}

/// `d` without the names of weekdays that contain a hyphen or a space (`segunda-feira`), which the
/// scan would take for separators.
fn without_long_weekdays(d: &str, loc: &Locale) -> String {
    let mut out = d.to_string();
    let mut names: Vec<&String> = loc.days.iter().filter(|n| n.contains(['-', ' '])).collect();
    names.sort_by_key(|n| std::cmp::Reverse(n.len()));
    for n in names {
        let lower = out.to_lowercase();
        if let Some(at) = lower.find(&n.to_lowercase())
            && lower.is_char_boundary(at)
            && out.is_char_boundary(at)
            && out.is_char_boundary(at + n.len())
        {
            out.replace_range(at..at + n.len(), "");
            out = out.trim().trim_matches(',').trim().to_string();
        }
    }
    out
}

fn scan_named<'a>(d: &'a str, loc: &Locale) -> Option<Written<'a>> {
    let mut items = Vec::new();
    let mut seps: Vec<char> = Vec::new();
    let (mut day_dot, mut comma) = (false, false);
    // The separator before the next item, once an item came before it.
    let mut pending: Option<char> = None;
    let mut dotted = false;
    let mut i = 0;
    while i < d.len() {
        let c = d[i..].chars().next()?;
        if c.is_ascii_digit() {
            let end = d[i..]
                .find(|c: char| !c.is_ascii_digit())
                .map_or(d.len(), |p| i + p);
            if let Some(s) = pending.take() {
                seps.push(s);
            }
            items.push(Item::Num(&d[i..end]));
            i = end;
            // A dot after the day, before a space: `1. Mar`.
            if items.len() == 1 && d[i..].starts_with(". ") {
                day_dot = true;
                i += 1;
            }
        } else if c.is_alphabetic() {
            let mut end = d[i..]
                .find(|c: char| !c.is_alphabetic())
                .map_or(d.len(), |p| i + p);
            // An abbreviation's dot belongs to it when a separator or the end follows.
            if d[end..].starts_with('.') && !d[end + 1..].starts_with(|c: char| c.is_alphanumeric())
            {
                end += 1;
            }
            let word = &d[i..end];
            i = end;
            if let Some((m, added_dot)) = month_of(word, loc) {
                dotted |= added_dot;
                if let Some(s) = pending.take() {
                    seps.push(s);
                }
                items.push(Item::Month(m));
            } else if is_weekday(word, loc) {
                // Weekday names around a date are ignored, with the separator after them.
                pending = None;
                let rest = d[i..].trim_start_matches([',', ' ']);
                i = d.len() - rest.len();
            } else {
                return None;
            }
        } else {
            // A run of separators: spaces, one comma after the second item, one `-` or `/`.
            let end = d[i..]
                .find(|c: char| !matches!(c, ' ' | ',' | '-' | '/'))
                .map_or(d.len(), |p| i + p);
            if end == i {
                // Not a separator of a date (`:`, `.`).
                return None;
            }
            let run = &d[i..end];
            i = end;
            let sep = match run {
                _ if run.chars().all(|c| c == ' ') => ' ',
                "," | ", " if items.len() == 2 && !comma => {
                    comma = true;
                    ' '
                }
                "-" => '-',
                "/" => '/',
                _ => return None,
            };
            if !items.is_empty() {
                pending = Some(sep);
            }
        }
    }
    let first = *seps.first()?;
    if seps.iter().any(|&s| s != first) || !(2..=3).contains(&items.len()) {
        return None;
    }
    let sep = [' ', '-', '/'].iter().position(|&x| x == first)?;
    // With hyphens, a dotted abbreviation does not stand between two numbers (`13-mar.-2024`).
    let middle = matches!(
        items.as_slice(),
        [Item::Num(_), Item::Month(..), Item::Num(_)]
    );
    if dotted && sep == 1 && middle {
        return None;
    }
    Some(Written {
        items,
        sep,
        day_dot,
        comma,
    })
}

/// The numbers and the letter the locale gives the form (`d`: the first number is the day; `y`: the
/// year, or the year of the two-digit form's first number).
fn named_date(d: &str, loc: &Locale, dc: DateCtx) -> Option<(i64, i64, i64)> {
    let stripped = without_long_weekdays(d, loc);
    let w = scan_named(&stripped, loc)?;
    let big = |n: &str| n.len() >= 3;
    let num = |n: &str| -> Option<i64> { (n.len() <= 4).then(|| n.parse().ok())? };
    let year = |n: &str| -> Option<i64> {
        let v = num(n)?;
        Some(if n.len() <= 2 {
            expand_year(v, loc.two_digit_year_start)
        } else {
            v
        })
    };
    use Item::{Month, Num};
    // Index among the forms of `Locale::named`, and the roles of the numbers.
    let form = match w.items.as_slice() {
        [Month(_), Num(a)] => {
            if big(a) {
                5
            } else {
                3
            }
        }
        [Num(a), Month(_)] => {
            if big(a) {
                6
            } else {
                4
            }
        }
        [Month(_), Num(a), Num(b)] if !big(a) => {
            if big(b) {
                0
            } else {
                7
            }
        }
        [Num(a), Month(_), Num(b)] => match (big(a), big(b)) {
            (false, true) => 1,
            (false, false) => 8,
            (true, false) => 2,
            (true, true) => return None,
        },
        _ => return None,
    };
    let m = month_item(&w)?;
    let index = if w.day_dot || w.comma {
        // A dot after the day or a comma after the second item: forms of their own.
        match (form, w.day_dot, w.comma, w.sep) {
            (1, true, false, 0) => 27,
            (4, true, false, 0) => 28,
            (0, false, true, 0) => 29,
            _ => return None,
        }
    } else {
        w.sep * 9 + form
    };
    let letter = *loc.named.as_bytes().get(index)?;
    if letter != b'd' && letter != b'y' {
        return None;
    }
    let (y, day) = match (w.items.as_slice(), form) {
        // A year of four digits: the other number is the day.
        ([Item::Month(..), Num(a), Num(b)], 0) | ([Num(a), Item::Month(..), Num(b)], 1) => {
            (year(b)?, num(a)?)
        }
        ([Num(a), Item::Month(..), Num(b)], 2) => (year(a)?, num(b)?),
        // Two digits each: the locale says which is the year.
        ([Item::Month(..), Num(a), Num(b)], 7) | ([Num(a), Item::Month(..), Num(b)], 8) => {
            if letter == b'd' {
                (year(b)?, num(a)?)
            } else {
                (year(a)?, num(b)?)
            }
        }
        (_, 5 | 6) => {
            let Num(n) = w.items.iter().find(|i| matches!(i, Num(_)))? else {
                return None;
            };
            (year(n)?, 1)
        }
        (_, 3 | 4) => {
            let Num(n) = w.items.iter().find(|i| matches!(i, Num(_)))? else {
                return None;
            };
            if letter == b'd' {
                (dc.year(), num(n)?)
            } else {
                (year(n)?, 1)
            }
        }
        _ => return None,
    };
    Some((y, m, day))
}

fn month_item(w: &Written) -> Option<i64> {
    w.items.iter().find_map(|i| match i {
        Item::Month(m) => Some(*m),
        Item::Num(_) => None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Serials of the dates read below, and the clock at 2025-01-01 (45658).
    fn read(tag: &str, text: &str) -> Option<String> {
        let loc = Locale::by_tag(tag).unwrap();
        let dc = DateCtx::new(DateSystem::NULL_1899, Some(45658.0));
        let n = parse_date_text(text, &loc, dc)?;
        let (y, m, d) = dates::serial_to_ymd_quiet(n, DateSystem::NULL_1899)?;
        Some(format!("{y}-{m:02}-{d:02}"))
    }

    #[test]
    fn digits_follow_the_patterns_of_the_locale() {
        assert_eq!(read("en-US", "3/1/2024").as_deref(), Some("2024-03-01"));
        assert_eq!(read("es-ES", "3/1/2024").as_deref(), Some("2024-01-03"));
        assert_eq!(read("en-GB", "3/1/2024").as_deref(), Some("2024-01-03"));
        // The separator is the pattern's: `.` is not a date in English or Spanish.
        assert_eq!(read("en-US", "3.1.2024"), None);
        assert_eq!(read("es-ES", "3.1.2024"), None);
        // British lists `D-M` as well, but not with a year.
        assert_eq!(read("en-GB", "3-1").as_deref(), Some("2025-01-03"));
        assert_eq!(read("en-US", "3-1"), None);
        // A day that does not exist is no date, and no other order is tried.
        assert_eq!(read("en-US", "13/1/2024"), None);
        assert_eq!(read("en-US", "2/30/2024"), None);
        assert_eq!(read("en-US", "3/1/99").as_deref(), Some("1999-03-01"));
        assert_eq!(read("en-US", "3/1/29").as_deref(), Some("2029-03-01"));
    }

    #[test]
    fn years_go_up_to_32767() {
        assert_eq!(read("en-US", "3/1/12345").as_deref(), Some("12345-03-01"));
        assert_eq!(read("en-US", "3/1/32767").as_deref(), Some("32767-03-01"));
        assert_eq!(read("en-US", "3/1/32768"), None);
        assert_eq!(read("en-US", "12345-3-1").as_deref(), Some("12345-03-01"));
        assert_eq!(read("en-US", "3/1/005").as_deref(), Some("5-03-01"));
    }

    #[test]
    fn a_pattern_without_a_year_takes_the_year_of_the_clock() {
        assert_eq!(read("en-US", "3/1").as_deref(), Some("2025-03-01"));
        assert_eq!(read("es-ES", "1 mar").as_deref(), Some("2025-03-01"));
        assert_eq!(read("en-US", "Mar 1").as_deref(), Some("2025-03-01"));
    }

    #[test]
    fn hyphenated_dates_are_year_month_day() {
        assert_eq!(read("en-US", "2024-3-1").as_deref(), Some("2024-03-01"));
        assert_eq!(read("en-US", "99-12-31").as_deref(), Some("1999-12-31"));
        assert_eq!(read("es-ES", "3-1-24").as_deref(), Some("2003-01-24"));
        assert_eq!(read("es-ES", "31-12-99"), None);
    }

    #[test]
    fn a_month_name_is_read_by_the_form_the_locale_gives_it() {
        // English: month first only.
        assert_eq!(read("en-US", "Mar 13 2024").as_deref(), Some("2024-03-13"));
        assert_eq!(read("en-US", "13 Mar 2024"), None);
        assert_eq!(read("en-US", "Mar 13, 2024").as_deref(), Some("2024-03-13"));
        // Spanish and British: either order; a lone number after the month is a year.
        assert_eq!(read("es-ES", "13 mar 2024").as_deref(), Some("2024-03-13"));
        assert_eq!(read("en-GB", "13 Mar 2024").as_deref(), Some("2024-03-13"));
        assert_eq!(read("es-ES", "mar 13").as_deref(), Some("2013-03-01"));
        assert_eq!(read("es-ES", "13. marzo 2024"), None);
        // The words are the full name, the abbreviation (with a dot added) and `Sept`; a prefix
        // is not enough, and the language is the locale's.
        assert_eq!(read("en-US", "Sept 1 2024").as_deref(), Some("2024-09-01"));
        assert_eq!(read("en-US", "Marc 1 2024"), None);
        assert_eq!(read("es-ES", "13 mar. 2024").as_deref(), Some("2024-03-13"));
        assert_eq!(read("en-GB", "13-Mar.-2024"), None);
        assert_eq!(read("es-ES", "March 1 2024"), None);
        // Weekday names around the date are ignored.
        assert_eq!(
            read("en-US", "Monday Mar 13 2024").as_deref(),
            Some("2024-03-13")
        );
        assert_eq!(
            read("es-ES", "13 marzo 2024 lunes").as_deref(),
            Some("2024-03-13")
        );
    }

    #[test]
    fn times_use_the_locales_text_for_am_and_pm() {
        let loc = Locale::by_tag("en-US").unwrap();
        let dc = DateCtx::new(DateSystem::NULL_1899, Some(45658.0));
        assert_eq!(parse_date_text("1:30 PM", &loc, dc), Some(13.5 / 24.0));
        let es = Locale::by_tag("es-ES").unwrap();
        assert_eq!(parse_date_text("1:30 PM", &es, dc), None);
        assert_eq!(parse_date_text("1:30 p. m.", &es, dc), Some(13.5 / 24.0));
    }
}
