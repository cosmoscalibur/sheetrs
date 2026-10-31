//! Value coercions and number/text conversion under a [`Locale`].

use crate::datetext::looks_like_numeric_date;
pub use crate::datetext::{DateCtx, parse_date_text};
use crate::options::Locale;
use crate::value::{ErrorKind, Value};
use std::fmt::Write as _;

type Res<T> = Result<T, ErrorKind>;

/// Number → text as the `General` format would show it: 15 significant
/// digits, positional between 1e-14 and 1e15, scientific otherwise with
/// the exponent in three digits (`E+017`), as LibreOffice writes it.
pub fn fmt_number(x: f64, loc: &Locale) -> String {
    if x == 0.0 {
        return "0".into();
    }
    let sci = format!("{:.14e}", x.abs());
    let (mant, exp) = sci.split_once('e').unwrap_or((&sci, "0"));
    let exp: i32 = exp.parse().unwrap_or(0);
    let digits: String = mant.chars().filter(char::is_ascii_digit).collect();
    let digits = digits.trim_end_matches('0');
    let digits = if digits.is_empty() { "0" } else { digits };
    let mut out = String::new();
    if x < 0.0 {
        out.push('-');
    }
    let positional = (-14..=15).contains(&exp);
    if positional {
        if exp >= 0 {
            let int_len = exp as usize + 1;
            if digits.len() <= int_len {
                out.push_str(digits);
                out.extend(std::iter::repeat_n('0', int_len - digits.len()));
            } else {
                out.push_str(&digits[..int_len]);
                out.push(loc.decimal);
                out.push_str(&digits[int_len..]);
            }
        } else {
            out.push('0');
            out.push(loc.decimal);
            out.extend(std::iter::repeat_n('0', (-exp - 1) as usize));
            out.push_str(digits);
        }
    } else {
        out.push_str(&digits[..1]);
        if digits.len() > 1 {
            out.push(loc.decimal);
            out.push_str(&digits[1..]);
        }
        let width = 3;
        let _ = write!(
            out,
            "E{}{:0width$}",
            if exp < 0 { '-' } else { '+' },
            exp.abs()
        );
    }
    out
}

/// Text → number: locale decimal/group separators, `%` suffix, and ISO dates.
pub fn parse_number(s: &str, loc: &Locale, dc: DateCtx) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    if looks_like_numeric_date(t, loc) {
        return parse_date_text(t, loc, dc);
    }
    let (t, pct) = match t.strip_suffix('%') {
        Some(rest) => (rest.trim_end(), true),
        None => (t, false),
    };
    // A group written with a space (French, Swedish…) is read with any of the spaces.
    let spaced = matches!(loc.group, ' ' | '\u{a0}' | '\u{202f}');
    let is_group = |c: char| c == loc.group || (spaced && matches!(c, ' ' | '\u{a0}' | '\u{202f}'));
    // A group separator must stand between groups of three digits: LibreOffice
    // does not read `1.5` as 15 under a locale that groups with `.`.
    if t.contains(is_group) {
        let unsigned = t.trim_start_matches(['+', '-']);
        let int = unsigned.split(loc.decimal).next().unwrap_or("");
        let mut groups = int.split(is_group);
        let first_ok = groups
            .next()
            .is_some_and(|g| (1..=3).contains(&g.len()) && g.chars().all(|c| c.is_ascii_digit()));
        let rest_ok = groups.all(|g| g.len() == 3 && g.chars().all(|c| c.is_ascii_digit()));
        let tail_clean = !unsigned
            .split(loc.decimal)
            .skip(1)
            .any(|d| d.contains(is_group));
        if !(first_ok && rest_ok && tail_clean) {
            return parse_date_text(t, loc, dc);
        }
    }
    // A `.` or `,` that is neither the decimal nor the group of the locale is not a number's.
    if t.chars()
        .any(|c| matches!(c, '.' | ',') && c != loc.decimal && !is_group(c))
    {
        return parse_date_text(t, loc, dc);
    }
    let mut clean = String::with_capacity(t.len());
    for ch in t.chars() {
        if is_group(ch) || ch == '\u{a0}' {
            continue;
        }
        clean.push(if ch == loc.decimal { '.' } else { ch });
    }
    let valid = clean
        .chars()
        .all(|c| c.is_ascii_digit() || matches!(c, '.' | '-' | '+' | 'e' | 'E'))
        && clean.chars().any(|c| c.is_ascii_digit());
    if valid && let Ok(n) = clean.parse::<f64>() {
        return Some(if pct { n / 100.0 } else { n });
    }
    parse_date_text(t, loc, dc)
}

/// Complete a two-digit year inside the 100-year window starting at `start`.
pub fn expand_year(y: i64, start: u16) -> i64 {
    if y >= 100 {
        return y;
    }
    let start = i64::from(start);
    let cand = start / 100 * 100 + y;
    if cand < start { cand + 100 } else { cand }
}

pub fn to_num(v: &Value, loc: &Locale, dc: DateCtx) -> Res<f64> {
    match v {
        Value::Number(n) => Ok(*n),
        Value::Empty => Ok(0.0),
        Value::Bool(b) => Ok(f64::from(u8::from(*b))),
        Value::Text(s) => parse_number(s, loc, dc).ok_or(ErrorKind::Value),
        Value::Error(e) => Err(*e),
        Value::Array(_) | Value::Ref(_) => Err(ErrorKind::Value),
    }
}

/// Value → text as LibreOffice writes it (booleans are the numbers 1/0, numbers use its
/// general format).
pub fn to_text(v: &Value, loc: &Locale) -> Res<String> {
    match v {
        Value::Text(s) => Ok(s.to_string()),
        Value::Number(n) => Ok(fmt_number(*n, loc)),
        Value::Bool(b) => Ok(if *b { "1" } else { "0" }.into()),
        Value::Empty => Ok(String::new()),
        Value::Error(e) => Err(*e),
        Value::Array(_) | Value::Ref(_) => Err(ErrorKind::Value),
    }
}

pub fn to_bool(v: &Value, loc: &Locale) -> Res<bool> {
    match v {
        Value::Bool(b) => Ok(*b),
        Value::Number(n) => Ok(*n != 0.0),
        Value::Empty => Ok(false),
        Value::Text(s) => {
            if s.eq_ignore_ascii_case("true") || s.eq_ignore_ascii_case(&loc.true_text) {
                Ok(true)
            } else if s.eq_ignore_ascii_case("false") || s.eq_ignore_ascii_case(&loc.false_text) {
                Ok(false)
            } else {
                Err(ErrorKind::Value)
            }
        }
        Value::Error(e) => Err(*e),
        Value::Array(_) | Value::Ref(_) => Err(ErrorKind::Value),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn general_format() {
        let en = Locale::en_us();
        assert_eq!(fmt_number(0.1 + 0.2, &en), "0.3");
        assert_eq!(fmt_number(1234.5, &en), "1234.5");
        assert_eq!(fmt_number(-42.0, &en), "-42");
        assert_eq!(fmt_number(0.00012, &en), "0.00012");
        assert_eq!(fmt_number(1.5, &Locale::es_es()), "1,5");
        assert_eq!(fmt_number(100.0, &en), "100");
        // Positional down to 1e-14, three-digit exponent.
        assert_eq!(fmt_number(1e-10, &en), "0.0000000001");
        assert_eq!(fmt_number(1.5e-10, &en), "0.00000000015");
        assert_eq!(fmt_number(1e-15, &en), "1E-015");
        assert_eq!(fmt_number(1e20, &en), "1E+020");
        assert_eq!(fmt_number(1e16, &en), "1E+016");
    }

    #[test]
    fn parsing() {
        let en = Locale::en_us();
        let es = Locale::es_es();
        // The clock stands at 2025-01-01.
        let s = DateCtx::new(crate::options::DateSystem::NULL_1899, Some(45658.0));
        assert_eq!(parse_number(" 1,234.5 ", &en, s), Some(1234.5));
        assert_eq!(parse_number("1.234,5", &es, s), Some(1234.5));
        assert_eq!(parse_number("50%", &en, s), Some(0.5));
        assert_eq!(parse_number("abc", &en, s), None);
        assert_eq!(parse_number("inf", &en, s), None);
        assert_eq!(parse_number("2024-01-01", &en, s), Some(45292.0));
        assert_eq!(parse_number("2024-01-01 12:00", &en, s), Some(45292.5));
        // Text dates follow the locale.
        assert_eq!(parse_number("3/1/2024", &en, s), Some(45352.0)); // March 1st
        assert_eq!(parse_number("3/1/2024", &es, s), Some(45294.0)); // January 3rd
        // The separator must be the pattern's: `3.1.2024` is not a date in Spanish.
        assert_eq!(parse_number("3.1.2024", &es, s), None);
        // A pattern without a year gives the year of the clock.
        assert_eq!(parse_number("3/1", &en, s), Some(45717.0)); // 2025-03-01
        assert_eq!(parse_number("1 Mar", &es, s), Some(45717.0));
        assert_eq!(parse_number("Mar 1", &en, s), Some(45717.0));
        // Month names follow the locale's order and words (LibreOffice 26.8).
        assert_eq!(parse_number("1 Mar 2024", &en, s), None); // month in the middle: day-first only
        assert_eq!(parse_number("1-Mar-2024", &en, s), Some(45352.0));
        assert_eq!(parse_number("Mar 1 2024", &en, s), Some(45352.0));
        assert_eq!(parse_number("March 1, 2024", &en, s), Some(45352.0));
        assert_eq!(parse_number("Mar. 1, 2024", &en, s), Some(45352.0));
        assert_eq!(parse_number("Monday Mar 1 2024", &en, s), Some(45352.0));
        assert_eq!(parse_number("Mar 2024", &en, s), Some(45352.0));
        assert_eq!(parse_number("2024-Mar-1", &en, s), Some(45352.0));
        assert_eq!(parse_number("Mar-1-2024", &en, s), None);
        assert_eq!(parse_number("2024 Mar 1", &en, s), None);
        assert_eq!(parse_number("1 Mar, 2024", &en, s), None);
        assert_eq!(parse_number("Mar 32 2024", &en, s), None);
        assert_eq!(parse_number("1 marzo 2024", &es, s), Some(45352.0));
        assert_eq!(parse_number("marzo 1 2024", &es, s), Some(45352.0));
        assert_eq!(parse_number("1 mar. 2024", &es, s), Some(45352.0));
        assert_eq!(parse_number("1 Mar 24", &es, s), Some(45352.0));
        assert_eq!(parse_number("Mar 24", &es, s), Some(45352.0)); // March of 2024
        assert_eq!(parse_number("marzo 1, 2024", &es, s), None);
        assert_eq!(parse_number("1 de marzo de 2024", &es, s), None);
        assert_eq!(parse_number("March 1 2024", &es, s), None); // English words are not Spanish
        assert_eq!(parse_number("3/1/24", &en, s), Some(45352.0));
        assert_eq!(parse_number("3/1/99", &en, s), Some(36220.0)); // 1999: window 1930–2029
        assert_eq!(parse_number("2/30/2024", &en, s), None);
        assert_eq!(
            parse_number("3/1/2024 1:30 PM", &en, s),
            Some(45352.0 + 13.5 / 24.0)
        );
        assert_eq!(parse_number("13:30", &en, s), Some(13.5 / 24.0));
        assert_eq!(parse_number("25:00", &en, s), None);
        assert_eq!(expand_year(29, 1930), 2029);
        assert_eq!(expand_year(30, 1930), 1930);
        assert_eq!(expand_year(5, 1950), 2005);
    }
}
