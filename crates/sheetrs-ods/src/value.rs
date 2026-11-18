//! The value of a cell as `LibreOffice` 26.8 loads it from the attributes and paragraphs of an
//! ODS cell.

use sheetrs_core::DateSystem;
use sheetrs_io::{DateRules, XML_WHITESPACE, decimal_prefix, iso_date};

use crate::tokens::Attrs;

/// What a cell holds.
#[derive(Debug, Clone, PartialEq)]
pub(crate) enum Scalar {
    Empty,
    Number(f64),
    Bool(bool),
    Text(String),
}

/// The value of a cell, and whether it is a cached text that `LibreOffice` keeps when the cell
/// also has a formula.
pub(crate) fn scalar(attrs: &Attrs, paragraphs: &[String], system: DateSystem) -> (Scalar, bool) {
    // The `calcext` type overrides the `office` one, except for `error`, which is a string
    // the cell shows and has no type of its own.
    let kind = match attrs.calcext_type.as_deref() {
        None | Some("error") => attrs.value_type.as_deref(),
        other => other,
    };
    let stored = attrs.value.as_deref();
    match kind {
        Some("float" | "percentage" | "currency") => {
            (Scalar::Number(stored.map_or(0.0, number)), false)
        }
        Some("string") => string(attrs.string_value.as_deref(), paragraphs),
        Some("boolean") => (boolean(attrs.boolean_value.as_deref().or(stored)), false),
        Some("date") => (
            Scalar::Number(match attrs.date_value.as_deref() {
                Some(text) => iso_date(text, system, DateRules::ODS).unwrap_or(0.0),
                None => stored.map_or(0.0, number),
            }),
            false,
        ),
        Some("time") => (
            Scalar::Number(match attrs.time_value.as_deref() {
                Some(text) => duration(text).unwrap_or(0.0),
                None => stored.map_or(0.0, number),
            }),
            false,
        ),
        None if !paragraphs.is_empty() => (Scalar::Text(paragraphs.join("\n")), false),
        _ => (Scalar::Empty, false),
    }
}

fn string(attribute: Option<&str>, paragraphs: &[String]) -> (Scalar, bool) {
    match attribute.filter(|text| !text.is_empty()) {
        Some(text) => (Scalar::Text(text.to_owned()), true),
        None if paragraphs.is_empty() => (Scalar::Empty, false),
        None => (Scalar::Text(paragraphs.join("\n")), false),
    }
}

/// A stored boolean is `true`, or the number the text starts with; it is a boolean of the model
/// only when that is exactly 0 or 1.
fn boolean(attribute: Option<&str>) -> Scalar {
    let value = match attribute {
        Some("true") => 1.0,
        Some(text) => number(text),
        None => 0.0,
    };
    if value.to_bits() == 1.0_f64.to_bits() {
        Scalar::Bool(true)
    } else if value.to_bits() == 0.0_f64.to_bits() {
        Scalar::Bool(false)
    } else {
        Scalar::Number(value)
    }
}

/// The number a stored text stands for: after optional leading whitespace and a sign, the
/// longest prefix that is a decimal number; text without one, and a number that overflows, is
/// zero.
pub(crate) fn number(text: &str) -> f64 {
    let rest = text.trim_start_matches(XML_WHITESPACE);
    let (negative, unsigned) = match rest.strip_prefix('-') {
        Some(after) => (true, after),
        None => (false, rest.strip_prefix('+').unwrap_or(rest)),
    };
    let magnitude = decimal_prefix(unsigned)
        .and_then(|prefix| prefix.parse::<f64>().ok())
        .unwrap_or(0.0);
    if !magnitude.is_finite() {
        0.0
    } else if negative {
        -magnitude
    } else {
        magnitude
    }
}

/// The days of a `PnDTnHnMnS` duration, or `None` when the text is not one.
///
/// A decimal point can follow any amount of the time part, and then the amount counts as
/// seconds whatever its unit is.
pub(crate) fn duration(text: &str) -> Option<f64> {
    let (negative, rest) = match text.strip_prefix('-') {
        Some(after) => (true, after),
        None => (false, text),
    };
    let rest = rest.strip_prefix(['P', 'p'])?;
    let (date, time) = match rest.find(['T', 't']) {
        Some(at) => (rest.get(..at)?, rest.get(at + 1..)?),
        None => (rest, ""),
    };
    let mut days = 0.0;
    for amount in amounts(date)? {
        match (amount.unit, amount.fraction) {
            ('D', None) => days = amount.whole,
            _ => return None,
        }
    }
    let (mut hours, mut minutes, mut seconds) = (0.0, 0.0, 0.0);
    for amount in amounts(time)? {
        match (amount.unit, amount.fraction) {
            ('H', None) => hours = amount.whole,
            ('M', None) => minutes = amount.whole,
            ('H' | 'M' | 'S', Some(fraction)) => seconds = amount.whole + fraction,
            ('S', None) => seconds = amount.whole,
            _ => return None,
        }
    }
    let total = (days * 86_400.0 + hours * 3600.0 + minutes * 60.0 + seconds) / 86_400.0;
    Some(if negative { -total } else { total })
}

/// The most an integer part of a duration may reach before another digit is added.
const DURATION_DIGIT_LIMIT: f64 = 214_748_363.0;

/// One `amount` and upper-case `unit` of a duration part, with the fraction after a decimal
/// point if there is one.
struct Amount {
    whole: f64,
    fraction: Option<f64>,
    unit: char,
}

fn amounts(part: &str) -> Option<Vec<Amount>> {
    let mut found = Vec::new();
    let mut rest = part;
    while !rest.is_empty() {
        let digits = rest.bytes().take_while(u8::is_ascii_digit).count();
        let (whole_digits, tail) = rest.split_at(digits);
        let mut whole = 0.0_f64;
        for digit in whole_digits.bytes() {
            if whole > DURATION_DIGIT_LIMIT {
                return None;
            }
            whole = whole * 10.0 + f64::from(digit - b'0');
        }
        let (fraction, tail) = match tail.strip_prefix(['.', ',']) {
            Some(after) => {
                let count = after.bytes().take_while(u8::is_ascii_digit).count();
                let (fraction_digits, tail) = after.split_at(count);
                let fraction = if fraction_digits.is_empty() {
                    0.0
                } else {
                    format!("0.{fraction_digits}").parse().ok()?
                };
                (Some(fraction), tail)
            }
            None => (None, tail),
        };
        if whole_digits.is_empty() && fraction.is_none() {
            return None;
        }
        let unit = tail.chars().next()?;
        found.push(Amount {
            whole,
            fraction,
            unit: unit.to_ascii_uppercase(),
        });
        rest = tail.get(unit.len_utf8()..)?;
    }
    Some(found)
}
