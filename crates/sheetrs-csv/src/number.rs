//! The numbers a field holds as plain decimals.

/// The characters skipped around a number.
fn is_padding(character: char) -> bool {
    character == ' ' || character == '\u{a0}'
}

/// The number a field stands for, if it is `[+-]digits[.digits][e[+-]digits]` or
/// `[+-].digits[...]` between spaces. A number that overflows or underflows to zero is not one.
pub(crate) fn number(text: &str) -> Option<f64> {
    let token = text.trim_matches(is_padding);
    let bytes = token.as_bytes();
    let mut at = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    let digits = |from: usize| {
        bytes.get(from..).map_or(0, |tail| {
            tail.iter().take_while(|byte| byte.is_ascii_digit()).count()
        })
    };
    let integer = digits(at);
    at += integer;
    let mut mantissa_digits = integer;
    if bytes.get(at) == Some(&b'.') {
        let fraction = digits(at + 1);
        at += 1 + fraction;
        mantissa_digits += fraction;
    }
    if mantissa_digits == 0 {
        return None;
    }
    if matches!(bytes.get(at), Some(b'e' | b'E')) {
        let sign = usize::from(matches!(bytes.get(at + 1), Some(b'+' | b'-')));
        let exponent = digits(at + 1 + sign);
        if exponent == 0 {
            return None;
        }
        at += 1 + sign + exponent;
    }
    if at != bytes.len() {
        return None;
    }
    let value: f64 = token.parse().ok()?;
    let mantissa_is_zero = token
        .bytes()
        .take_while(|byte| !matches!(byte, b'e' | b'E'))
        .all(|byte| !byte.is_ascii_digit() || byte == b'0');
    (value.is_finite() && (value != 0.0 || mantissa_is_zero)).then_some(value)
}
