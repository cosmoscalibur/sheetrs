//! The text `LibreOffice` writes for a number in an XLSX file.

/// The number as `LibreOffice` writes it: 15 significant digits, no trailing zeros, and an
/// exponent of at least three digits where the number is too small or too large to be written
/// without one (`1E+021`, `1E-007`). A number that 15 digits would make infinite keeps the
/// digits that tell it apart.
pub(crate) fn number_text(number: f64) -> String {
    if number == 0.0 {
        return "0".to_owned();
    }
    let scientific = format!("{number:.14e}");
    let (mantissa, exponent) = scientific.split_once('e').unwrap_or((&scientific, "0"));
    let exponent: i32 = exponent.parse().unwrap_or(0);
    let (sign, mantissa) = mantissa
        .strip_prefix('-')
        .map_or(("", mantissa), |rest| ("-", rest));
    let digits: String = mantissa.chars().filter(char::is_ascii_digit).collect();
    let digits = digits.trim_end_matches('0');
    if !(-4..15).contains(&exponent) {
        let text = scientific_text(sign, digits, exponent);
        // Rounded to 15 digits, the largest numbers no longer fit a double.
        if text_to_number(&text).is_some_and(f64::is_infinite) {
            return format!("{number:E}")
                .replace("E-", "E-0")
                .replace('E', "E+")
                .replace("E+-", "E-");
        }
        return text;
    }
    let point = usize::try_from(exponent + 1).ok();
    let body = match point {
        Some(whole) if whole >= digits.len() => {
            format!("{digits}{}", "0".repeat(whole - digits.len()))
        }
        Some(whole) => format!("{}.{}", &digits[..whole], &digits[whole..]),
        None => format!(
            "0.{}{digits}",
            "0".repeat(usize::try_from(-exponent - 1).unwrap_or(0))
        ),
    };
    format!("{sign}{body}")
}

fn scientific_text(sign: &str, digits: &str, exponent: i32) -> String {
    let (first, rest) = digits.split_at(1.min(digits.len()));
    let fraction = if rest.is_empty() {
        String::new()
    } else {
        format!(".{rest}")
    };
    let exponent_sign = if exponent < 0 { '-' } else { '+' };
    format!(
        "{sign}{first}{fraction}E{exponent_sign}{:03}",
        exponent.abs()
    )
}

fn text_to_number(text: &str) -> Option<f64> {
    text.replace('E', "e").parse().ok()
}
