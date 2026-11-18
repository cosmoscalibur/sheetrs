//! The text `LibreOffice` writes for a number in an XLSX file.

/// The number as `LibreOffice` writes it: 15 significant digits, no trailing zeros, and an
/// exponent of at least three digits where the number is too small (below `1E-4`) or too large
/// (from `1E+17`) to be written without one (`1E+021`, `1E-007`). A number from `1E+15` up to
/// but not including `1E+16` keeps all its digits.
pub(crate) fn number_text(number: f64) -> String {
    let Some(found) = sheetrs_io::fifteen_digits(number) else {
        return "0".to_owned();
    };
    let (sign, digits, exponent) = (
        if found.negative { "-" } else { "" },
        found.digits.as_str(),
        found.exponent,
    );
    if !(-4..=16).contains(&exponent) {
        return scientific_text(sign, digits, exponent);
    }
    if exponent == 15 {
        // A whole part of sixteen digits is written in full.
        return format!("{number:.0}");
    }
    let point = usize::try_from(exponent + 1).ok();
    let body = match point {
        Some(whole) if whole >= digits.len() => {
            format!("{digits}{}", "0".repeat(whole - digits.len()))
        }
        Some(0) => format!("0.{digits}"),
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
