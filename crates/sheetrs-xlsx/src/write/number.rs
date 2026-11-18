//! The text `LibreOffice` writes for a number in an XLSX file.

/// The number as `LibreOffice` writes it: 15 significant digits, no trailing zeros, and an
/// exponent of at least three digits where the number is too small (below `1E-4`) or too large
/// (from `1E+17`) to be written without one (`1E+021`, `1E-007`). A number from `1E+15` up to
/// but not including `1E+16` keeps all its digits.
pub(crate) fn number_text(number: f64) -> String {
    let Some(found) = sheetrs_io::fifteen_digits(number) else {
        return "0".to_owned();
    };
    if (-4..=16).contains(&found.exponent) {
        found.plain(number)
    } else {
        found.scientific(3)
    }
}
