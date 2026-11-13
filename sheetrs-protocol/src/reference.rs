//! A range as a caller writes it: `Sheet1!B2:C9`, `'My Sheet'!A:A`, `B2`.

use crate::error::{Error, Result};
use serde::{Deserialize, Serialize};
use sheetrs_formula::addr::{Area, parse_area, split_sheet};

/// A range of one sheet. Without a sheet it is the first one of the workbook.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(try_from = "String", into = "String")]
pub struct Reference {
    /// The name of the sheet, as written (not quoted).
    pub sheet: Option<String>,
    /// The cells: zero-based, with open ends (`A:A`, `3:5`) reaching the limits of a sheet.
    pub area: Area,
    text: String,
}

impl std::str::FromStr for Reference {
    type Err = Error;

    /// A range as written: `Sheet1!B2:C9`, `'My Sheet'!A:A`, `B2`.
    ///
    /// # Errors
    ///
    /// When the text is not a range.
    fn from_str(text: &str) -> Result<Self> {
        let text = text.trim();
        let (sheet, tail) = split_sheet(text);
        let area =
            parse_area(tail).ok_or_else(|| Error::Request(format!("`{text}` is not a range")))?;
        Ok(Self {
            sheet,
            area,
            text: text.to_string(),
        })
    }
}

impl Reference {
    /// The range as it was written.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.text
    }
}

impl TryFrom<String> for Reference {
    type Error = Error;

    fn try_from(text: String) -> Result<Self> {
        text.parse()
    }
}

impl From<Reference> for String {
    fn from(r: Reference) -> Self {
        r.text
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_range_with_and_without_a_sheet() {
        let r = "'My Sheet'!B2:C9".parse::<Reference>().unwrap();
        assert_eq!(r.sheet.as_deref(), Some("My Sheet"));
        assert_eq!((r.area.r0, r.area.c0, r.area.r1, r.area.c1), (1, 1, 8, 2));
        assert_eq!("A1".parse::<Reference>().unwrap().sheet, None);
        assert!("not a range".parse::<Reference>().is_err());
    }
}
