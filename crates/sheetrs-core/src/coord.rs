//! Cell coordinates: one canonical representation with A1 and R1C1 text views.

use thiserror::Error;

/// Rows in a default sheet (`A1:XFD1048576`).
pub const MAX_ROWS: u32 = 1_048_576;
/// Columns in a default sheet (`A1:XFD1048576`).
pub const MAX_COLS: u32 = 16_384;

const LETTER_COUNT: u32 = 26;

/// A zero-based cell position, used as the base of relative references.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct CellPos {
    /// Zero-based row index.
    pub row: u32,
    /// Zero-based column index.
    pub col: u32,
}

/// One axis of a cell reference: a zero-based index that is either fixed or relative.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct AxisRef {
    /// Zero-based index on the axis.
    pub index: u32,
    /// Whether the index stays fixed when the reference is copied (`$` in A1).
    pub absolute: bool,
}

/// A single-cell reference, independent of the text syntax it came from.
///
/// The parsers only return indices inside the sheet; the printing methods
/// assume the same, so building a value with an index beyond [`MAX_ROWS`] or
/// [`MAX_COLS`] is a caller error.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellRef {
    /// The row axis.
    pub row: AxisRef,
    /// The column axis.
    pub col: AxisRef,
}

/// A reference that cannot be turned into a [`CellRef`].
#[derive(Debug, Clone, PartialEq, Eq, Error)]
pub enum RefError {
    /// The text does not follow the A1 or R1C1 grammar.
    #[error("invalid reference syntax: {0:?}")]
    Syntax(String),
    /// The reference is well formed but lies outside the sheet.
    #[error("reference outside the sheet: {0:?}")]
    OutOfBounds(String),
}

impl CellRef {
    /// Parses an A1 reference such as `B2`, `$B$2` or `xfd1048576`.
    ///
    /// # Errors
    ///
    /// [`RefError::Syntax`] when the text is not an A1 reference and
    /// [`RefError::OutOfBounds`] when it names a cell outside the sheet.
    pub fn from_a1(text: &str) -> Result<Self, RefError> {
        let syntax = || RefError::Syntax(text.to_owned());
        let out_of_bounds = || RefError::OutOfBounds(text.to_owned());

        let (col_absolute, rest) = strip_dollar(text);
        let letters_len = rest.bytes().take_while(u8::is_ascii_alphabetic).count();
        let (letters, rest) = rest.split_at(letters_len);
        let (row_absolute, digits) = strip_dollar(rest);
        if letters.is_empty()
            || digits.is_empty()
            || !digits.bytes().all(|byte| byte.is_ascii_digit())
        {
            return Err(syntax());
        }

        let column_number = letters
            .bytes()
            .try_fold(0_u32, |number, letter| {
                number
                    .checked_mul(LETTER_COUNT)?
                    .checked_add(u32::from(letter.to_ascii_uppercase() - b'A') + 1)
            })
            .ok_or_else(out_of_bounds)?;
        let row_number: u32 = digits.parse().map_err(|_| out_of_bounds())?;
        if !(1..=MAX_COLS).contains(&column_number) || !(1..=MAX_ROWS).contains(&row_number) {
            return Err(out_of_bounds());
        }

        Ok(Self {
            row: AxisRef {
                index: row_number - 1,
                absolute: row_absolute,
            },
            col: AxisRef {
                index: column_number - 1,
                absolute: col_absolute,
            },
        })
    }

    /// Prints the reference in A1 syntax with uppercase column letters.
    #[must_use]
    pub fn to_a1(&self) -> String {
        let mut text = String::new();
        if self.col.absolute {
            text.push('$');
        }
        text.push_str(&column_letters(self.col.index));
        if self.row.absolute {
            text.push('$');
        }
        text.push_str(&(self.row.index + 1).to_string());
        text
    }

    /// Parses an R1C1 reference such as `R3C3`, `R[2]C[-1]` or `RC`.
    ///
    /// Relative axes are offsets from `base`; absolute axes are one-based.
    /// `R` and `C` are case-insensitive, and a zero offset (`R[0]`) is read as
    /// the empty form (`R`).
    ///
    /// # Errors
    ///
    /// [`RefError::Syntax`] when the text is not an R1C1 reference and
    /// [`RefError::OutOfBounds`] when it resolves to a cell outside the sheet.
    pub fn from_r1c1(text: &str, base: CellPos) -> Result<Self, RefError> {
        let syntax = || RefError::Syntax(text.to_owned());
        let rest = text.strip_prefix(['R', 'r']).ok_or_else(syntax)?;
        let (row_token, col_token) = rest.split_once(['C', 'c']).ok_or_else(syntax)?;
        Ok(Self {
            row: axis_from_r1c1(row_token, base.row, MAX_ROWS, text)?,
            col: axis_from_r1c1(col_token, base.col, MAX_COLS, text)?,
        })
    }

    /// Prints the reference in R1C1 syntax, with relative axes as offsets from `base`.
    #[must_use]
    pub fn to_r1c1(&self, base: CellPos) -> String {
        let mut text = String::from("R");
        push_axis_r1c1(&mut text, self.row, base.row);
        text.push('C');
        push_axis_r1c1(&mut text, self.col, base.col);
        text
    }
}

fn strip_dollar(text: &str) -> (bool, &str) {
    text.strip_prefix('$')
        .map_or((false, text), |rest| (true, rest))
}

fn column_letters(index: u32) -> String {
    let mut letters = Vec::new();
    let mut number = index + 1;
    while number > 0 {
        number -= 1;
        let Ok(offset) = u8::try_from(number % LETTER_COUNT) else {
            unreachable!("a remainder modulo {LETTER_COUNT} fits in u8");
        };
        letters.push(char::from(b'A' + offset));
        number /= LETTER_COUNT;
    }
    letters.iter().rev().collect()
}

fn axis_from_r1c1(token: &str, base: u32, limit: u32, text: &str) -> Result<AxisRef, RefError> {
    let syntax = || RefError::Syntax(text.to_owned());
    let out_of_bounds = || RefError::OutOfBounds(text.to_owned());

    let relative = |offset: i64| {
        i64::from(base)
            .checked_add(offset)
            .and_then(|index| u32::try_from(index).ok())
            .filter(|index| *index < limit)
            .map(|index| AxisRef {
                index,
                absolute: false,
            })
            .ok_or_else(out_of_bounds)
    };

    if token.is_empty() {
        return relative(0);
    }
    if let Some(inner) = token
        .strip_prefix('[')
        .and_then(|rest| rest.strip_suffix(']'))
    {
        let digits = inner.strip_prefix('-').unwrap_or(inner);
        if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
            return Err(syntax());
        }
        return inner
            .parse()
            .map_err(|_| out_of_bounds())
            .and_then(relative);
    }
    if !token.bytes().all(|byte| byte.is_ascii_digit()) {
        return Err(syntax());
    }
    let number: u32 = token.parse().map_err(|_| out_of_bounds())?;
    if (1..=limit).contains(&number) {
        Ok(AxisRef {
            index: number - 1,
            absolute: true,
        })
    } else {
        Err(out_of_bounds())
    }
}

fn push_axis_r1c1(text: &mut String, axis: AxisRef, base: u32) {
    if axis.absolute {
        text.push_str(&(axis.index + 1).to_string());
        return;
    }
    let offset = i64::from(axis.index) - i64::from(base);
    if offset != 0 {
        text.push('[');
        text.push_str(&offset.to_string());
        text.push(']');
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn rel(row: u32, col: u32) -> CellRef {
        CellRef {
            row: AxisRef {
                index: row,
                absolute: false,
            },
            col: AxisRef {
                index: col,
                absolute: false,
            },
        }
    }

    #[test]
    fn a1_parses_relative_and_absolute_axes() -> Result<(), RefError> {
        assert_eq!(CellRef::from_a1("A1")?, rel(0, 0));
        assert_eq!(CellRef::from_a1("AA10")?, rel(9, 26));
        let mixed = CellRef::from_a1("$B2")?;
        assert_eq!(
            mixed.col,
            AxisRef {
                index: 1,
                absolute: true
            }
        );
        assert_eq!(
            mixed.row,
            AxisRef {
                index: 1,
                absolute: false
            }
        );
        let both = CellRef::from_a1("$B$2")?;
        assert!(both.col.absolute && both.row.absolute);
        Ok(())
    }

    #[test]
    fn a1_is_case_insensitive_and_prints_uppercase() -> Result<(), RefError> {
        assert_eq!(CellRef::from_a1("b3")?.to_a1(), "B3");
        Ok(())
    }

    #[test]
    fn a1_accepts_the_sheet_corners() -> Result<(), RefError> {
        let last = CellRef::from_a1("XFD1048576")?;
        assert_eq!(
            (last.row.index, last.col.index),
            (MAX_ROWS - 1, MAX_COLS - 1)
        );
        assert_eq!(last.to_a1(), "XFD1048576");
        Ok(())
    }

    #[test]
    fn a1_rejects_cells_outside_the_sheet() {
        for text in ["XFE1", "A0", "A1048577", "ZZZZZZZZZZZZ1"] {
            assert!(
                matches!(CellRef::from_a1(text), Err(RefError::OutOfBounds(_))),
                "{text}"
            );
        }
    }

    #[test]
    fn a1_rejects_malformed_text() {
        for text in ["", "A", "1A", "$", "A$", "A-1", "A1 ", "A1B"] {
            assert!(
                matches!(CellRef::from_a1(text), Err(RefError::Syntax(_))),
                "{text:?}"
            );
        }
    }

    #[test]
    fn a1_round_trips() -> Result<(), RefError> {
        for text in ["A1", "Z9", "AA1", "$AZ$100", "C$3", "$C3", "XFD1048576"] {
            assert_eq!(CellRef::from_a1(text)?.to_a1(), text);
        }
        Ok(())
    }

    #[test]
    fn r1c1_prints_relative_axes_as_offsets_from_the_base() -> Result<(), RefError> {
        let base = CellPos { row: 0, col: 0 };
        assert_eq!(CellRef::from_a1("C3")?.to_r1c1(base), "R[2]C[2]");
        assert_eq!(CellRef::from_a1("A1")?.to_r1c1(base), "RC");
        let lower = CellPos { row: 4, col: 4 };
        assert_eq!(CellRef::from_a1("A1")?.to_r1c1(lower), "R[-4]C[-4]");
        Ok(())
    }

    #[test]
    fn r1c1_prints_absolute_axes_as_one_based_numbers() -> Result<(), RefError> {
        let base = CellPos { row: 0, col: 0 };
        assert_eq!(CellRef::from_a1("$C$3")?.to_r1c1(base), "R3C3");
        assert_eq!(CellRef::from_a1("C$3")?.to_r1c1(base), "R3C[2]");
        assert_eq!(CellRef::from_a1("$C3")?.to_r1c1(base), "R[2]C3");
        Ok(())
    }

    #[test]
    fn r1c1_parses_what_it_prints() -> Result<(), RefError> {
        let base = CellPos { row: 7, col: 3 };
        for text in ["A1", "$C$3", "C$3", "$C3", "XFD1048576", "D8"] {
            let cell = CellRef::from_a1(text)?;
            assert_eq!(
                CellRef::from_r1c1(&cell.to_r1c1(base), base)?,
                cell,
                "{text}"
            );
        }
        Ok(())
    }

    #[test]
    fn r1c1_accepts_offsets_up_to_the_sheet_edge() -> Result<(), RefError> {
        let base = CellPos { row: 0, col: 0 };
        let last = CellRef::from_r1c1("R[1048575]C[16383]", base)?;
        assert_eq!(
            (last.row.index, last.col.index),
            (MAX_ROWS - 1, MAX_COLS - 1)
        );
        for text in ["R[1048576]C", "RC[16384]"] {
            assert!(
                matches!(
                    CellRef::from_r1c1(text, base),
                    Err(RefError::OutOfBounds(_))
                ),
                "{text}"
            );
        }
        Ok(())
    }

    #[test]
    fn r1c1_is_case_insensitive() -> Result<(), RefError> {
        let base = CellPos { row: 2, col: 2 };
        assert_eq!(
            CellRef::from_r1c1("r[1]c[-1]", base)?,
            CellRef::from_r1c1("R[1]C[-1]", base)?
        );
        assert_eq!(CellRef::from_r1c1("rc", base)?, CellRef::from_a1("C3")?);
        Ok(())
    }

    #[test]
    fn r1c1_reads_a_zero_offset_as_the_empty_form() -> Result<(), RefError> {
        let base = CellPos { row: 2, col: 2 };
        assert_eq!(CellRef::from_r1c1("R[0]C[0]", base)?.to_r1c1(base), "RC");
        Ok(())
    }

    #[test]
    fn r1c1_classifies_numeric_overflow_as_out_of_bounds() {
        let base = CellPos { row: 0, col: 0 };
        for text in [
            "R[99999999999999999999]C",
            "R[-99999999999999999999]C",
            "R99999999999999999999C1",
        ] {
            assert!(
                matches!(
                    CellRef::from_r1c1(text, base),
                    Err(RefError::OutOfBounds(_))
                ),
                "{text}"
            );
        }
        for text in ["R[+1]C", "R[]C", "R[-]C", "R[1 ]C"] {
            assert!(
                matches!(CellRef::from_r1c1(text, base), Err(RefError::Syntax(_))),
                "{text}"
            );
        }
    }

    #[test]
    fn r1c1_rejects_a_base_outside_the_sheet() {
        let outside = CellPos {
            row: MAX_ROWS,
            col: 0,
        };
        assert!(matches!(
            CellRef::from_r1c1("RC", outside),
            Err(RefError::OutOfBounds(_))
        ));
    }

    #[test]
    fn a1_normalizes_leading_zeros() -> Result<(), RefError> {
        assert_eq!(CellRef::from_a1("A01")?.to_a1(), "A1");
        Ok(())
    }

    #[test]
    fn r1c1_rejects_out_of_sheet_and_malformed_text() {
        let base = CellPos { row: 0, col: 0 };
        for text in ["R[-1]C", "RC[-1]", "R0C1", "R1048577C1", "R1C16385"] {
            assert!(
                matches!(
                    CellRef::from_r1c1(text, base),
                    Err(RefError::OutOfBounds(_))
                ),
                "{text}"
            );
        }
        for text in ["", "X", "R1", "RC1X", "R[a]C", "R[1C"] {
            assert!(
                matches!(CellRef::from_r1c1(text, base), Err(RefError::Syntax(_))),
                "{text:?}"
            );
        }
    }
}
