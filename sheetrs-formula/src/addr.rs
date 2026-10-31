//! A1 addressing: column letters, area text, sheet qualifiers.

pub const MAX_ROWS: u32 = 1_048_576;
pub const MAX_COLS: u32 = 16_384;

/// Column letters to a 0-based index (`A` → 0). `None` unless 1–3 ASCII
/// letters within `XFD`.
pub fn col_index(s: &str) -> Option<u32> {
    if s.is_empty() || s.len() > 3 || !s.bytes().all(|b| b.is_ascii_alphabetic()) {
        return None;
    }
    let n = s.bytes().fold(0u32, |a, b| {
        a * 26 + u32::from(b.to_ascii_uppercase() - b'A' + 1)
    });
    (n <= MAX_COLS).then(|| n - 1)
}

/// 0-based column index to letters, allocation free.
pub fn col_name(index: u32) -> impl std::fmt::Display {
    std::fmt::from_fn(move |f| {
        let mut i = index;
        let mut buf = [0u8; 4];
        let mut n = 0;
        i += 1;
        while i > 0 {
            i -= 1;
            buf[n] = b'A' + (i % 26) as u8;
            n += 1;
            i /= 26;
        }
        buf[..n].reverse();
        f.write_str(std::str::from_utf8(&buf[..n]).unwrap_or("?"))
    })
}

/// `A1` text for a zero-based position.
pub fn a1(row: u32, col: u32) -> String {
    format!("{}{}", col_name(col), row + 1)
}

/// An area as written, with open ends (`A:A`, `1:3`) already expanded to the
/// sheet limits. Absolute markers are dropped.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Area {
    pub r0: u32,
    pub c0: u32,
    pub r1: u32,
    pub c1: u32,
}

/// `[$]COL[$]ROW` → `(row, col, row_abs, col_abs)`, zero-based.
pub fn parse_cell(s: &str) -> Option<(u32, u32, bool, bool)> {
    let (s, col_abs) = match s.strip_prefix('$') {
        Some(r) => (r, true),
        None => (s, false),
    };
    let split = s.find(|c: char| !c.is_ascii_alphabetic())?;
    let (letters, digits) = s.split_at(split);
    let (digits, row_abs) = match digits.strip_prefix('$') {
        Some(r) => (r, true),
        None => (digits, false),
    };
    if digits.is_empty() || !digits.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let row: u32 = digits.parse().ok()?;
    if row == 0 || row > MAX_ROWS {
        return None;
    }
    Some((row - 1, col_index(letters)?, row_abs, col_abs))
}

/// `[$]COL` → `(col, abs)`.
pub fn parse_col(s: &str) -> Option<(u32, bool)> {
    let (t, abs) = match s.strip_prefix('$') {
        Some(r) => (r, true),
        None => (s, false),
    };
    Some((col_index(t)?, abs))
}

/// `[$]ROW` → `(row, abs)`, zero-based.
pub fn parse_row(s: &str) -> Option<(u32, bool)> {
    let (t, abs) = match s.strip_prefix('$') {
        Some(r) => (r, true),
        None => (s, false),
    };
    if t.is_empty() || !t.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let r: u32 = t.parse().ok()?;
    (1..=MAX_ROWS).contains(&r).then(|| (r - 1, abs))
}

/// Parse `A1`, `$A$1`, `A1:B2`, `A:C` or `2:5` (no sheet qualifier).
pub fn parse_area(s: &str) -> Option<Area> {
    match s.split_once(':') {
        None => {
            let (r, c, ..) = parse_cell(s)?;
            Some(Area {
                r0: r,
                c0: c,
                r1: r,
                c1: c,
            })
        }
        Some((a, b)) => {
            if let (Some((ra, ca, ..)), Some((rb, cb, ..))) = (parse_cell(a), parse_cell(b)) {
                return Some(Area {
                    r0: ra.min(rb),
                    c0: ca.min(cb),
                    r1: ra.max(rb),
                    c1: ca.max(cb),
                });
            }
            if let (Some((ca, _)), Some((cb, _))) = (parse_col(a), parse_col(b)) {
                return Some(Area {
                    r0: 0,
                    c0: ca.min(cb),
                    r1: MAX_ROWS - 1,
                    c1: ca.max(cb),
                });
            }
            if let (Some((ra, _)), Some((rb, _))) = (parse_row(a), parse_row(b)) {
                return Some(Area {
                    r0: ra.min(rb),
                    c0: 0,
                    r1: ra.max(rb),
                    c1: MAX_COLS - 1,
                });
            }
            None
        }
    }
}

/// Split an optional sheet qualifier: `Sheet1!A1`, `'My Sheet'!A1:B2`
/// (`''` escapes a quote). Returns the unquoted sheet name and the remainder.
pub fn split_sheet(s: &str) -> (Option<String>, &str) {
    if let Some(rest) = s.strip_prefix('\'') {
        let mut name = String::new();
        let mut it = rest.char_indices().peekable();
        while let Some((i, ch)) = it.next() {
            if ch == '\'' {
                if matches!(it.peek(), Some((_, '\''))) {
                    name.push('\'');
                    it.next();
                } else if let Some(tail) = rest[i + 1..].strip_prefix('!') {
                    return (Some(name), tail);
                } else {
                    break;
                }
            } else {
                name.push(ch);
            }
        }
        return (None, s);
    }
    match s.rsplit_once('!') {
        Some((sheet, tail)) if !sheet.is_empty() => (Some(sheet.to_string()), tail),
        _ => (None, s),
    }
}

/// Sheet, name and function names are case-insensitive, for every script.
pub fn same_name(a: &str, b: &str) -> bool {
    // Both ASCII (the usual case, and the hot one when lowering): the lower-case mapping of a
    // character is the ASCII one. A non-ASCII text can match an ASCII one (the Kelvin sign
    // lowers to `k`), so a mixed pair takes the general path.
    if a.is_ascii() && b.is_ascii() {
        return a.eq_ignore_ascii_case(b);
    }
    a.chars()
        .flat_map(char::to_lowercase)
        .eq(b.chars().flat_map(char::to_lowercase))
}

/// Whether a sheet name must be quoted in a reference.
pub fn needs_quotes(name: &str) -> bool {
    name.is_empty()
        || name.starts_with(|c: char| c.is_ascii_digit())
        || !name.chars().all(|c| c.is_alphanumeric() || c == '_')
        // Looks like a cell reference (`A1`, `R1C1`): Excel quotes those too.
        || parse_cell(name).is_some()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn columns_roundtrip() {
        for (s, i) in [("A", 0), ("Z", 25), ("AA", 26), ("XFD", 16383)] {
            assert_eq!(col_index(s), Some(i));
            assert_eq!(col_name(i).to_string(), s);
        }
        assert_eq!(col_index("XFE"), None);
    }

    #[test]
    fn cells_and_areas() {
        assert_eq!(parse_cell("$B$2"), Some((1, 1, true, true)));
        assert_eq!(parse_cell("B$2"), Some((1, 1, true, false)));
        assert_eq!(
            parse_area("B3:A1"),
            Some(Area {
                r0: 0,
                c0: 0,
                r1: 2,
                c1: 1
            })
        );
        assert_eq!(
            parse_area("A:B").map(|a| (a.r0, a.r1)),
            Some((0, MAX_ROWS - 1))
        );
        assert_eq!(
            parse_area("2:3").map(|a| (a.c0, a.c1)),
            Some((0, MAX_COLS - 1))
        );
        assert_eq!(parse_area("A0"), None);
        assert_eq!(parse_area("SUM"), None);
    }

    #[test]
    fn sheet_qualifiers() {
        assert_eq!(split_sheet("S1!A1"), (Some("S1".into()), "A1"));
        assert_eq!(split_sheet("'It''s'!B2:C3"), (Some("It's".into()), "B2:C3"));
        assert_eq!(split_sheet("A1"), (None, "A1"));
        assert!(needs_quotes("My Sheet") && needs_quotes("A1") && needs_quotes("1x"));
        assert!(!needs_quotes("Sheet1") && !needs_quotes("Hoja_Ñ"));
        assert!(same_name("Hoja Ñ", "HOJA ñ"));
        assert!(same_name("Sheet1", "SHEET1") && !same_name("Sheet1", "Sheet2"));
        // An ASCII text and a non-ASCII one that lowers to it.
        assert!(same_name("k", "\u{212a}") && same_name("K", "\u{212a}"));
    }
}
