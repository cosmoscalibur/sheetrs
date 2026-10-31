//! Cell addressing: zero-based coordinates, A1 text parsing and formatting.

use serde::{Deserialize, Serialize};

pub use sheetrs_formula::addr::{
    Area, MAX_COLS, MAX_ROWS, a1, col_index, col_name, parse_area, same_name, split_sheet,
};

/// A cell of a sheet. `sheet` indexes `Program::sheets`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord, Serialize, Deserialize)]
pub struct CellRef {
    pub sheet: u16,
    pub row: u32,
    pub col: u32,
}

impl CellRef {
    pub fn new(sheet: u16, row: u32, col: u32) -> Self {
        Self { sheet, row, col }
    }
}

/// Inclusive, normalized rectangle (`r0 <= r1`, `c0 <= c1`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct Rect {
    pub sheet: u16,
    pub r0: u32,
    pub c0: u32,
    pub r1: u32,
    pub c1: u32,
}

impl Rect {
    pub fn new(sheet: u16, ra: u32, ca: u32, rb: u32, cb: u32) -> Self {
        Self {
            sheet,
            r0: ra.min(rb),
            c0: ca.min(cb),
            r1: ra.max(rb),
            c1: ca.max(cb),
        }
    }

    pub fn single(c: CellRef) -> Self {
        Self::new(c.sheet, c.row, c.col, c.row, c.col)
    }

    pub fn rows(&self) -> u32 {
        self.r1 - self.r0 + 1
    }

    pub fn cols(&self) -> u32 {
        self.c1 - self.c0 + 1
    }

    pub fn area(&self) -> u64 {
        u64::from(self.rows()) * u64::from(self.cols())
    }

    pub fn is_single(&self) -> bool {
        self.r0 == self.r1 && self.c0 == self.c1
    }

    pub fn contains(&self, c: CellRef) -> bool {
        c.sheet == self.sheet
            && (self.r0..=self.r1).contains(&c.row)
            && (self.c0..=self.c1).contains(&c.col)
    }

    pub fn top_left(&self) -> CellRef {
        CellRef::new(self.sheet, self.r0, self.c0)
    }
}

/// Bind an area written in a formula to a sheet.
pub fn bind(area: Area, sheet: u16) -> Rect {
    Rect::new(sheet, area.r0, area.c0, area.r1, area.c1)
}

/// Canonical text of a range with its sheet: `Sheet1!B2:C9`, quoting the sheet
/// name when it needs it (`'My Sheet'!A1`).
pub fn qualified(sheets: &[String], r: &Rect) -> String {
    let name = sheets.get(r.sheet as usize).map_or("?", String::as_str);
    let sheet = sheetrs_formula::print::sheet_text(name);
    format!("{sheet}!{}", rect_text(r))
}

/// Format a rectangle as A1 text (without sheet).
pub fn rect_text(r: &Rect) -> String {
    if r.is_single() {
        a1(r.r0, r.c0)
    } else {
        format!("{}:{}", a1(r.r0, r.c0), a1(r.r1, r.c1))
    }
}

/// Calc A1 notation: `[$]Sheet.A1[:[Sheet].B2]`, the sheet name quoted when it needs it.
/// Returns the sheet and the area as `A1[:B2]`; text without a sheet separator is a plain area.
fn calc_a1(text: &str) -> (Option<String>, String) {
    let text = text.strip_prefix('$').unwrap_or(text);
    // First reference: the sheet is what precedes the first `.` outside quotes.
    let (sheet, rest) = if let Some(q) = text.strip_prefix('\'') {
        let mut name = String::new();
        let mut chars = q.char_indices().peekable();
        let mut end = None;
        while let Some((i, c)) = chars.next() {
            if c == '\'' {
                if matches!(chars.peek(), Some((_, '\''))) {
                    name.push('\'');
                    chars.next();
                } else {
                    end = Some(i + 1);
                    break;
                }
            } else {
                name.push(c);
            }
        }
        match end.and_then(|i| q[i..].strip_prefix('.')) {
            Some(rest) => (Some(name), rest),
            None => return (None, text.to_string()),
        }
    } else {
        match text.split_once('.') {
            Some((n, rest)) if !n.is_empty() => (Some(n.to_string()), rest),
            _ => (None, text),
        }
    };
    // Second reference of a range: `A1:Sheet.B2`, `A1:.B2`, `A1:B2`.
    let rest = match rest.split_once(':') {
        Some((a, b)) => {
            let b = b.strip_prefix('$').unwrap_or(b);
            let b = match b.rsplit_once('.') {
                Some((_, cell)) => cell,
                None => b,
            };
            format!("{a}:{b}")
        }
        None => rest.to_string(),
    };
    (sheet, rest)
}

/// The sheet and the area of the text `INDIRECT` receives, read in the document's syntax:
/// Excel A1 (`Sheet!A1`, what an XLSX import sets) or Calc A1 (`Sheet.A1`, a document created
/// in LibreOffice). The other notation is not a reference. `None`: not a reference.
pub fn string_ref(text: &str, excel: bool) -> Option<(Option<String>, String)> {
    let text = text.trim();
    if excel {
        let (sheet, tail) = split_sheet(text);
        // Calc's `Sheet.A1` is not a reference here.
        (!tail.contains('.')).then(|| (sheet, tail.to_string()))
    } else {
        let (sheet, tail) = calc_a1(text);
        // Excel's `Sheet!A1` is not a reference here.
        (sheet.is_some() || !tail.contains('!')).then_some((sheet, tail))
    }
}
