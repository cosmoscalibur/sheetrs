//! What a command does to the layout of a workbook: rows and columns deleted from sheets, sheets
//! deleted. Everything is expressed in the coordinates of the input file, so the user needs no
//! idea of the positions the cells will have afterwards.
//!
//! References follow `LibreOffice` 26.8, which `sheetrs_formula::edit` implements: a reference to
//! deleted cells is `#REF!`, a range shrinks around the deletion, whole rows and columns do not
//! move when the other axis changes.

use super::error::{Error, Result, ScopeError};
use super::{ValueScope, ValueTarget};
use sheetrs_formula::edit::Edit;
use sheetrs_formula::parse::parse_spanned;
use std::collections::{HashMap, HashSet};

/// Disjoint, ordered, inclusive 0-based bands of deleted rows (or columns), and how many indexes
/// the bands before each one delete, so a position is found by search and not by a walk.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct Bands(Vec<(u32, u32)>, Vec<u32>);

impl Bands {
    /// The bands of `ranges`, which may overlap or touch.
    #[must_use]
    pub fn new(mut ranges: Vec<(u32, u32)>) -> Self {
        ranges.sort_unstable();
        let mut merged: Vec<(u32, u32)> = Vec::with_capacity(ranges.len());
        for (first, last) in ranges {
            match merged.last_mut() {
                Some(band) if first <= band.1.saturating_add(1) => band.1 = band.1.max(last),
                _ => merged.push((first, last)),
            }
        }
        let mut before = 0u32;
        let deleted_before_each = merged
            .iter()
            .map(|&(first, last)| {
                let n = before;
                before += last - first + 1;
                n
            })
            .collect();
        Self(merged, deleted_before_each)
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.0.is_empty()
    }

    /// How many deleted indexes are before `index`.
    fn deleted_before(&self, index: u32) -> u32 {
        // The bands that start before `index`: all but the last are whole before it.
        let started = self.0.partition_point(|&(first, _)| first < index);
        match started.checked_sub(1) {
            None => 0,
            Some(i) => {
                let (first, last) = self.0[i];
                self.1[i] + index.min(last.saturating_add(1)) - first
            }
        }
    }

    /// How many of `first..first + count` stay.
    pub(crate) fn kept_in(&self, first: u32, count: u32) -> u32 {
        let end = first.saturating_add(count);
        count - (self.deleted_before(end) - self.deleted_before(first))
    }

    /// Whether `index` is deleted.
    pub(crate) fn contains(&self, index: u32) -> bool {
        self.0
            .binary_search_by(|&(first, last)| {
                if index < first {
                    std::cmp::Ordering::Greater
                } else if index > last {
                    std::cmp::Ordering::Less
                } else {
                    std::cmp::Ordering::Equal
                }
            })
            .is_ok()
    }

    /// The index that the first row or column at or after `index` that stays has once the bands
    /// are gone.
    pub(crate) fn shifted(&self, index: u32) -> u32 {
        index - self.deleted_before(index)
    }

    /// Where `index` is once the bands are gone, `None` when it is deleted.
    pub(crate) fn moved(&self, index: u32) -> Option<u32> {
        (!self.contains(index)).then(|| index - self.deleted_before(index))
    }

    /// The bounds of `first..=last` once the bands are gone, `None` when nothing of it remains.
    pub(crate) fn span(&self, first: u32, last: u32) -> Option<(u32, u32)> {
        let band_of = |index: u32| self.0.iter().find(|&&(a, b)| (a..=b).contains(&index));
        let first = band_of(first).map_or(first, |&(_, b)| b.saturating_add(1));
        let last = match band_of(last) {
            Some(&(a, _)) => a.checked_sub(1)?,
            None => last,
        };
        if first > last {
            return None;
        }
        Some((self.moved(first)?, self.moved(last)?))
    }

    /// The index in the input of the row (or column) that is at `current` once the bands are
    /// gone: what a position given after an earlier deletion refers to.
    #[must_use]
    pub fn input_index(&self, current: u32) -> u32 {
        let mut index = current;
        for &(first, last) in &self.0 {
            if first <= index {
                index = index.saturating_add(last - first + 1);
            } else {
                break;
            }
        }
        index
    }

    /// The bands as inclusive `(first, last)` ranges, in order.
    pub fn ranges(&self) -> impl Iterator<Item = (u32, u32)> + '_ {
        self.0.iter().copied()
    }

    /// The bands as `(first, count)`, the last one first.
    pub(crate) fn last_first(&self) -> impl Iterator<Item = (u32, u32)> + '_ {
        self.0
            .iter()
            .rev()
            .map(|&(first, last)| (first, last - first + 1))
    }
}

/// The error of a command that meets a 3D reference (a span of sheets, `Sheet1:Sheet3!A1`) while it
/// deletes rows or columns or removes sheets: what the reference becomes is not implemented, and
/// leaving it as it is would be a wrong formula without a word.
pub(super) fn three_d_unsupported(formula: &str) -> Error {
    Error::ThreeDReference {
        formula: formula.to_string(),
    }
}

/// Finds, in text read in pieces, a reference across sheets: a `:` followed by a sheet name and
/// `!` (`Sheet1:Sheet3!A1`, an XLSX) or `.` (`$S.A1:$T.B2`, an ODS), or a quoted sheet name that
/// holds a `:` (`'Sheet 1:Sheet 3'!A1`). It also matches texts that are not one (`A1:A5.5`): the
/// cost of that is a formula read that did not need it, where the formula is then refused if it
/// is one. A part where it finds none is copied without reading, so it must not miss one.
#[derive(Default)]
struct SpanScan {
    state: ScanState,
    found: bool,
}

#[derive(Default, Clone, Copy, PartialEq)]
enum ScanState {
    #[default]
    Idle,
    AfterColon,
    Name,
    Quoted,
    QuoteEnd,
    /// Inside a quoted name that has no `:` yet, and one that has.
    InQuote,
    InQuoteColon,
    /// A quote met inside a quoted name that held a `:`: closing it, or the first of a doubled one.
    QuoteEndColon,
}

impl SpanScan {
    // A transition table: one arm per (state, byte class), kept apart even where two arms share
    // a next state so that each row can be checked against the diagram on its own.
    #[allow(clippy::match_same_arms)]
    fn feed(&mut self, bytes: &[u8]) {
        use ScanState::{
            AfterColon, Idle, InQuote, InQuoteColon, Name, QuoteEnd, QuoteEndColon, Quoted,
        };
        let name_byte = |b: u8| b.is_ascii_alphanumeric() || b == b'_' || b >= 0x80;
        for &b in bytes {
            self.state = match (self.state, b) {
                (InQuote, b':') => InQuoteColon,
                (InQuote, b'\'') => Idle,
                (InQuote, _) => InQuote,
                (InQuoteColon, b'\'') => QuoteEndColon,
                (InQuoteColon, _) => InQuoteColon,
                (QuoteEndColon, b'\'') => InQuoteColon,
                (QuoteEndColon, b'!' | b'.') => {
                    self.found = true;
                    Idle
                }
                (QuoteEndColon, _) => Idle,
                (Quoted, b'\'') => QuoteEnd,
                (Quoted, _) => Quoted,
                (QuoteEnd, b'\'') => Quoted,
                (Name | QuoteEnd, b'!' | b'.') => {
                    self.found = true;
                    Idle
                }
                (_, b':') => AfterColon,
                (AfterColon, b'$') => AfterColon,
                (AfterColon, b'\'') => Quoted,
                (AfterColon | Name, b) if name_byte(b) => Name,
                (Idle, b'\'') => InQuote,
                _ => Idle,
            };
        }
    }
}

/// A reader that shows what it reads to a [`SpanScan`].
struct SpanTap<'a, R> {
    inner: R,
    scan: &'a mut SpanScan,
}

impl<R: std::io::Read> std::io::Read for SpanTap<'_, R> {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        let n = self.inner.read(buf)?;
        self.scan.feed(&buf[..n]);
        Ok(n)
    }
}

/// The rows and columns deleted from each sheet and the sheets deleted, by sheet name.
#[derive(Debug, Default, Clone)]
pub(crate) struct StructureEdits {
    rows: HashMap<String, Bands>,
    cols: HashMap<String, Bands>,
    deleted_sheets: HashSet<String>,
    /// Finds the name of an edited or deleted sheet in a formula, to tell a formula of another
    /// sheet that cannot change from one that can. `None` when a name is not ASCII (its case
    /// would need more than this) or there is none.
    mentions: Option<aho_corasick::AhoCorasick>,
}

impl StructureEdits {
    /// The plan of a command. A scope on a deleted sheet is dropped with its sheet; a scope on a
    /// sheet the workbook does not have is an error.
    pub(crate) fn new(
        sheets: Option<&[String]>,
        deleted_sheets: &HashSet<String>,
        delete_rows: &[ValueScope],
        delete_columns: &[ValueScope],
    ) -> Result<Self> {
        let mut rows: HashMap<String, Vec<(u32, u32)>> = HashMap::new();
        let mut cols: HashMap<String, Vec<(u32, u32)>> = HashMap::new();
        for (scopes, want_rows) in [(delete_rows, true), (delete_columns, false)] {
            for scope in scopes {
                let name = match sheets {
                    Some(sheets) => sheets
                        .iter()
                        .find(|name| **name == scope.sheet)
                        .ok_or_else(|| Error::SheetNotFound {
                            sheet: scope.sheet.clone(),
                            part: "the workbook".to_string(),
                        })?,
                    None => &scope.sheet,
                };
                if deleted_sheets.contains(name) {
                    continue;
                }
                match (scope.target, want_rows) {
                    (ValueTarget::Rows(first, last), true) => {
                        rows.entry(name.clone()).or_default().push((first, last));
                    }
                    (ValueTarget::Columns(first, last), false) => {
                        cols.entry(name.clone()).or_default().push((first, last));
                    }
                    _ if want_rows => {
                        return Err(ScopeError::NotRows {
                            spec: scope.to_string(),
                        }
                        .into());
                    }
                    _ => {
                        return Err(ScopeError::NotColumns {
                            spec: scope.to_string(),
                        }
                        .into());
                    }
                }
            }
        }
        let rows: HashMap<String, Bands> =
            rows.into_iter().map(|(k, v)| (k, Bands::new(v))).collect();
        let cols: HashMap<String, Bands> =
            cols.into_iter().map(|(k, v)| (k, Bands::new(v))).collect();
        let names: Vec<&String> = rows
            .keys()
            .chain(cols.keys())
            .chain(deleted_sheets)
            .collect();
        let mentions = (!names.is_empty() && names.iter().all(|name| name.is_ascii())).then(|| {
            aho_corasick::AhoCorasick::builder()
                .ascii_case_insensitive(true)
                .build(names)
                .expect("sheet names are valid patterns")
        });
        Ok(Self {
            rows,
            cols,
            deleted_sheets: deleted_sheets.clone(),
            mentions,
        })
    }

    /// Whether the command moves or deletes any cell or reference.
    pub(crate) fn is_empty(&self) -> bool {
        self.rows.is_empty() && self.cols.is_empty() && self.deleted_sheets.is_empty()
    }

    /// Whether the plan deletes any row or column.
    pub(crate) fn deletes_rows_or_columns(&self) -> bool {
        !self.rows.is_empty() || !self.cols.is_empty()
    }

    /// Whether a worksheet part that has no row or column deleted can be copied as it is: its
    /// text names no sheet the plan edits or removes and holds no reference across sheets (which
    /// spans sheets it does not name), so no formula in it can change. Only for names whose text
    /// is the same in the XML (no `&`, quotes or angle brackets to escape); otherwise `None`, and
    /// the part is read.
    pub(crate) fn text_names_no_edited_sheet(
        &self,
        text: impl std::io::Read,
    ) -> std::io::Result<Option<bool>> {
        let plain = |name: &String| !name.contains(['&', '\'', '"', '<', '>']);
        if !self
            .rows
            .keys()
            .chain(self.cols.keys())
            .chain(&self.deleted_sheets)
            .all(plain)
        {
            return Ok(None);
        }
        let Some(finder) = &self.mentions else {
            return Ok(None);
        };
        let mut scan = SpanScan::default();
        let named = {
            let mut found = finder
                .try_stream_find_iter(SpanTap {
                    inner: text,
                    scan: &mut scan,
                })
                .map_err(std::io::Error::other)?;
            found.next().transpose()?.is_some()
        };
        Ok(Some(!named && !scan.found))
    }

    /// Whether `formula`, held by a sheet whose layout does not change, can have a reference
    /// the plan changes: only if it names an edited or deleted sheet, or holds a reference across
    /// sheets, which spans sheets it does not name.
    pub(crate) fn may_reach_edited_sheet(&self, formula: &str) -> bool {
        self.mentions.as_ref().is_none_or(|finder| {
            finder.is_match(formula) || {
                let mut scan = SpanScan::default();
                scan.feed(formula.as_bytes());
                scan.found
            }
        })
    }

    /// The deleted rows of `sheet`.
    pub(crate) fn rows_of(&self, sheet: &str) -> Bands {
        self.rows.get(sheet).cloned().unwrap_or_default()
    }

    /// The deleted columns of `sheet`.
    pub(crate) fn cols_of(&self, sheet: &str) -> Bands {
        self.cols.get(sheet).cloned().unwrap_or_default()
    }

    /// The edits to apply to a formula, in the canonical order of the plan: what the command
    /// gives is what these edits give one after the other, whatever the order the operations were
    /// written in and whatever a hash order would have been.
    ///
    /// Within a sheet, the rows deleted come from the last band to the first (a deletion shifts
    /// only what is below it, so the positions of the next one, above, are as they were), then
    /// the columns the same way. The sheets deleted come last. Edits of different sheets do not
    /// depend on each other, because a reference moves only with the rows and columns of the sheet
    /// it names (a reference that spans sheets would tie them, and is refused), so the order
    /// between sheets is free; they go by name, only so that it is always the same.
    pub(crate) fn edits(&self) -> Vec<Edit<'_>> {
        let mut sheets: Vec<&str> = self
            .rows
            .keys()
            .chain(self.cols.keys())
            .map(String::as_str)
            .collect();
        sheets.sort_unstable();
        sheets.dedup();
        let mut deleted: Vec<&str> = self.deleted_sheets.iter().map(String::as_str).collect();
        deleted.sort_unstable();

        let mut edits = Vec::new();
        for &sheet in &sheets {
            if let Some(bands) = self.rows.get(sheet) {
                edits.extend(bands.last_first().map(|(at, count)| Edit::DeleteRows {
                    sheet,
                    at,
                    count,
                }));
            }
        }
        for &sheet in &sheets {
            if let Some(bands) = self.cols.get(sheet) {
                edits.extend(bands.last_first().map(|(at, count)| Edit::DeleteCols {
                    sheet,
                    at,
                    count,
                }));
            }
        }
        edits.extend(deleted.into_iter().map(|sheet| Edit::DeleteSheet { sheet }));
        edits
    }

    /// The text of the formula `source` held by a cell of `home_sheet` once the plan is applied,
    /// `None` when it does not change.
    pub(crate) fn rewrite_formula(
        edits: &[Edit<'_>],
        source: &str,
        home_sheet: &str,
    ) -> Result<Option<String>> {
        Ok(Self::rewrite_formula_tracked(edits, source, home_sheet)?.map(|(text, _)| text))
    }

    /// As [`rewrite_formula`](Self::rewrite_formula), and whether the plan only moved what the
    /// formula reads, so that the result it cached is still its result.
    pub(crate) fn rewrite_formula_tracked(
        edits: &[Edit<'_>],
        source: &str,
        home_sheet: &str,
    ) -> Result<Option<(String, bool)>> {
        let formula = parse_spanned(source).map_err(|e| Error::Formula {
            formula: source.to_string(),
            message: e.to_string(),
        })?;
        if formula.has_three_d() {
            return Err(three_d_unsupported(source));
        }
        Ok(formula.rewrite_all_tracked(edits, home_sheet))
    }

    /// The range `r0..=r1` by `c0..=c1` of `sheet` once the plan is applied, `None` when no cell
    /// of it remains.
    pub(crate) fn range(
        &self,
        sheet: &str,
        r0: u32,
        c0: u32,
        r1: u32,
        c1: u32,
    ) -> Option<(u32, u32, u32, u32)> {
        let (r0, r1) = match self.rows.get(sheet) {
            Some(bands) => bands.span(r0, r1)?,
            None => (r0, r1),
        };
        let (c0, c1) = match self.cols.get(sheet) {
            Some(bands) => bands.span(c0, c1)?,
            None => (c0, c1),
        };
        Some((r0, c0, r1, c1))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn bands_merge_and_move_indexes() {
        // Rows 3-4, 4-6 and 10 (0-based): 3-6 and 10.
        let bands = Bands::new(vec![(10, 10), (3, 4), (4, 6)]);
        assert_eq!(bands.0, [(3, 6), (10, 10)]);
        assert_eq!(bands.kept_in(0, 20), 15);
        assert_eq!(bands.kept_in(5, 6), 3);
        assert_eq!(bands.moved(2), Some(2));
        assert_eq!(bands.moved(3), None);
        assert_eq!(bands.moved(7), Some(3));
        assert_eq!(bands.moved(9), Some(5));
        assert_eq!(bands.moved(10), None);
        assert_eq!(bands.moved(11), Some(6));
        // Touching bands are one.
        assert_eq!(Bands::new(vec![(0, 1), (2, 3)]).0, [(0, 3)]);
    }

    #[test]
    fn a_range_shrinks_around_the_bands() {
        let bands = Bands::new(vec![(3, 4)]);
        assert_eq!(bands.span(1, 2), Some((1, 2)));
        assert_eq!(bands.span(1, 3), Some((1, 2)));
        assert_eq!(bands.span(3, 4), None);
        assert_eq!(bands.span(3, 7), Some((3, 5)));
        assert_eq!(bands.span(1, 7), Some((1, 5)));
        assert_eq!(bands.span(4, 5), Some((3, 3)));
        assert_eq!(bands.span(5, 6), Some((3, 4)));
        assert_eq!(Bands::new(vec![(0, 2)]).span(0, 1), None);
    }
}

#[cfg(test)]
mod span_scan_tests {
    use super::SpanScan;

    fn found(text: &str) -> bool {
        let mut scan = SpanScan::default();
        // Fed in two pieces, as a reader gives it.
        let (a, b) = text.as_bytes().split_at(text.len() / 2);
        scan.feed(a);
        scan.feed(b);
        scan.found
    }

    #[test]
    fn a_span_of_sheets_is_found_however_it_is_written() {
        assert!(found("SUM(Alpha:Gamma!A2:A5)"));
        assert!(found("SUM('Alpha 1':'Gamma 3'!A2)"));
        assert!(found("SUM('Alpha 1:Gamma 3'!A2:A5)"));
        assert!(found("SUM('It''s:Gamma 3'!A2)"));
        assert!(found("SUM([$Alpha.A1:$Gamma.B2])"));
        assert!(!found("SUM('Alpha 1'!A2:A5)+B1:B2"));
        assert!(!found("SUM(Alpha!A2:A5)"));
        assert!(!found("\"a:b\""));
    }
}
