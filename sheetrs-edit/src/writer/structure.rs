//! What a command does to the layout of a workbook: rows and columns deleted from sheets, sheets
//! removed. Everything is expressed in the coordinates of the input file, so the user needs no
//! idea of the positions the cells will have afterwards.
//!
//! References follow LibreOffice 26.8, which `sheetrs_formula::edit` implements: a reference to
//! deleted cells is `#REF!`, a range shrinks around the deletion, whole rows and columns do not
//! move when the other axis changes.

use super::{ValueScope, ValueTarget};
use anyhow::{Context, Result, bail};
use sheetrs_formula::edit::{Edit, Workbook};
use sheetrs_formula::parse::parse_spanned;
use std::collections::{HashMap, HashSet};

/// Disjoint, ordered, inclusive 0-based bands of deleted rows (or columns), and how many indexes
/// the bands before each one delete, so a position is found by search and not by a walk.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) struct Bands(Vec<(u32, u32)>, Vec<u32>);

impl Bands {
    /// The bands of `ranges`, which may overlap or touch.
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

    pub fn is_empty(&self) -> bool {
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
    pub fn kept_in(&self, first: u32, count: u32) -> u32 {
        let end = first.saturating_add(count);
        count - (self.deleted_before(end) - self.deleted_before(first))
    }

    /// Whether `index` is deleted.
    pub fn contains(&self, index: u32) -> bool {
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

    /// Where an anchor at `index` is once the bands are gone: its place, or its own index when it
    /// is deleted. LibreOffice keeps the position of a shape whose cell goes and anchors it to the
    /// cell that is there now, which is the cell of that index when the rows or columns that pass
    /// have the same height or width; the size of those is not in the file.
    pub fn anchored_at(&self, index: u32) -> u32 {
        self.moved(index).unwrap_or(index)
    }

    /// The index that the first row or column at or after `index` that stays has once the bands
    /// are gone.
    pub fn shifted(&self, index: u32) -> u32 {
        index - self.deleted_before(index)
    }

    /// Whether `index` is deleted together with the index before it, which a shape anchored
    /// there has in a band that starts above it.
    pub fn deleted_below_a_deleted(&self, index: u32) -> bool {
        index > 0 && self.contains(index) && self.contains(index - 1)
    }

    /// Where `index` is once the bands are gone, `None` when it is deleted.
    pub fn moved(&self, index: u32) -> Option<u32> {
        (!self.contains(index)).then(|| index - self.deleted_before(index))
    }

    /// The bounds of `first..=last` once the bands are gone, `None` when nothing of it remains.
    pub fn span(&self, first: u32, last: u32) -> Option<(u32, u32)> {
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

    /// The bands as `(first, count)`, the last one first.
    pub fn last_first(&self) -> impl Iterator<Item = (u32, u32)> + '_ {
        self.0
            .iter()
            .rev()
            .map(|&(first, last)| (first, last - first + 1))
    }
}

/// The error of a deletion that removes the rows (or columns) above the cell a shape is anchored
/// to together with that cell. LibreOffice places the shape then by its size and the size of the
/// rows (or columns), which are not in the file.
pub(super) fn shape_depends_on_sizes(sheet: &str) -> anyhow::Error {
    anyhow::anyhow!(
        "a deletion in sheet '{sheet}' removes a cell that a shape is anchored to together with the rows or columns before it: LibreOffice places the shape by its size and the height of the rows (or width of the columns), which the file does not hold"
    )
}

/// Reads the names of the sheets of the document at a path, in order.
pub(crate) type SheetReader = fn(&std::path::Path) -> Result<Vec<String>>;

/// Finds, in text read in pieces, a reference across sheets: a `:` followed by a sheet name and
/// `!` (`Sheet1:Sheet3!A1`, an XLSX) or `.` (`$S.A1:$T.B2`, an ODS). It also matches texts that are
/// not one (`A1:A5.5`): the cost of that is a formula read that did not need it.
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
}

impl SpanScan {
    fn feed(&mut self, bytes: &[u8]) {
        use ScanState::{AfterColon, Idle, Name, QuoteEnd, Quoted};
        let name_byte = |b: u8| b.is_ascii_alphanumeric() || b == b'_' || b >= 0x80;
        for &b in bytes {
            self.state = match (self.state, b) {
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

/// The rows and columns deleted from each sheet and the sheets removed, by sheet name.
#[derive(Debug, Default, Clone)]
pub(crate) struct StructureEdits {
    rows: HashMap<String, Bands>,
    cols: HashMap<String, Bands>,
    removed_sheets: HashSet<String>,
    /// The sheets of the workbook in order, which a 3D reference spans, or the error of reading
    /// them. Read when the first 3D reference needs them if the document does not give them first.
    order: std::cell::OnceCell<Result<Vec<String>, String>>,
    /// Where to read the sheets from when a 3D reference needs them.
    order_source: Option<(std::path::PathBuf, SheetReader)>,
    /// Finds the name of an edited or removed sheet in a formula, to tell a formula of another
    /// sheet that cannot change from one that can. `None` when a name is not ASCII (its case
    /// would need more than this) or there is none.
    mentions: Option<aho_corasick::AhoCorasick>,
}

impl StructureEdits {
    /// The plan of a command. A scope on a removed sheet is dropped with its sheet; a scope on a
    /// sheet the workbook does not have is an error.
    pub fn new(
        sheets: Option<&[String]>,
        removed_sheets: &HashSet<String>,
        delete_rows: &[ValueScope],
        delete_cols: &[ValueScope],
    ) -> Result<Self> {
        let mut rows: HashMap<String, Vec<(u32, u32)>> = HashMap::new();
        let mut cols: HashMap<String, Vec<(u32, u32)>> = HashMap::new();
        for (scopes, want_rows) in [(delete_rows, true), (delete_cols, false)] {
            for scope in scopes {
                let name = match sheets {
                    Some(sheets) => sheets
                        .iter()
                        .find(|name| **name == scope.sheet)
                        .with_context(|| format!("Sheet '{}' not found", scope.sheet))?,
                    None => &scope.sheet,
                };
                if removed_sheets.contains(name) {
                    continue;
                }
                match (scope.target, want_rows) {
                    (ValueTarget::Rows(first, last), true) => {
                        rows.entry(name.clone()).or_default().push((first, last));
                    }
                    (ValueTarget::Columns(first, last), false) => {
                        cols.entry(name.clone()).or_default().push((first, last));
                    }
                    _ if want_rows => bail!("'{scope}' is not a row range such as 'Sheet!3:10'"),
                    _ => bail!("'{scope}' is not a column range such as 'Sheet!B:D'"),
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
            .chain(removed_sheets)
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
            removed_sheets: removed_sheets.clone(),
            order: sheets
                .map(|sheets| std::cell::OnceCell::from(Ok(sheets.to_vec())))
                .unwrap_or_default(),
            order_source: None,
            mentions,
        })
    }

    /// The plan with the sheets of the workbook in order.
    #[cfg(test)]
    pub fn with_order(mut self, order: Vec<String>) -> Self {
        self.order = std::cell::OnceCell::from(Ok(order));
        self
    }

    /// The plan reading the sheets of the document at `path` with `read` when a 3D reference
    /// needs them: a document that holds them after the formulas that name them.
    pub fn with_order_source(mut self, path: std::path::PathBuf, read: SheetReader) -> Self {
        self.order_source = Some((path, read));
        self
    }

    /// The sheets of the workbook in order; none when they cannot be read (see `order_error`).
    fn order(&self) -> &[String] {
        let order = self.order.get_or_init(|| match &self.order_source {
            Some((path, read)) => read(path).map_err(|e| format!("{e:#}")),
            None => Ok(Vec::new()),
        });
        order.as_deref().unwrap_or_default()
    }

    /// Why the sheets could not be read, when a 3D reference needed them.
    pub fn order_error(&self) -> Option<&str> {
        self.order.get()?.as_ref().err().map(String::as_str)
    }

    /// Whether the command moves or deletes any cell or reference.
    pub fn is_empty(&self) -> bool {
        self.rows.is_empty() && self.cols.is_empty() && self.removed_sheets.is_empty()
    }

    /// Whether the plan deletes any row or column.
    pub fn deletes_rows_or_columns(&self) -> bool {
        !self.rows.is_empty() || !self.cols.is_empty()
    }

    /// Whether a worksheet part that has no row or column deleted can be copied as it is: its
    /// text names no sheet the plan edits or removes and holds no reference across sheets (which
    /// spans sheets it does not name), so no formula in it can change. Only for names whose text
    /// is the same in the XML (no `&`, quotes or angle brackets to escape); otherwise `None`, and
    /// the part is read.
    pub fn text_names_no_edited_sheet(
        &self,
        text: impl std::io::Read,
    ) -> std::io::Result<Option<bool>> {
        let plain = |name: &String| !name.contains(['&', '\'', '"', '<', '>']);
        if !self
            .rows
            .keys()
            .chain(self.cols.keys())
            .chain(&self.removed_sheets)
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
    /// the plan changes: only if it names an edited or removed sheet, or holds a reference across
    /// sheets, which spans sheets it does not name.
    pub fn may_reach_edited_sheet(&self, formula: &str) -> bool {
        self.mentions.as_ref().is_none_or(|finder| {
            finder.is_match(formula) || {
                let mut scan = SpanScan::default();
                scan.feed(formula.as_bytes());
                scan.found
            }
        })
    }

    /// The deleted rows of `sheet`.
    pub fn rows_of(&self, sheet: &str) -> Bands {
        self.rows.get(sheet).cloned().unwrap_or_default()
    }

    /// The deleted columns of `sheet`.
    pub fn cols_of(&self, sheet: &str) -> Bands {
        self.cols.get(sheet).cloned().unwrap_or_default()
    }

    /// The edits to apply to a formula, in the order that leaves the positions of the next one as
    /// they were.
    pub fn edits(&self) -> Vec<Edit<'_>> {
        let mut edits = Vec::new();
        for sheet in &self.removed_sheets {
            edits.push(Edit::DeleteSheet { sheet });
        }
        for (sheet, bands) in &self.rows {
            edits.extend(bands.last_first().map(|(at, count)| Edit::DeleteRows {
                sheet,
                at,
                count,
            }));
        }
        for (sheet, bands) in &self.cols {
            edits.extend(bands.last_first().map(|(at, count)| Edit::DeleteCols {
                sheet,
                at,
                count,
            }));
        }
        edits
    }

    /// The text of the formula `source` held by a cell of `home_sheet` once the plan is applied,
    /// `None` when it does not change.
    pub fn rewrite_formula(
        &self,
        edits: &[Edit<'_>],
        source: &str,
        home_sheet: &str,
    ) -> Result<Option<String>> {
        let formula =
            parse_spanned(source).map_err(|e| anyhow::anyhow!("cannot read '{source}': {e}"))?;
        let removed = |name: &str| self.removed_sheets.contains(name);
        let book = Workbook {
            order: self.order(),
            removed: &removed,
        };
        Ok(formula.rewrite_all(edits, home_sheet, Some(&book)))
    }

    /// Whether the command removes the sheet.
    pub fn removes_sheet(&self, name: &str) -> bool {
        self.removed_sheets.contains(name)
    }

    /// The sheets of the workbook in order, for 3D references; `removed` tells the removed ones.
    pub fn workbook<'a>(&'a self, removed: &'a dyn Fn(&str) -> bool) -> Workbook<'a> {
        Workbook {
            order: self.order(),
            removed,
        }
    }

    /// The range `r0..=r1` by `c0..=c1` of `sheet` once the plan is applied, `None` when no cell
    /// of it remains.
    pub fn range(
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
