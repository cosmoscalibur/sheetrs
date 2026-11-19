//! What a workbook carries without interpreting it, so that a writer of the same format can
//! put it back.

use std::cmp::Reverse;
use std::collections::{BTreeMap, BinaryHeap, HashMap};

use crate::coord::CellPos;
use crate::workbook::DateSystem;

/// A file format whose package a workbook can carry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PackageFormat {
    /// An XLSX package.
    Xlsx,
    /// An ODS package.
    Ods,
}

/// An entry of a package that the model does not interpret.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaquePart {
    /// The path of the entry in the package.
    pub name: String,
    /// The bytes of the entry, as they were read.
    pub bytes: Vec<u8>,
}

/// The bytes of an XML part before and after the element that the model replaces.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Frame {
    /// Everything up to the start of the replaced element.
    pub head: Vec<u8>,
    /// Everything after the end of the replaced element.
    pub tail: Vec<u8>,
    /// Whether the name of the replaced element has a namespace prefix, which a writer that
    /// builds the element cannot reproduce.
    pub prefixed: bool,
}

/// The part of a package that holds the cells, as it was read, with what the model held of it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Pristine {
    /// The part, compressed as the package had it, alone in a ZIP archive.
    pub entry: Vec<u8>,
    /// What [`crate::Workbook::fingerprint`] was when the part was read. A writer that finds
    /// the same fingerprint writes the part as it was.
    pub fingerprint: u64,
}

/// The package a workbook was read from, apart from what the model holds.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaquePackage {
    /// The format of the package.
    pub format: PackageFormat,
    /// The entries the writer copies as they are.
    pub parts: Vec<OpaquePart>,
    /// The main part of the workbook around its list of sheets.
    pub frame: Frame,
    /// How the package counted dates, which `frame` repeats.
    pub date_system: DateSystem,
    /// The defined names the model does not hold, as the package wrote them.
    pub defined_names: Vec<String>,
    /// The part with the cells, when the writer can put it back as it was while the workbook is
    /// as it was read.
    pub pristine: Option<Pristine>,
    /// Where the shared strings are; text is written inline when the package had none.
    pub shared_strings: Option<String>,
}

/// What a sheet carries from its package, apart from its cells.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SheetOpaque {
    /// The part of the sheet around its cells.
    pub frame: Frame,
    /// The path of the part of the sheet in the package.
    pub part: String,
    /// The attributes of the entry of the sheet in the workbook part, except its name and state.
    pub entry: String,
}

/// Rows or cells that share an id, as runs of consecutive rows.
#[derive(Debug, Clone, Default)]
pub(crate) struct Runs {
    /// First row to (last row, id).
    runs: BTreeMap<u32, (u32, u32)>,
}

impl Runs {
    /// Gives rows `first..=last` the id `id`; rows that already have one keep it.
    pub(crate) fn set(&mut self, first: u32, last: u32, id: u32) {
        if first > last {
            return;
        }
        let before = self
            .runs
            .range(..=first)
            .next_back()
            .map(|(start, (end, found))| (*start, *end, *found));
        if let Some((_, end, _)) = before.filter(|(_, end, _)| *end >= first) {
            if end < last {
                self.set(end + 1, last, id);
            }
            return;
        }
        let after = self
            .runs
            .range(first..)
            .next()
            .map(|(start, (end, _))| (*start, *end))
            .filter(|(start, _)| *start <= last);
        if let Some((start, end)) = after {
            self.set(first, start - 1, id);
            if end < last {
                self.set(end + 1, last, id);
            }
            return;
        }
        match before {
            Some((start, end, found)) if end + 1 == first && found == id => {
                self.runs.insert(start, (last, id));
            }
            _ => {
                self.runs.insert(first, (last, id));
            }
        }
    }

    pub(crate) fn iter(&self) -> impl Iterator<Item = (u32, u32, u32)> + '_ {
        self.runs
            .iter()
            .map(|(first, (last, id))| (*first, *last, *id))
    }

    pub(crate) fn is_empty(&self) -> bool {
        self.runs.is_empty()
    }
}

/// The opaque state of the rows and cells of a sheet.
#[derive(Debug, Clone, Default)]
pub(crate) struct Appearance {
    pub(crate) opaque: Option<SheetOpaque>,
    /// The attribute text of the rows, by run.
    pub(crate) rows: Runs,
    pub(crate) row_texts: Vec<String>,
    row_ids: HashMap<String, u32>,
    /// The style id of the cells, by column.
    pub(crate) styles: BTreeMap<u32, Runs>,
    /// The cells whose text has formatting, with the formatted text and the text it has
    /// without formatting.
    pub(crate) rich: HashMap<CellPos, usize>,
    pub(crate) rich_texts: Vec<(String, String)>,
    rich_ids: HashMap<(String, String), usize>,
}

impl Appearance {
    pub(crate) fn set_row(&mut self, row: u32, attributes: &str) {
        let id = self.row_id(attributes);
        self.rows.set(row, row, id);
    }

    fn row_id(&mut self, attributes: &str) -> u32 {
        if let Some(id) = self.row_ids.get(attributes) {
            return *id;
        }
        let id = u32::try_from(self.row_texts.len()).unwrap_or(u32::MAX);
        self.row_ids.insert(attributes.to_owned(), id);
        self.row_texts.push(attributes.to_owned());
        id
    }

    pub(crate) fn set_rich(&mut self, pos: CellPos, plain: &str, formatted: &str) {
        let key = (plain.to_owned(), formatted.to_owned());
        let known = self.rich_ids.get(&key).copied();
        let id = known.unwrap_or_else(|| {
            let id = self.rich_texts.len();
            self.rich_ids.insert(key.clone(), id);
            self.rich_texts.push(key);
            id
        });
        self.rich.insert(pos, id);
    }

    pub(crate) fn set_style(&mut self, pos: CellPos, style: u32) {
        self.styles
            .entry(pos.col)
            .or_default()
            .set(pos.row, pos.row, style);
    }

    /// Brings in the rows and cell styles of `other`, whose rows and cells this sheet does not
    /// hold.
    pub(crate) fn absorb(&mut self, other: &Appearance) {
        for (first, last, id) in other.rows.iter() {
            let id = self.row_id(&other.row_texts[id as usize]);
            self.rows.set(first, last, id);
        }
        for (pos, id) in &other.rich {
            let (plain, formatted) = &other.rich_texts[*id];
            self.set_rich(*pos, plain, formatted);
        }
        for (col, runs) in &other.styles {
            let mine = self.styles.entry(*col).or_default();
            for (first, last, id) in runs.iter() {
                mine.set(first, last, id);
            }
        }
    }

    /// The rows and columns the styled cells span: first row, last row, first column, last
    /// column.
    pub(crate) fn styled_extent(&self) -> Option<(u32, u32, u32, u32)> {
        let mut extent: Option<(u32, u32, u32, u32)> = None;
        for (col, runs) in &self.styles {
            let first = runs.iter().next().map(|(first, _, _)| first);
            let last = runs.iter().last().map(|(_, last, _)| last);
            if let (Some(first), Some(last)) = (first, last) {
                extent = Some(match extent {
                    None => (first, last, *col, *col),
                    Some((a, b, c, d)) => (a.min(first), b.max(last), c.min(*col), d.max(*col)),
                });
            }
        }
        extent
    }

    pub(crate) fn styled_cells(&self) -> StyledCells {
        let columns: Vec<ColumnRuns> = self
            .styles
            .iter()
            .filter(|(_, runs)| !runs.is_empty())
            .map(|(col, runs)| (*col, runs.iter().collect()))
            .collect();
        let mut heap = BinaryHeap::new();
        for (index, (_, runs)) in columns.iter().enumerate() {
            if let Some((first, _, _)) = runs.first() {
                heap.push(Reverse((*first, index)));
            }
        }
        StyledCells {
            cursors: vec![0; columns.len()],
            columns,
            heap,
        }
    }
}

/// A column and its runs of rows, each as first row, last row and id.
type ColumnRuns = (u32, Vec<(u32, u32, u32)>);

/// The cells that have a style, row by row, with the style of each.
#[derive(Debug)]
pub struct StyledCells {
    /// Each column with its runs of rows.
    columns: Vec<ColumnRuns>,
    /// For each column, the run it is in.
    cursors: Vec<usize>,
    heap: BinaryHeap<Reverse<(u32, usize)>>,
}

impl Iterator for StyledCells {
    type Item = (CellPos, u32);

    fn next(&mut self) -> Option<Self::Item> {
        let Reverse((row, index)) = self.heap.pop()?;
        let (col, runs) = &self.columns[index];
        let run = &mut self.cursors[index];
        let (_, last, id) = runs[*run];
        if row < last {
            self.heap.push(Reverse((row + 1, index)));
        } else {
            *run += 1;
            if let Some((first, _, _)) = runs.get(*run) {
                self.heap.push(Reverse((*first, index)));
            }
        }
        Some((CellPos { row, col: *col }, id))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn runs_extend_and_keep_the_first_id() {
        let mut runs = Runs::default();
        runs.set(0, 0, 1);
        runs.set(1, 1, 1);
        runs.set(2, 2, 2);
        runs.set(1, 4, 3);
        assert_eq!(
            runs.iter().collect::<Vec<_>>(),
            vec![(0, 1, 1), (2, 2, 2), (3, 4, 3)]
        );
    }

    #[test]
    fn runs_fill_the_gap_before_a_later_run() {
        let mut runs = Runs::default();
        runs.set(5, 6, 7);
        runs.set(2, 9, 8);
        assert_eq!(
            runs.iter().collect::<Vec<_>>(),
            vec![(2, 4, 8), (5, 6, 7), (7, 9, 8)]
        );
    }

    #[test]
    fn styled_cells_come_row_by_row() {
        let mut appearance = Appearance::default();
        for (row, col, style) in [(0, 1, 5), (1, 1, 5), (0, 0, 2), (1, 0, 3), (3, 0, 3)] {
            appearance.set_style(CellPos { row, col }, style);
        }
        let found: Vec<_> = appearance
            .styled_cells()
            .map(|(pos, style)| (pos.row, pos.col, style))
            .collect();
        assert_eq!(
            found,
            vec![(0, 0, 2), (0, 1, 5), (1, 0, 3), (1, 1, 5), (3, 0, 3)]
        );
    }

    #[test]
    fn rows_with_the_same_attributes_share_a_text() {
        let mut appearance = Appearance::default();
        appearance.set_row(0, "ht=\"15\"");
        appearance.set_row(1, "ht=\"15\"");
        appearance.set_row(2, "hidden=\"1\"");
        assert_eq!(appearance.row_texts.len(), 2);
        assert_eq!(appearance.rows.iter().count(), 2);
    }
}
