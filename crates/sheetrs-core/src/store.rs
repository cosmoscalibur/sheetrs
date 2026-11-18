//! Column-wise storage of the cells of a sheet, and the tables that intern the strings and
//! the formulas of a workbook.
//!
//! A column keeps its cells in parallel channels: the sorted rows, a tag with the kind of the
//! value and the `stale` bit, an eight-byte payload and a formula slot. Text, errors outside
//! the classic ones and formula texts are numbers that index the tables of the workbook.
//!
//! Numbers issued by the tables are produced here and read back here, so a number that does
//! not resolve is a bug and stops the program. What the input controls (how many strings,
//! how much formula text) is checked when storing and reported as a [`LimitError`].

use std::cmp::Reverse;
use std::collections::{BTreeMap, BinaryHeap};
use std::sync::Arc;

use rustc_hash::FxHashMap;
use thiserror::Error;

use crate::coord::CellPos;
use crate::workbook::{Cell, ErrorKind, Formula, Value, Visibility, Workbook};

const KIND_MASK: u8 = 0b0111;
/// Marks a cell whose formula slot is a shared-formula group number, not a text slot.
const SHARED_BIT: u8 = 0b0100_0000;
const STALE_BIT: u8 = 0b1000_0000;
const EMPTY: u8 = 0;
const NUMBER: u8 = 1;
const BOOL: u8 = 2;
const TEXT: u8 = 3;
const ERROR: u8 = 4;

/// Error payloads from this value on are `ErrorKind::Other` texts: the rest is the text number.
const OTHER_ERROR_BASE: u64 = 16;

/// A capacity of the cell store that the input exceeded.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
pub enum LimitError {
    /// More distinct strings than numbers of 32 bits can name.
    #[error("more than 4294967295 distinct strings")]
    Strings,
    /// More formula text than offsets of 32 bits can address.
    #[error("more than 4 GiB of formula text")]
    FormulaText,
    /// More formulas than numbers of 32 bits can name.
    #[error("more than 4294967295 formulas")]
    Formulas,
}

/// The table index a stored number stands for.
fn index_of(number: u32) -> usize {
    usize::try_from(number).expect("usize holds 32 bits")
}

fn string_id(count: usize) -> Result<u32, LimitError> {
    u32::try_from(count).map_err(|_| LimitError::Strings)
}

fn formula_end(stored: usize, added: usize) -> Result<u32, LimitError> {
    stored
        .checked_add(added)
        .and_then(|end| u32::try_from(end).ok())
        .ok_or(LimitError::FormulaText)
}

fn formula_slot(count: usize) -> Result<u32, LimitError> {
    count
        .checked_add(1)
        .and_then(|slot| u32::try_from(slot).ok())
        .ok_or(LimitError::Formulas)
}

/// Strings stored once and referred to by number.
#[derive(Debug, Clone, Default)]
pub(crate) struct Interner {
    strings: Vec<Arc<str>>,
    ids: FxHashMap<Arc<str>, u32>,
}

impl Interner {
    fn intern(&mut self, text: &str) -> Result<u32, LimitError> {
        if let Some(id) = self.ids.get(text) {
            return Ok(*id);
        }
        let id = string_id(self.strings.len())?;
        let text: Arc<str> = Arc::from(text);
        self.strings.push(Arc::clone(&text));
        self.ids.insert(text, id);
        Ok(id)
    }

    /// Interns every string of `other`; the result maps the numbers of `other` to these.
    fn merge(&mut self, other: &Interner) -> Result<Vec<u32>, LimitError> {
        other.strings.iter().map(|text| self.intern(text)).collect()
    }

    fn get(&self, id: u32) -> &str {
        self.strings
            .get(index_of(id))
            .expect("string numbers come from this table")
    }
}

/// Formula texts stored back to back; a slot is the 1-based position of a text.
///
/// Texts are not deduplicated: a hash table over distinct formula texts costs more than the
/// repeated texts it would save.
#[derive(Debug, Clone, Default)]
pub(crate) struct Arena {
    text: String,
    ends: Vec<u32>,
}

impl Arena {
    fn push(&mut self, text: &str) -> Result<u32, LimitError> {
        let end = formula_end(self.text.len(), text.len())?;
        let slot = formula_slot(self.ends.len())?;
        self.text.push_str(text);
        self.ends.push(end);
        Ok(slot)
    }

    /// Appends every text of `other`; a slot of `other` reads as itself plus the result.
    fn merge(&mut self, other: &Arena) -> Result<u32, LimitError> {
        let end = formula_end(self.text.len(), other.text.len())?;
        let count = self.ends.len() + other.ends.len();
        u32::try_from(count).map_err(|_| LimitError::Formulas)?;
        let slot_base = u32::try_from(self.ends.len()).expect("checked with the total");
        let text_base = end - u32::try_from(other.text.len()).expect("checked with the total");
        self.text.push_str(&other.text);
        self.ends
            .extend(other.ends.iter().map(|end| end + text_base));
        Ok(slot_base)
    }

    fn get(&self, slot: u32) -> &str {
        let index = index_of(slot)
            .checked_sub(1)
            .expect("formula slots start at 1");
        let start = index
            .checked_sub(1)
            .map_or(0, |previous| index_of(self.ends[previous]));
        let end = index_of(*self.ends.get(index).expect("slots come from this table"));
        self.text
            .get(start..end)
            .expect("formula texts are stored whole")
    }
}

/// The string and formula tables of a workbook.
#[derive(Debug, Clone, Default)]
pub(crate) struct Tables {
    strings: Interner,
    formulas: Arena,
}

const KNOWN_ERRORS: [ErrorKind; 7] = [
    ErrorKind::Null,
    ErrorKind::DivZero,
    ErrorKind::Value,
    ErrorKind::Ref,
    ErrorKind::Name,
    ErrorKind::Num,
    ErrorKind::Na,
];

/// The number a stored payload holds when the kind of the cell refers to a table.
fn table_number(payload: u64) -> u32 {
    u32::try_from(payload).expect("table numbers fit 32 bits")
}

impl Tables {
    fn encode_value(&mut self, value: &ValueRef<'_>) -> Result<(u8, u64), LimitError> {
        Ok(match value {
            ValueRef::Empty => (EMPTY, 0),
            ValueRef::Number(number) => (NUMBER, number.to_bits()),
            ValueRef::Bool(flag) => (BOOL, u64::from(*flag)),
            ValueRef::Text(text) => (TEXT, u64::from(self.strings.intern(text)?)),
            ValueRef::Error(kind) => {
                let payload = match KNOWN_ERRORS.iter().position(|known| known == kind) {
                    Some(index) => u64::try_from(index).expect("seven classic errors"),
                    None => OTHER_ERROR_BASE + u64::from(self.strings.intern(kind.as_text())?),
                };
                (ERROR, payload)
            }
        })
    }

    /// The formula slot and the tag bits that say how to read it.
    fn encode_formula(&mut self, formula: Option<FormulaRef<'_>>) -> Result<(u32, u8), LimitError> {
        Ok(match formula {
            None => (0, 0),
            Some(FormulaRef::Text(text)) => (self.formulas.push(text)?, 0),
            Some(FormulaRef::Shared { index }) => (index, SHARED_BIT),
        })
    }

    fn decode_value(&self, kind: u8, payload: u64) -> ValueRef<'_> {
        match kind {
            NUMBER => ValueRef::Number(f64::from_bits(payload)),
            BOOL => ValueRef::Bool(payload != 0),
            TEXT => ValueRef::Text(self.strings.get(table_number(payload))),
            ERROR => ValueRef::Error(match payload.checked_sub(OTHER_ERROR_BASE) {
                Some(id) => ErrorKind::Other(self.strings.get(table_number(id)).into()),
                None => KNOWN_ERRORS
                    .get(usize::try_from(payload).expect("classic error codes are small"))
                    .expect("classic error codes are issued by `encode_value`")
                    .clone(),
            }),
            _ => ValueRef::Empty,
        }
    }

    fn decode_formula(&self, tag: u8, slot: u32) -> Option<FormulaRef<'_>> {
        if tag & SHARED_BIT != 0 {
            Some(FormulaRef::Shared { index: slot })
        } else if slot == 0 {
            None
        } else {
            Some(FormulaRef::Text(self.formulas.get(slot)))
        }
    }

    /// Brings the strings and formula texts of `other` into these tables.
    fn merge(&mut self, other: &Tables) -> Result<Remap, LimitError> {
        Ok(Remap {
            strings: self.strings.merge(&other.strings)?,
            formula_base: self.formulas.merge(&other.formulas)?,
        })
    }
}

/// How the numbers stored by the cells of one workbook read in another, after a merge.
struct Remap {
    strings: Vec<u32>,
    formula_base: u32,
}

impl Remap {
    /// The payload and formula slot of a cell of the merged workbook as they read in the other.
    fn cell(&self, tag: u8, payload: u64, slot: u32) -> (u64, u32) {
        let string = |number: u64| {
            let merged = self
                .strings
                .get(index_of(table_number(number)))
                .expect("string numbers of the other workbook are in its table");
            u64::from(*merged)
        };
        let payload = match tag & KIND_MASK {
            TEXT => string(payload),
            ERROR => match payload.checked_sub(OTHER_ERROR_BASE) {
                Some(id) => OTHER_ERROR_BASE + string(id),
                None => payload,
            },
            _ => payload,
        };
        let slot = if tag & SHARED_BIT != 0 || slot == 0 {
            slot
        } else {
            slot.checked_add(self.formula_base)
                .expect("merged slots fit the total checked by the merge")
        };
        (payload, slot)
    }
}

/// The encoded parts of a cell: tag, payload and formula slot.
type Encoded = (u8, u64, u32);

/// The cells of one column, as parallel channels sorted by row.
#[derive(Debug, Clone, Default)]
struct Column {
    rows: Vec<u32>,
    tags: Vec<u8>,
    payloads: Vec<u64>,
    formulas: Vec<u32>,
    /// Cells that arrived above the last row of the channels, kept ordered apart so that a
    /// sheet written in any order costs no more than a sorted insert per cell; `seal` merges
    /// them into the channels.
    pending: BTreeMap<u32, Encoded>,
}

impl Column {
    /// Where `row` is among the channels, or would go; cells arrive mostly below the last
    /// one, which is checked first.
    fn position(&self, row: u32) -> Result<usize, usize> {
        match self.rows.last() {
            None => Err(0),
            Some(last) if *last < row => Err(self.rows.len()),
            Some(_) => self.rows.binary_search(&row),
        }
    }

    fn push(&mut self, row: u32, (tag, payload, formula): Encoded) {
        self.rows.push(row);
        self.tags.push(tag);
        self.payloads.push(payload);
        self.formulas.push(formula);
    }

    fn reserve(&mut self, additional: usize) {
        self.rows.reserve(additional);
        self.tags.reserve(additional);
        self.payloads.reserve(additional);
        self.formulas.reserve(additional);
    }

    /// The cell at `row`, wherever it is held.
    fn entry(&self, row: u32) -> Option<Encoded> {
        match self.position(row) {
            Ok(index) => Some((self.tags[index], self.payloads[index], self.formulas[index])),
            Err(_) => self.pending.get(&row).copied(),
        }
    }

    /// Stores a cell, replacing the one at `row`; returns whether the cell is new.
    fn set(&mut self, row: u32, cell: Encoded) -> bool {
        match self.position(row) {
            Ok(index) => {
                (self.tags[index], self.payloads[index], self.formulas[index]) = cell;
                false
            }
            Err(index) if index == self.rows.len() => {
                self.push(row, cell);
                true
            }
            Err(_) => self.pending.insert(row, cell).is_none(),
        }
    }

    /// Merges the pending cells into the channels and releases the spare capacity.
    fn seal(&mut self) {
        if self.pending.is_empty() {
            self.rows.shrink_to_fit();
            self.tags.shrink_to_fit();
            self.payloads.shrink_to_fit();
            self.formulas.shrink_to_fit();
            return;
        }
        let total = self.rows.len() + self.pending.len();
        let mut merged = Column {
            rows: Vec::with_capacity(total),
            tags: Vec::with_capacity(total),
            payloads: Vec::with_capacity(total),
            formulas: Vec::with_capacity(total),
            pending: BTreeMap::new(),
        };
        let mut pending = std::mem::take(&mut self.pending).into_iter().peekable();
        for index in 0..self.rows.len() {
            let row = self.rows[index];
            while let Some((pending_row, cell)) =
                pending.next_if(|(pending_row, _)| *pending_row < row)
            {
                merged.push(pending_row, cell);
            }
            merged.push(
                row,
                (self.tags[index], self.payloads[index], self.formulas[index]),
            );
        }
        for (pending_row, cell) in pending {
            merged.push(pending_row, cell);
        }
        *self = merged;
    }
}

/// A sheet: its name, visibility and cells by column.
#[derive(Debug, Clone)]
pub(crate) struct Sheet {
    name: String,
    visibility: Visibility,
    /// Columns that hold cells, sorted by column number.
    columns: Vec<(u32, Column)>,
    shared_anchors: BTreeMap<u32, CellPos>,
    len: usize,
}

impl Sheet {
    pub(crate) fn name(&self) -> &str {
        &self.name
    }

    fn column(&self, col: u32) -> Option<&Column> {
        let index = self
            .columns
            .binary_search_by_key(&col, |(number, _)| *number)
            .ok()?;
        Some(&self.columns[index].1)
    }

    fn column_mut(&mut self, col: u32) -> &mut Column {
        let index = match self
            .columns
            .binary_search_by_key(&col, |(number, _)| *number)
        {
            Ok(index) => index,
            Err(index) => {
                self.columns.insert(index, (col, Column::default()));
                index
            }
        };
        &mut self.columns[index].1
    }

    fn set_encoded(&mut self, pos: CellPos, cell: Encoded) {
        if self.column_mut(pos.col).set(pos.row, cell) {
            self.len += 1;
        }
    }

    /// The cell of a finished sheet at `pos`, as its column and index in the channels.
    fn find(&self, pos: CellPos) -> Option<(&Column, usize)> {
        let column = self.column(pos.col)?;
        let row = column.position(pos.row).ok()?;
        Some((column, row))
    }
}

/// Adds a sheet to a workbook cell by cell.
#[derive(Debug)]
pub struct SheetBuilder<'a> {
    workbook: &'a mut Workbook,
    sheet: Sheet,
}

impl<'a> SheetBuilder<'a> {
    pub(crate) fn new(workbook: &'a mut Workbook, name: String) -> Self {
        Self {
            workbook,
            sheet: Sheet {
                name,
                visibility: Visibility::Visible,
                columns: Vec::new(),
                shared_anchors: BTreeMap::new(),
                len: 0,
            },
        }
    }

    /// Sets whether the sheet is shown.
    pub fn set_visibility(&mut self, visibility: Visibility) {
        self.sheet.visibility = visibility;
    }

    /// Stores `cell` at `pos`, replacing any previous cell.
    ///
    /// # Errors
    ///
    /// Returns a [`LimitError`] when the workbook would hold more strings or formula text
    /// than the store can address.
    pub fn set(&mut self, pos: CellPos, cell: &Cell) -> Result<(), LimitError> {
        let value = match &cell.value {
            Value::Empty => ValueRef::Empty,
            Value::Number(number) => ValueRef::Number(*number),
            Value::Bool(flag) => ValueRef::Bool(*flag),
            Value::Text(text) => ValueRef::Text(text),
            Value::Error(kind) => ValueRef::Error(kind.clone()),
        };
        let formula = cell.formula.as_ref().map(|formula| match formula {
            Formula::Text(text) => FormulaRef::Text(text),
            Formula::Shared { index } => FormulaRef::Shared { index: *index },
        });
        self.set_view(pos, &value, formula, cell.stale)
    }

    /// Stores a cell given by borrowed parts at `pos`, replacing any previous cell.
    ///
    /// # Errors
    ///
    /// Returns a [`LimitError`] when the workbook would hold more strings or formula text
    /// than the store can address.
    pub fn set_view(
        &mut self,
        pos: CellPos,
        value: &ValueRef<'_>,
        formula: Option<FormulaRef<'_>>,
        stale: bool,
    ) -> Result<(), LimitError> {
        let tables = &mut self.workbook.tables;
        let (kind, payload) = tables.encode_value(value)?;
        let (formula, shared) = tables.encode_formula(formula)?;
        let tag = kind | shared | if stale { STALE_BIT } else { 0 };
        self.sheet.set_encoded(pos, (tag, payload, formula));
        Ok(())
    }

    /// Whether the cell at `pos` holds a formula.
    #[must_use]
    pub fn has_formula_at(&self, pos: CellPos) -> bool {
        self.sheet
            .column(pos.col)
            .and_then(|column| column.entry(pos.row))
            .is_some_and(|(tag, _, formula)| formula != 0 || tag & SHARED_BIT != 0)
    }

    /// Records the cell that carries the formula text of shared group `index`.
    pub fn set_shared_anchor(&mut self, index: u32, anchor: CellPos) {
        self.sheet.shared_anchors.insert(index, anchor);
    }

    /// Stores every cell and shared-formula anchor of `other`, a sheet of another workbook;
    /// on a shared address the cell of `other` wins.
    ///
    /// All strings and formula texts of the workbook of `other` are brought in, so `other`
    /// should belong to a workbook built for that sheet alone.
    ///
    /// # Errors
    ///
    /// Returns a [`LimitError`] when the workbook would hold more strings or formula text
    /// than the store can address.
    pub fn absorb(&mut self, other: &SheetView<'_>) -> Result<(), LimitError> {
        let remap = self.workbook.tables.merge(other.tables)?;
        for (col, source) in &other.sheet.columns {
            // Cells that all follow the last row of this column are appended as a block.
            let below = match (
                self.sheet
                    .column(*col)
                    .and_then(|column| column.rows.last()),
                source.rows.first(),
            ) {
                (Some(last), Some(first)) => last < first,
                _ => true,
            };
            let column = self.sheet.column_mut(*col);
            if below {
                column.reserve(source.rows.len());
            }
            let mut added = 0;
            for index in 0..source.rows.len() {
                let (payload, slot) = remap.cell(
                    source.tags[index],
                    source.payloads[index],
                    source.formulas[index],
                );
                let cell = (source.tags[index], payload, slot);
                if below {
                    column.push(source.rows[index], cell);
                    added += 1;
                } else if column.set(source.rows[index], cell) {
                    added += 1;
                }
            }
            self.sheet.len += added;
        }
        for (index, anchor) in &other.sheet.shared_anchors {
            self.sheet.shared_anchors.insert(*index, *anchor);
        }
        Ok(())
    }

    /// Adds the sheet to the workbook.
    pub fn finish(mut self) {
        for (_, column) in &mut self.sheet.columns {
            column.seal();
        }
        self.workbook.sheets.push(self.sheet);
    }
}

/// The value of a cell, borrowed from the workbook.
#[derive(Debug, Clone, PartialEq)]
pub enum ValueRef<'a> {
    /// No value.
    Empty,
    /// A number; dates are numbers too.
    Number(f64),
    /// A boolean.
    Bool(bool),
    /// Text.
    Text(&'a str),
    /// An error value.
    Error(ErrorKind),
}

/// The formula of a cell, borrowed from the workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaRef<'a> {
    /// Formula text as stored in the file, without a leading `=`.
    Text(&'a str),
    /// A cell of a shared group; see [`SheetView::shared_anchor`].
    Shared {
        /// The group the cell belongs to.
        index: u32,
    },
}

/// One cell of a sheet, read from its columns.
#[derive(Debug, Clone, PartialEq)]
pub struct CellView<'a> {
    /// The constant or cached result.
    pub value: ValueRef<'a>,
    /// The formula, when the cell has one.
    pub formula: Option<FormulaRef<'a>>,
    /// Whether `value` is a cache that `LibreOffice` recalculates on load.
    pub stale: bool,
}

impl CellView<'_> {
    /// The cell as an owned [`Cell`].
    #[must_use]
    pub fn to_cell(&self) -> Cell {
        Cell {
            value: match &self.value {
                ValueRef::Empty => Value::Empty,
                ValueRef::Number(number) => Value::Number(*number),
                ValueRef::Bool(flag) => Value::Bool(*flag),
                ValueRef::Text(text) => Value::Text(Arc::from(*text)),
                ValueRef::Error(kind) => Value::Error(kind.clone()),
            },
            formula: self.formula.map(|formula| match formula {
                FormulaRef::Text(text) => Formula::Text(text.to_owned()),
                FormulaRef::Shared { index } => Formula::Shared { index },
            }),
            stale: self.stale,
        }
    }
}

fn view_at<'a>(tables: &'a Tables, column: &Column, index: usize) -> CellView<'a> {
    let tag = column.tags[index];
    CellView {
        value: tables.decode_value(tag & KIND_MASK, column.payloads[index]),
        formula: tables.decode_formula(tag, column.formulas[index]),
        stale: tag & STALE_BIT != 0,
    }
}

/// A sheet of a workbook, with access to its cells.
#[derive(Debug, Clone, Copy)]
pub struct SheetView<'a> {
    tables: &'a Tables,
    sheet: &'a Sheet,
}

impl<'a> SheetView<'a> {
    pub(crate) fn new(tables: &'a Tables, sheet: &'a Sheet) -> Self {
        Self { tables, sheet }
    }

    /// The sheet name.
    #[must_use]
    pub fn name(&self) -> &'a str {
        self.sheet.name()
    }

    /// Whether the sheet is shown.
    #[must_use]
    pub fn visibility(&self) -> Visibility {
        self.sheet.visibility
    }

    /// The number of non-empty cells.
    #[must_use]
    pub fn len(&self) -> usize {
        self.sheet.len
    }

    /// Whether the sheet has no cells.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.sheet.len == 0
    }

    /// The cell at `pos`, if it is not empty.
    #[must_use]
    pub fn get(&self, pos: CellPos) -> Option<CellView<'a>> {
        let (column, row) = self.sheet.find(pos)?;
        Some(view_at(self.tables, column, row))
    }

    /// The cell that carries the formula text of shared group `index`.
    #[must_use]
    pub fn shared_anchor(&self, index: u32) -> Option<CellPos> {
        self.sheet.shared_anchors.get(&index).copied()
    }

    /// The non-empty cells column by column, each column from the top.
    pub fn cells_by_column(&self) -> impl Iterator<Item = (CellPos, CellView<'a>)> + use<'a> {
        let tables = self.tables;
        self.sheet.columns.iter().flat_map(move |(col, column)| {
            (0..column.rows.len()).map(move |index| {
                let pos = CellPos {
                    row: column.rows[index],
                    col: *col,
                };
                (pos, view_at(tables, column, index))
            })
        })
    }

    /// The non-empty cells row by row. It merges the columns, so it costs more than
    /// [`SheetView::cells_by_column`].
    #[must_use]
    pub fn cells(&self) -> RowMajor<'a> {
        let columns = self.sheet.columns.as_slice();
        let heap = columns
            .iter()
            .enumerate()
            .filter_map(|(index, (_, column))| Some(Reverse((*column.rows.first()?, index))))
            .collect();
        RowMajor {
            tables: self.tables,
            columns,
            cursors: vec![0; columns.len()],
            heap,
        }
    }
}

/// The cells of a sheet in row-major order.
#[derive(Debug)]
pub struct RowMajor<'a> {
    tables: &'a Tables,
    columns: &'a [(u32, Column)],
    cursors: Vec<usize>,
    heap: BinaryHeap<Reverse<(u32, usize)>>,
}

impl<'a> Iterator for RowMajor<'a> {
    type Item = (CellPos, CellView<'a>);

    fn next(&mut self) -> Option<Self::Item> {
        let Reverse((row, index)) = self.heap.pop()?;
        let (col, column) = &self.columns[index];
        let cursor = &mut self.cursors[index];
        let view = view_at(self.tables, column, *cursor);
        *cursor += 1;
        if let Some(next_row) = column.rows.get(*cursor) {
            self.heap.push(Reverse((*next_row, index)));
        }
        Some((CellPos { row, col: *col }, view))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[cfg(target_pointer_width = "64")]
    #[test]
    fn capacities_beyond_32_bits_are_reported() {
        let limit = usize::try_from(u32::MAX).expect("64-bit");
        assert_eq!(string_id(limit), Ok(u32::MAX));
        assert_eq!(string_id(limit + 1), Err(LimitError::Strings));
        assert_eq!(formula_end(limit - 1, 1), Ok(u32::MAX));
        assert_eq!(formula_end(limit, 1), Err(LimitError::FormulaText));
        assert_eq!(formula_end(usize::MAX, 1), Err(LimitError::FormulaText));
        assert_eq!(formula_slot(limit - 1), Ok(u32::MAX));
        assert_eq!(formula_slot(limit), Err(LimitError::Formulas));
    }

    #[test]
    fn pushes_return_consecutive_slots_that_read_back() {
        let mut arena = Arena::default();
        assert_eq!(arena.push("A1"), Ok(1));
        assert_eq!(arena.push("B2"), Ok(2));
        assert_eq!((arena.get(1), arena.get(2)), ("A1", "B2"));
    }
}
