//! The used area of a sheet: one definition, one accumulator, for every format and every consumer.

use super::{Cell, CellSink};

/// The rectangle of the **used cells** of a sheet, 0-based and inclusive: from the first used row
/// and column to the last.
///
/// A cell is used when it holds a value or a formula, in an XLSX and in an ODS alike (see
/// [`is_used`]); a cell kept only for its style, a row or column that is only hidden or sized, and
/// the `<dimension>` an XLSX declares do not count. It is the used area of `LibreOffice`
/// (`GetCellArea`, which looks at the cell contents), so a sheet whose only content is at `D10` has
/// the area `D10:D10`, and a sheet without content has none.
///
/// Both ways of reading a workbook build it with [`AreaSink`], as the cells are read: the loaded
/// model ([`Sheet::used_area`](super::Sheet::used_area)) and the streaming consumers (the statistics
/// read it from the sheet, they keep no area of their own).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct UsedArea {
    /// First used row.
    pub first_row: u32,
    /// First used column.
    pub first_col: u32,
    /// Last used row.
    pub last_row: u32,
    /// Last used column.
    pub last_col: u32,
}

impl UsedArea {
    /// The area of the single cell at `(row, col)`.
    #[must_use]
    pub fn at(row: u32, col: u32) -> Self {
        Self {
            first_row: row,
            first_col: col,
            last_row: row,
            last_col: col,
        }
    }

    /// The area from `A1` to `(last_row, last_col)`.
    #[must_use]
    pub fn from_a1(last_row: u32, last_col: u32) -> Self {
        Self {
            first_row: 0,
            first_col: 0,
            last_row,
            last_col,
        }
    }

    /// The area that holds this one and the cell at `(row, col)`.
    #[must_use]
    pub fn including(self, row: u32, col: u32) -> Self {
        Self {
            first_row: self.first_row.min(row),
            first_col: self.first_col.min(col),
            last_row: self.last_row.max(row),
            last_col: self.last_col.max(col),
        }
    }

    /// The extent counted from `A1`, as `(rows, columns)`: the rows `0..=last_row` and the columns
    /// `0..=last_col`.
    #[must_use]
    pub fn extent_from_a1(&self) -> (u32, u32) {
        (self.last_row + 1, self.last_col + 1)
    }

    /// The number of cells of the rectangle.
    #[must_use]
    pub fn cells(&self) -> u64 {
        (u64::from(self.last_row - self.first_row) + 1)
            * (u64::from(self.last_col - self.first_col) + 1)
    }

    /// The number of cells from `A1` to the last used cell.
    #[must_use]
    pub fn cells_from_a1(&self) -> u64 {
        let (rows, cols) = self.extent_from_a1();
        u64::from(rows) * u64::from(cols)
    }
}

/// Whether a cell counts as used: it holds a value or a formula. Cells kept only for their style do
/// not.
#[must_use]
pub fn is_used(cell: &Cell) -> bool {
    cell.is_formula() || !cell.value.is_empty()
}

/// The two rectangles the reader finds in a sheet, from the same pass over its cells.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct SheetAreas {
    /// The cells with a value or a formula: [`Sheet::used_area`](super::Sheet::used_area).
    pub used: Option<UsedArea>,
    /// Every cell the file keeps, with a value, a formula or only a style:
    /// [`Sheet::kept_area`](super::Sheet::kept_area).
    pub kept: Option<UsedArea>,
}

/// A sink that hands every cell on to `inner` and accumulates the [`SheetAreas`] of each sheet. The
/// readers put it in front of whatever sink they are given, so the areas do not depend on who
/// reads: the cells are not copied or kept, a sheet costs eight numbers.
pub struct AreaSink<'a> {
    inner: &'a mut dyn CellSink,
    areas: Vec<SheetAreas>,
}

impl<'a> AreaSink<'a> {
    #[must_use]
    pub fn new(inner: &'a mut dyn CellSink) -> Self {
        Self {
            inner,
            areas: Vec::new(),
        }
    }

    /// The areas of each sheet that was read, by sheet index.
    #[must_use]
    pub fn into_areas(self) -> Vec<SheetAreas> {
        self.areas
    }
}

impl CellSink for AreaSink<'_> {
    fn hint(&mut self, sheet: u16, cells: usize) {
        self.inner.hint(sheet, cells);
    }

    fn cell(&mut self, sheet: u16, at: (u32, u32), cell: Cell) {
        let i = usize::from(sheet);
        if self.areas.len() <= i {
            self.areas.resize(i + 1, SheetAreas::default());
        }
        let grow = |area: Option<UsedArea>| {
            Some(area.map_or_else(|| UsedArea::at(at.0, at.1), |a| a.including(at.0, at.1)))
        };
        let areas = &mut self.areas[i];
        areas.kept = grow(areas.kept);
        if is_used(&cell) {
            areas.used = grow(areas.used);
        }
        self.inner.cell(sheet, at, cell);
    }

    fn comment(&mut self, sheet: u16, at: (u32, u32), text: &str) {
        self.inner.comment(sheet, at, text);
    }

    fn sheet(&mut self, index: u16, name: &str) {
        self.inner.sheet(index, name);
    }

    fn filtered_rows(&mut self, sheet: u16, first: u32, last: u32) {
        self.inner.filtered_rows(sheet, first, last);
    }

    fn array_member(&mut self, sheet: u16, at: (u32, u32), offset: (u32, u32)) {
        self.inner.array_member(sheet, at, offset);
    }
}

/// Puts the areas the [`AreaSink`] found into the sheets it was reading. The kept area also reaches
/// the last hidden row and column of the sheet, which the file keeps as well.
pub(crate) fn assign(sheets: &mut [super::Sheet], areas: &[SheetAreas]) {
    for (sheet, areas) in sheets.iter_mut().zip(areas) {
        sheet.used_area = areas.used;
        let hidden_row = sheet.hidden_rows.iter().map(|&(_, last)| last).max();
        let hidden_col = sheet.hidden_columns.iter().map(|&(_, last)| last).max();
        sheet.kept_area = match (areas.kept, hidden_row, hidden_col) {
            (kept, None, None) => kept,
            (Some(kept), row, col) => Some(UsedArea {
                last_row: row.map_or(kept.last_row, |row| row.max(kept.last_row)),
                last_col: col.map_or(kept.last_col, |col| col.max(kept.last_col)),
                ..kept
            }),
            (None, row, col) => Some(UsedArea::from_a1(row.unwrap_or(0), col.unwrap_or(0))),
        };
    }
}
