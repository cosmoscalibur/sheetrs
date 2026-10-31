//! The cells of a model by position: one [`Grid`] per sheet, the store the reader's other consumer
//! (the linter's workbook) uses too. A workbook reader finds cells in row order, so reading
//! appends; a lookup is two array reads and a search in a row's few columns, with no hashing and
//! no table to grow, and a range read walks neighbouring memory.

use crate::addr::CellRef;
use crate::model::ModelCell;
use sheetrs_formula::grid::Grid;
use std::ops::Index;

#[derive(Debug, Clone, Default)]
pub struct CellStore {
    sheets: Vec<Grid<ModelCell>>,
    len: usize,
}

impl CellStore {
    pub fn len(&self) -> usize {
        self.len
    }

    pub fn is_empty(&self) -> bool {
        self.len == 0
    }

    /// Room for `cells` more cells in `sheet`.
    pub fn reserve(&mut self, sheet: u16, cells: usize) {
        self.sheet_mut(sheet).reserve(cells);
    }

    fn sheet_mut(&mut self, sheet: u16) -> &mut Grid<ModelCell> {
        let i = usize::from(sheet);
        if self.sheets.len() <= i {
            self.sheets.resize_with(i + 1, Grid::new);
        }
        &mut self.sheets[i]
    }

    pub fn get(&self, at: &CellRef) -> Option<&ModelCell> {
        self.sheets
            .get(usize::from(at.sheet))?
            .get(&(at.row, at.col))
    }

    pub fn contains_key(&self, at: &CellRef) -> bool {
        self.get(at).is_some()
    }

    /// Set a cell; a cell already there is replaced and returned. Appending in row order, as a
    /// reader does, is constant time; any other position moves the cells after it.
    pub fn insert(&mut self, at: CellRef, cell: ModelCell) -> Option<ModelCell> {
        let old = self.sheet_mut(at.sheet).insert((at.row, at.col), cell);
        if old.is_none() {
            self.len += 1;
        }
        old
    }

    /// The cells in order of sheet, row and column.
    pub fn iter(&self) -> impl Iterator<Item = (CellRef, &ModelCell)> {
        self.sheets.iter().enumerate().flat_map(|(s, grid)| {
            grid.iter()
                .map(move |((row, col), cell)| (CellRef::new(s as u16, row, col), cell))
        })
    }

    pub fn keys(&self) -> impl Iterator<Item = CellRef> {
        self.iter().map(|(at, _)| at)
    }
}

impl<'a> IntoIterator for &'a CellStore {
    type Item = (CellRef, &'a ModelCell);
    type IntoIter = Box<dyn Iterator<Item = (CellRef, &'a ModelCell)> + 'a>;

    fn into_iter(self) -> Self::IntoIter {
        Box::new(self.iter())
    }
}

impl Index<&CellRef> for CellStore {
    type Output = ModelCell;

    fn index(&self, at: &CellRef) -> &ModelCell {
        self.get(at).expect("no cell at that position")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::value::Value;

    fn cell(n: f64) -> ModelCell {
        ModelCell {
            value: Value::Number(n),
            formula: None,
            array: false,
        }
    }

    #[test]
    fn orders_by_sheet_row_and_column_across_sheets() {
        let mut s = CellStore::default();
        s.insert(CellRef::new(1, 3, 3), cell(33.0));
        s.insert(CellRef::new(0, 9, 9), cell(99.0));
        s.insert(CellRef::new(0, 0, 1), cell(1.0));
        assert_eq!(s.len(), 3);
        let order: Vec<_> = s.keys().map(|a| (a.sheet, a.row, a.col)).collect();
        assert_eq!(order, vec![(0, 0, 1), (0, 9, 9), (1, 3, 3)]);
        assert_eq!(s[&CellRef::new(1, 3, 3)].value, Value::Number(33.0));
        assert!(s.get(&CellRef::new(7, 0, 0)).is_none());
    }

    #[test]
    fn a_cell_at_the_same_position_replaces_the_old_one() {
        let mut s = CellStore::default();
        s.insert(CellRef::new(0, 1, 1), cell(1.0));
        let old = s.insert(CellRef::new(0, 1, 1), cell(2.0));
        assert_eq!(old.map(|c| c.value), Some(Value::Number(1.0)));
        assert_eq!(s.len(), 1);
        assert_eq!(s[&CellRef::new(0, 1, 1)].value, Value::Number(2.0));
    }
}
