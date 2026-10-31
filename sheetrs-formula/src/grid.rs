//! Cells by position, stored by row: the one store of cells of the workspace.
//!
//! A sorted list of columns and items per sheet, with the offset where each row starts. A reader
//! finds cells in row order, so reading appends; a lookup is two array reads and a search in a
//! row's few columns, with no hashing and no table to grow, and a walk goes through neighbouring
//! memory in row order. Both consumers of the reader keep their cells in it: the linter's
//! workbook as `Grid<Cell>`, and the compiler's model as `Grid<ModelCell>`. It lives here, below
//! the readers, so the compiler can use it without them.

use std::collections::HashMap;
use std::ops::Index;

/// Items by `(row, column)`, 0-based.
#[derive(Debug, Clone)]
pub struct Grid<T> {
    /// `row_start[r]..row_start[r + 1]` are the entries of row `r`; a row past the end is empty.
    row_start: Vec<u32>,
    cols: Vec<u32>,
    items: Vec<T>,
}

impl<T> Default for Grid<T> {
    fn default() -> Self {
        Self {
            row_start: Vec::new(),
            cols: Vec::new(),
            items: Vec::new(),
        }
    }
}

impl<T> Grid<T> {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn len(&self) -> usize {
        self.items.len()
    }

    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }

    /// Room for `items` more items.
    pub fn reserve(&mut self, items: usize) {
        self.cols.reserve(items);
        self.items.reserve(items);
    }

    /// Give back the room that was reserved and not used.
    pub fn shrink_to_fit(&mut self) {
        self.row_start.shrink_to_fit();
        self.cols.shrink_to_fit();
        self.items.shrink_to_fit();
    }

    fn rows(&self) -> usize {
        self.row_start.len().saturating_sub(1)
    }

    fn row(&self, row: u32) -> std::ops::Range<usize> {
        let r = row as usize;
        if r < self.rows() {
            self.row_start[r] as usize..self.row_start[r + 1] as usize
        } else {
            0..0
        }
    }

    /// `Ok(index)` of the item, or `Err(index)` where it would go.
    fn locate(&self, row: u32, col: u32) -> Result<usize, usize> {
        if row as usize >= self.rows() {
            return Err(self.items.len());
        }
        let range = self.row(row);
        match self.cols[range.clone()].binary_search(&col) {
            Ok(i) => Ok(range.start + i),
            Err(i) => Err(range.start + i),
        }
    }

    pub fn get(&self, at: &(u32, u32)) -> Option<&T> {
        self.locate(at.0, at.1).ok().map(|i| &self.items[i])
    }

    pub fn get_mut(&mut self, at: &(u32, u32)) -> Option<&mut T> {
        self.locate(at.0, at.1).ok().map(|i| &mut self.items[i])
    }

    pub fn contains_key(&self, at: &(u32, u32)) -> bool {
        self.get(at).is_some()
    }

    /// Set an item; one already there is replaced and returned. Appending in row order, as a
    /// reader does, is constant time; any other position moves the items after it.
    pub fn insert(&mut self, at: (u32, u32), item: T) -> Option<T> {
        let (row, col) = (at.0 as usize, at.1);
        // After every item stored: a new row, or a later column of the last row. No search, no
        // shifting of what follows.
        let appends = row >= self.rows()
            || (row + 1 == self.rows() && self.cols.last().is_some_and(|&c| c < col));
        if appends {
            let end = u32::try_from(self.items.len()).expect("more than 2^32 cells in a sheet");
            while self.row_start.len() < row + 1 {
                self.row_start.push(end);
            }
            if self.row_start.len() == row + 1 {
                self.row_start.push(end);
            }
            self.row_start[row + 1] = end + 1;
            self.cols.push(col);
            self.items.push(item);
            return None;
        }
        match self.locate(at.0, col) {
            Ok(i) => Some(std::mem::replace(&mut self.items[i], item)),
            Err(i) => {
                for start in &mut self.row_start[row + 1..] {
                    *start += 1;
                }
                self.cols.insert(i, col);
                self.items.insert(i, item);
                None
            }
        }
    }

    /// The items in order of row and column.
    pub fn iter(&self) -> impl Iterator<Item = ((u32, u32), &T)> {
        (0..self.rows()).flat_map(move |r| {
            self.row(r as u32)
                .map(move |i| ((r as u32, self.cols[i]), &self.items[i]))
        })
    }

    pub fn keys(&self) -> impl Iterator<Item = (u32, u32)> {
        self.iter().map(|(at, _)| at)
    }

    /// The items in order of row and column.
    pub fn values(&self) -> impl Iterator<Item = &T> {
        self.items.iter()
    }
}

impl<T> Index<&(u32, u32)> for Grid<T> {
    type Output = T;

    fn index(&self, at: &(u32, u32)) -> &T {
        self.get(at).expect("no cell at that position")
    }
}

impl<'a, T> IntoIterator for &'a Grid<T> {
    type Item = ((u32, u32), &'a T);
    type IntoIter = Box<dyn Iterator<Item = ((u32, u32), &'a T)> + 'a>;

    fn into_iter(self) -> Self::IntoIter {
        Box::new(self.iter())
    }
}

impl<T> IntoIterator for Grid<T> {
    type Item = ((u32, u32), T);
    type IntoIter = std::vec::IntoIter<((u32, u32), T)>;

    fn into_iter(self) -> Self::IntoIter {
        let mut rows = Vec::with_capacity(self.items.len());
        let mut r = 0u32;
        for (i, item) in self.items.into_iter().enumerate() {
            while self.row_start[r as usize + 1] as usize <= i {
                r += 1;
            }
            rows.push(((r, self.cols[i]), item));
        }
        rows.into_iter()
    }
}

impl<T> FromIterator<((u32, u32), T)> for Grid<T> {
    fn from_iter<I: IntoIterator<Item = ((u32, u32), T)>>(iter: I) -> Self {
        let mut grid = Self::new();
        for (at, item) in iter {
            grid.insert(at, item);
        }
        grid
    }
}

impl<T> From<HashMap<(u32, u32), T>> for Grid<T> {
    fn from(map: HashMap<(u32, u32), T>) -> Self {
        let mut items: Vec<_> = map.into_iter().collect();
        items.sort_unstable_by_key(|(at, _)| *at);
        items.into_iter().collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn appends_in_row_order_and_inserts_anywhere() {
        let mut g = Grid::new();
        for (r, c) in [(0, 0), (0, 3), (2, 1), (2, 2), (5, 0)] {
            assert!(g.insert((r, c), r * 10 + c).is_none());
        }
        // Out of order, into an earlier row, a later column and a row between.
        g.insert((2, 0), 20);
        g.insert((0, 1), 1);
        g.insert((1, 7), 17);
        g.insert((9, 9), 99);
        assert_eq!(g.len(), 9);
        let order: Vec<_> = g.keys().collect();
        let mut sorted = order.clone();
        sorted.sort_unstable();
        assert_eq!(order, sorted);
        for (a, v) in [((0, 1), 1), ((1, 7), 17), ((2, 0), 20), ((9, 9), 99)] {
            assert_eq!(g[&a], v);
        }
        assert!(g.get(&(1, 0)).is_none());
        assert!(g.get(&(4, 4)).is_none());
        assert!(g.get(&(100, 0)).is_none());
        assert_eq!(g.values().count(), 9);
    }

    #[test]
    fn an_item_at_the_same_position_replaces_the_old_one() {
        let mut g = Grid::new();
        g.insert((1, 1), 1);
        assert_eq!(g.insert((1, 1), 2), Some(1));
        assert_eq!(g.len(), 1);
        assert_eq!(g[&(1, 1)], 2);
    }

    #[test]
    fn converts_from_a_map_and_into_pairs() {
        let map: HashMap<(u32, u32), char> = [((3, 1), 'c'), ((0, 2), 'a'), ((3, 0), 'b')]
            .into_iter()
            .collect();
        let g = Grid::from(map);
        assert_eq!(
            g.into_iter().collect::<Vec<_>>(),
            vec![((0, 2), 'a'), ((3, 0), 'b'), ((3, 1), 'c')]
        );
    }

    #[test]
    fn an_empty_grid_has_nothing() {
        let g: Grid<u8> = Grid::new();
        assert!(g.is_empty() && g.get(&(0, 0)).is_none() && g.iter().next().is_none());
        assert_eq!(g.into_iter().count(), 0);
    }
}
