//! The constants of a program, kept apart so that an executable decodes only the ones a request
//! reads. A program of a whole workbook (no outputs declared) holds millions of them; the formulas and
//! the rest stay in the program, which then starts in constant time and memory.
//!
//! The constants are grouped by sheet and row into blocks of about [`BLOCK_CELLS`] cells, each
//! encoded on its own. The block index and the extent of each sheet travel with the program; a
//! [`StoredConsts`] reads and decodes a block the first time a cell of it is asked for.

use crate::addr::CellRef;
use crate::engine::ConstSource;
use crate::ir::{CellDef, CellKind, Program};
use crate::value::Value;
use serde::{Deserialize, Serialize};
use std::cell::RefCell;
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};

/// Constants a block holds, at least (whole rows go together).
pub const BLOCK_CELLS: usize = 4096;

/// Programs with fewer constants keep them in the program: it is as quick to decode them all.
pub const MIN_STORED: usize = 50_000;

/// Where the blocks of a program are and what they span.
#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct ConstIndex {
    sheets: Vec<SheetConsts>,
    cells: u64,
}

impl ConstIndex {
    /// The number of constants kept apart.
    #[must_use]
    pub fn cells(&self) -> u64 {
        self.cells
    }
}

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
struct SheetConsts {
    /// Rows and columns spanned by the constants.
    extent: (u32, u32),
    /// Sorted by `row0`: a block holds the rows from its `row0` up to the next block's.
    blocks: Vec<BlockRef>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
struct BlockRef {
    row0: u32,
    /// Position of the encoded block among the blocks, and its size.
    offset: u64,
    len: u32,
}

/// Cells `(row, col, value)` sorted by position.
type Block = Vec<(u32, u32, Value)>;

/// A decoded block with where each of its rows starts, so that a cell is found by its row and then by
/// its column among the few its row holds, not by a search over the whole block.
///
/// A block whose rows are close together keeps one entry per row between its first and its last (the row
/// is a subtraction); a block that spans far more rows than it has cells (a few constants at both ends of
/// a sheet) keeps only the rows that hold cells, found by binary search, so the table is proportional to
/// the cells and not to the distance between them.
#[derive(Debug)]
struct Decoded {
    row0: u32,
    /// Empty when `starts` has one entry per row from `row0`; otherwise the rows that hold cells, sorted.
    rows: Vec<u32>,
    /// `starts[r]` is the position of the first cell of the `r`-th row (of the span, or of `rows`); one more
    /// entry ends the last.
    starts: Vec<u32>,
    cells: Block,
}

/// A block is dense when its span is at most this many rows per cell it holds (and a fixed allowance).
const DENSE_ROWS_PER_CELL: usize = 4;

impl Decoded {
    fn new(cells: Block) -> Self {
        let row0 = cells.first().map_or(0, |c| c.0);
        let last = cells.last().map_or(0, |c| c.0);
        let span = (last - row0) as usize + 1;
        let mut starts = Vec::new();
        let mut rows = Vec::new();
        if span <= DENSE_ROWS_PER_CELL * cells.len() + 64 {
            starts.reserve(span + 1);
            let mut at = 0u32;
            for row in row0..=last {
                starts.push(at);
                while (at as usize) < cells.len() && cells[at as usize].0 == row {
                    at += 1;
                }
            }
            starts.push(at);
        } else {
            for (i, c) in cells.iter().enumerate() {
                if rows.last() != Some(&c.0) {
                    rows.push(c.0);
                    starts.push(i as u32);
                }
            }
            starts.push(cells.len() as u32);
        }
        Self {
            row0,
            rows,
            starts,
            cells,
        }
    }

    fn get(&self, row: u32, col: u32) -> Option<&Value> {
        let r = if self.rows.is_empty() {
            row.checked_sub(self.row0)? as usize
        } else {
            self.rows.binary_search(&row).ok()?
        };
        let (from, to) = (
            *self.starts.get(r)? as usize,
            *self.starts.get(r + 1)? as usize,
        );
        let cells = &self.cells[from..to];
        let i = cells.binary_search_by_key(&col, |c| c.1).ok()?;
        Some(&cells[i].2)
    }
}

/// Move the constants of `program` out of its cells, returning their index and encoded blocks.
/// With fewer than `min` constants the program is left as it is and the index is empty. The slots
/// of `cycles` follow the cells that stay.
///
/// # Errors
///
/// When the constants cannot be split off the program.
pub fn split(program: &mut Program, min: usize) -> crate::Result<(ConstIndex, Vec<u8>)> {
    let total = program
        .cells
        .iter()
        .filter(|c| matches!(c.kind, CellKind::Const(_)))
        .count();
    if total < min || total == 0 {
        return Ok((ConstIndex::default(), Vec::new()));
    }
    let cells = std::mem::take(&mut program.cells);
    let mut kept = Vec::with_capacity(cells.len() - total);
    // Slots before each old slot that stay, to renumber the cycles.
    let mut kept_before = Vec::with_capacity(cells.len() + 1);
    let mut consts: Vec<(CellRef, Value)> = Vec::with_capacity(total);
    for c in cells {
        kept_before.push(kept.len() as u32);
        match c {
            CellDef {
                at,
                kind: CellKind::Const(v),
            } => consts.push((at, v)),
            c => kept.push(c),
        }
    }
    kept_before.push(kept.len() as u32);
    program.cells = kept;
    for (s, e) in &mut program.cycles {
        *s = kept_before[*s as usize];
        *e = kept_before[*e as usize];
    }
    consts.sort_unstable_by_key(|(at, _)| *at);

    let mut index = ConstIndex {
        sheets: vec![SheetConsts::default(); program.sheets.len().max(1)],
        cells: total as u64,
    };
    let mut bytes = Vec::new();
    let mut i = 0;
    while i < consts.len() {
        let sheet = consts[i].0.sheet;
        let mut j = i;
        loop {
            // Whole rows: the block ends where a row does.
            let row = consts[j].0.row;
            while j < consts.len() && consts[j].0.sheet == sheet && consts[j].0.row == row {
                j += 1;
            }
            if j - i >= BLOCK_CELLS || j >= consts.len() || consts[j].0.sheet != sheet {
                break;
            }
        }
        let sh = index
            .sheets
            .get_mut(sheet as usize)
            .ok_or(crate::Error::ConstantOnUnknownSheet)?;
        let block: Vec<(u32, u32, &Value)> = consts[i..j]
            .iter()
            .map(|(at, v)| (at.row, at.col, v))
            .collect();
        let encoded = postcard::to_allocvec(&block)?;
        sh.blocks.push(BlockRef {
            row0: consts[i].0.row,
            offset: bytes.len() as u64,
            len: u32::try_from(encoded.len()).map_err(|_| crate::Error::BlockTooBig)?,
        });
        for (at, _) in &consts[i..j] {
            sh.extent.0 = sh.extent.0.max(at.row + 1);
            sh.extent.1 = sh.extent.1.max(at.col + 1);
        }
        bytes.extend_from_slice(&encoded);
        i = j;
    }
    Ok((index, bytes))
}

/// The constants of an executable, decoded block by block as they are read.
#[derive(Debug)]
pub struct StoredConsts {
    index: ConstIndex,
    backing: Backing,
    /// Where the blocks start in the backing.
    start: u64,
    /// Per sheet, per block of the sheet: decoded on the first read of a cell of it.
    decoded: RefCell<Vec<Vec<Option<Decoded>>>>,
}

/// Where the blocks are read from: the executable's file, or memory (a module that carries its
/// program in its own linear memory has no file).
#[derive(Debug)]
enum Backing {
    File(RefCell<File>),
    Memory(&'static [u8]),
}

impl StoredConsts {
    #[must_use]
    pub fn new(index: ConstIndex, file: File, start: u64) -> Self {
        Self {
            index,
            backing: Backing::File(RefCell::new(file)),
            start,
            decoded: RefCell::default(),
        }
    }

    /// Constants whose blocks are in `memory`, from `start`.
    #[must_use]
    pub fn from_memory(index: ConstIndex, memory: &'static [u8], start: u64) -> Self {
        Self {
            index,
            backing: Backing::Memory(memory),
            start,
            decoded: RefCell::default(),
        }
    }

    /// The number of constants.
    pub fn len(&self) -> u64 {
        self.index.cells
    }

    pub fn is_empty(&self) -> bool {
        self.index.cells == 0
    }

    fn load(&self, b: &BlockRef) -> Block {
        let damaged = |e: &dyn std::fmt::Display| -> ! {
            panic!("cannot read the constants of the program: {e}")
        };
        let from = usize::try_from(self.start + b.offset).unwrap_or(usize::MAX);
        match &self.backing {
            Backing::File(file) => {
                let mut buf = vec![0u8; b.len as usize];
                let mut f = file.borrow_mut();
                f.seek(SeekFrom::Start(self.start + b.offset))
                    .and_then(|_| f.read_exact(&mut buf))
                    .unwrap_or_else(|e| damaged(&e));
                postcard::from_bytes(&buf)
                    .unwrap_or_else(|e| panic!("the constants of the program are damaged: {e}"))
            }
            Backing::Memory(memory) => {
                let bytes = from
                    .checked_add(b.len as usize)
                    .and_then(|to| memory.get(from..to))
                    .unwrap_or_else(|| damaged(&"a block is out of the memory"));
                postcard::from_bytes(bytes)
                    .unwrap_or_else(|e| panic!("the constants of the program are damaged: {e}"))
            }
        }
    }
}

impl ConstSource for StoredConsts {
    fn get(&self, at: CellRef) -> Option<Value> {
        let sh = self.index.sheets.get(at.sheet as usize)?;
        if at.row >= sh.extent.0 || at.col >= sh.extent.1 {
            return None;
        }
        let b = sh
            .blocks
            .partition_point(|b| b.row0 <= at.row)
            .checked_sub(1)?;
        let mut decoded = self.decoded.borrow_mut();
        if decoded.len() <= at.sheet as usize {
            decoded.resize_with(self.index.sheets.len().max(at.sheet as usize + 1), Vec::new);
        }
        let blocks = &mut decoded[at.sheet as usize];
        if blocks.is_empty() {
            blocks.resize_with(sh.blocks.len(), || None);
        }
        let block = blocks[b].get_or_insert_with(|| Decoded::new(self.load(&sh.blocks[b])));
        block.get(at.row, at.col).cloned()
    }

    fn extent(&self, sheet: u16) -> (u32, u32) {
        self.index
            .sheets
            .get(sheet as usize)
            .map_or((0, 0), |s| s.extent)
    }
}

#[cfg(test)]
mod decoded_tests {
    use super::*;

    fn block(rows: &[u32]) -> Block {
        rows.iter()
            .enumerate()
            .flat_map(|(i, &r)| {
                [
                    (r, 1, Value::Number(i as f64)),
                    (r, 7, Value::Number(-(i as f64))),
                ]
            })
            .collect()
    }

    /// A few constants at both ends of a sheet cost entries for the rows that hold them, not for the distance.
    #[test]
    fn a_sparse_block_keeps_a_table_of_the_size_of_its_cells() {
        let rows = [0, 1_048_575];
        let d = Decoded::new(block(&rows));
        assert!(
            d.starts.len() <= 3 && d.rows.len() == 2,
            "{} entries",
            d.starts.len()
        );
        assert_eq!(d.get(0, 1), Some(&Value::Number(0.0)));
        assert_eq!(d.get(1_048_575, 7), Some(&Value::Number(-1.0)));
        assert_eq!(d.get(5, 1), None);
        assert_eq!(d.get(0, 2), None);
        assert_eq!(d.get(2_000_000, 1), None);
    }

    /// Dense and sparse blocks answer as a scan of the cells does.
    #[test]
    fn both_layouts_agree_with_a_scan() {
        let mut state = 12345u64;
        let mut next = move |m: u32| {
            state = state
                .wrapping_mul(6_364_136_223_846_793_005)
                .wrapping_add(1_442_695_040_888_963_407);
            ((state >> 33) as u32) % m
        };
        for span in [10u32, 200, 100_000] {
            let mut rows: Vec<u32> = (0..60).map(|_| 5 + next(span)).collect();
            rows.sort_unstable();
            rows.dedup();
            let cells = block(&rows);
            let d = Decoded::new(cells.clone());
            assert_eq!(d.rows.is_empty(), span < 1000, "layout for span {span}");
            for _ in 0..2000 {
                let (r, c) = (next(span + 20), [1, 7, 3][next(3) as usize]);
                let want = cells.iter().find(|x| x.0 == r && x.1 == c).map(|x| &x.2);
                assert_eq!(d.get(r, c), want, "row {r} col {c} span {span}");
            }
        }
    }
}
