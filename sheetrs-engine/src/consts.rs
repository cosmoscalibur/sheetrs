//! The constants of a program, kept apart so that an executable decodes only the ones a request
//! reads. A program of a whole workbook (`outputs = any`) holds millions of them; the formulas and
//! the rest stay in the program, which then starts in constant time and memory.
//!
//! The constants are grouped by sheet and row into blocks of about [`BLOCK_CELLS`] cells, each
//! encoded on its own. The block index and the extent of each sheet travel with the program; a
//! [`StoredConsts`] reads and decodes a block the first time a cell of it is asked for.

use crate::addr::CellRef;
use crate::engine::ConstSource;
use crate::fastmap::FastMap;
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

/// Move the constants of `program` out of its cells, returning their index and encoded blocks.
/// With fewer than `min` constants the program is left as it is and the index is empty. The slots
/// of `cycles` follow the cells that stay.
pub fn split(program: &mut Program, min: usize) -> Result<(ConstIndex, Vec<u8>), String> {
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
            .ok_or("a constant lies on a sheet the program does not list")?;
        let block: Vec<(u32, u32, &Value)> = consts[i..j]
            .iter()
            .map(|(at, v)| (at.row, at.col, v))
            .collect();
        let encoded = bincode::serialize(&block).map_err(|e| e.to_string())?;
        sh.blocks.push(BlockRef {
            row0: consts[i].0.row,
            offset: bytes.len() as u64,
            len: u32::try_from(encoded.len()).map_err(|_| "a block of constants is too big")?,
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
pub struct StoredConsts {
    index: ConstIndex,
    file: RefCell<File>,
    /// Where the blocks start in `file`.
    start: u64,
    decoded: RefCell<FastMap<(u16, u32), Block>>,
}

impl StoredConsts {
    pub fn new(index: ConstIndex, file: File, start: u64) -> Self {
        Self {
            index,
            file: RefCell::new(file),
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
        let mut buf = vec![0u8; b.len as usize];
        let mut f = self.file.borrow_mut();
        f.seek(SeekFrom::Start(self.start + b.offset))
            .and_then(|_| f.read_exact(&mut buf))
            .unwrap_or_else(|e| panic!("cannot read the constants of the program: {e}"));
        bincode::deserialize(&buf)
            .unwrap_or_else(|e| panic!("the constants of the program are damaged: {e}"))
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
        let block = decoded
            .entry((at.sheet, b as u32))
            .or_insert_with(|| self.load(&sh.blocks[b]));
        let i = block
            .binary_search_by_key(&(at.row, at.col), |(r, c, _)| (*r, *c))
            .ok()?;
        Some(block[i].2.clone())
    }

    fn extent(&self, sheet: u16) -> (u32, u32) {
        self.index
            .sheets
            .get(sheet as usize)
            .map_or((0, 0), |s| s.extent)
    }
}
