//! The execution engine: a demand-driven evaluator over a [`Program`].
//!
//! Cells are normally evaluated in the precomputed topological order, so every
//! reference finds its operand ready. References that the compiler could not
//! resolve statically (`INDIRECT`, `OFFSET` with run-time arguments) simply
//! demand the target cell: it is evaluated on the spot, memoized, and a cell
//! found in progress is a circular reference. That hybrid keeps one engine for
//! both the fully static and the dynamic case.

use crate::addr::{CellRef, Rect};
use crate::coerce;
use crate::fastmap::FastMap;
use crate::ir::{BinOp, CellKind, Expr, Program, UnOp};
use crate::options::EngineOptions;
use crate::value::{Array, Collators, ErrorKind, Value, compare};
use std::cmp::Ordering;
use std::sync::Arc;

pub type Res<T> = Result<T, ErrorKind>;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum St {
    Pending,
    Busy,
    Done,
    /// Member of a cycle being iterated: reads return the current value.
    Iter,
}

/// Cells that are not part of the program: constants of the source workbook,
/// read in place while folding so they need not be copied into a program.
pub trait ConstSource {
    fn get(&self, at: CellRef) -> Option<Value>;
    /// Rows and columns spanned on a sheet.
    fn extent(&self, sheet: u16) -> (u32, u32);
}

/// First position of each text (case-folded) and number among the keys of a lookup
/// column.
#[derive(Debug, Default)]
pub struct LookupIndex {
    text: crate::fastmap::FastMap<String, usize>,
    number: crate::fastmap::FastMap<u64, usize>,
}

impl LookupIndex {
    fn add(&mut self, v: &Value, at: usize) {
        match v {
            Value::Text(t) => {
                self.text
                    .entry(crate::casemap::fold(t).into_owned())
                    .or_insert(at);
            }
            // -0 and 0 are the same key, and an empty cell counts as 0.
            Value::Number(n) => {
                self.number.entry((n + 0.0).to_bits()).or_insert(at);
            }
            Value::Empty => {
                self.number.entry(0f64.to_bits()).or_insert(at);
            }
            _ => {}
        }
    }

    #[must_use]
    pub fn find_text(&self, folded: &str) -> Option<usize> {
        self.text.get(folded).copied()
    }

    #[must_use]
    pub fn find_number(&self, n: f64) -> Option<usize> {
        self.number.get(&(n + 0.0).to_bits()).copied()
    }
}

/// Ranges under this many cells are cheaper to rebuild than to remember.
const SMALL_RANGE: u64 = 128;

/// The format a formula ends with, so far (`Engine::logical_fmt`): none that is boolean.
pub(crate) const FMT_NONE: u8 = 0;
/// A boolean format that a function left on the number it gives (`VALUE(TRUE())`).
pub(crate) const FMT_KEPT: u8 = 1;
/// The one a minus left on a boolean operand (`-TRUE()`): `ISLOGICAL` of it is true.
pub(crate) const FMT_NEGATED: u8 = 2;

pub struct Engine<'p> {
    prog: &'p Program,
    base: Option<&'p dyn ConstSource>,
    pub opts: EngineOptions,
    index: FastMap<CellRef, u32>,
    by_sheet: Vec<Vec<u32>>,
    /// Per sheet: (rows, cols) spanned by its cells.
    extent: Vec<(u32, u32)>,
    values: Vec<Value>,
    /// Formula cells that ended in a blank. `LibreOffice` keeps such an "empty
    /// result" (cached as 0) that still equals `""` and is not a number.
    blank: Vec<bool>,
    state: Vec<St>,
    /// Cell whose formula is being evaluated.
    pub at: CellRef,
    rng: u64,
    /// >0 while evaluating arguments of array-context functions.
    array_depth: u32,
    /// Materialized ranges whose cells were all final when read.
    range_cache: FastMap<Rect, Arc<Array>>,
    /// The same ranges as [`range_array_full`](Engine::range_array_full) sees them.
    full_cache: FastMap<Rect, Arc<Array>>,
    /// First position of each key of a lookup column, for ranges whose values are settled.
    lookup_cache: FastMap<(Rect, bool), Arc<LookupIndex>>,
    /// Reads that saw a cell still in flux (circular or iterating).
    unstable_reads: u64,
    /// Buffer for the values of a small range read by `numbers`, kept between calls.
    scratch: Vec<Value>,
    coll: Collators,
    /// Formula cells that met something `LibreOffice` does and the engine does not reproduce
    /// (see `crate::notes`).
    pub violations: Vec<(CellRef, &'static str)>,
    /// Whether the formula that is being evaluated ends, so far, with the boolean format
    /// `LibreOffice` gives it: a number that the cell shows as `TRUE` or `FALSE` (`VALUE(TRUE())`).
    /// What sets a number format (a function, an operator) clears it; what keeps the one of its
    /// operand (`VALUE`, a minus on a boolean) sets it; `finalize` reads it.
    pub(crate) logical_fmt: u8,
}

impl std::fmt::Debug for Engine<'_> {
    /// The options and the size of what it holds; its state is not worth printing.
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Engine")
            .field("opts", &self.opts)
            .field("cells", &self.values.len())
            .finish_non_exhaustive()
    }
}

impl<'p> Engine<'p> {
    #[must_use]
    pub fn new(prog: &'p Program, opts: EngineOptions) -> Self {
        Self::build(prog, opts, None)
    }

    /// An engine whose cells missing from `prog` come from `base`.
    pub fn with_base(prog: &'p Program, opts: EngineOptions, base: &'p dyn ConstSource) -> Self {
        Self::build(prog, opts, Some(base))
    }

    fn build(prog: &'p Program, opts: EngineOptions, base: Option<&'p dyn ConstSource>) -> Self {
        let n = prog.cells.len();
        let mut index =
            FastMap::with_capacity_and_hasher(n, std::hash::BuildHasherDefault::default());
        let mut by_sheet = vec![Vec::new(); prog.sheets.len().max(1)];
        let mut extent = vec![(0u32, 0u32); prog.sheets.len().max(1)];
        let mut values = Vec::with_capacity(n);
        let mut state = Vec::with_capacity(n);
        let blank = vec![false; n];
        for (i, c) in prog.cells.iter().enumerate() {
            index.insert(c.at, i as u32);
            let s = c.at.sheet as usize;
            by_sheet[s].push(i as u32);
            extent[s].0 = extent[s].0.max(c.at.row + 1);
            extent[s].1 = extent[s].1.max(c.at.col + 1);
            match &c.kind {
                CellKind::Const(v) | CellKind::Input(v) => {
                    values.push(v.clone());
                    state.push(St::Done);
                }
                CellKind::Formula(_) => {
                    values.push(Value::Empty);
                    state.push(St::Pending);
                }
            }
        }
        let seed = opts.rand_seed.unwrap_or_else(|| {
            let nanos = crate::clock::unix_now().as_nanos() as u64;
            if nanos == 0 { 0x9E37_79B9 } else { nanos }
        });
        let coll = Collators::new(&opts.locale.tag);
        if let Some(b) = base {
            for (s, e) in extent.iter_mut().enumerate() {
                let (r, c) = b.extent(s as u16);
                e.0 = e.0.max(r);
                e.1 = e.1.max(c);
            }
        }
        Self {
            prog,
            base,
            opts,
            index,
            by_sheet,
            extent,
            values,
            blank,
            state,
            at: CellRef::new(0, 0, 0),
            rng: seed | 1,
            array_depth: 0,
            range_cache: FastMap::default(),
            full_cache: FastMap::default(),
            lookup_cache: FastMap::default(),
            scratch: Vec::new(),
            unstable_reads: 0,
            coll,
            violations: {
                let _ = crate::notes::take();
                Vec::new()
            },
            logical_fmt: FMT_NONE,
        }
    }

    /// Whether the cell is an input of the program.
    #[must_use]
    pub fn is_input(&self, at: CellRef) -> bool {
        self.index
            .get(&at)
            .is_some_and(|&s| matches!(self.prog.cells[s as usize].kind, CellKind::Input(_)))
    }

    /// Overwrite an input cell. Returns `false` if the cell is not an input.
    pub fn set_input(&mut self, at: CellRef, v: Value) -> bool {
        match self.index.get(&at) {
            Some(&s) if matches!(self.prog.cells[s as usize].kind, CellKind::Input(_)) => {
                self.values[s as usize] = v;
                true
            }
            _ => false,
        }
    }

    /// Evaluate every cell of the program in schedule order.
    pub fn calculate(&mut self) {
        let prog = self.prog;
        let mut cycles = prog.cycles.iter().peekable();
        let mut i = 0u32;
        let n = prog.cells.len() as u32;
        while i < n {
            if let Some(&&(s, e)) = cycles.peek()
                && s == i
            {
                cycles.next();
                self.run_cycle(s, e);
                i = e;
                continue;
            }
            self.ensure(i);
            i += 1;
        }
    }

    fn run_cycle(&mut self, s: u32, e: u32) {
        let Some(it) = self.opts.iteration else {
            // Each member hits a busy cell in the chain and ends as #CIRC!.
            for slot in s..e {
                self.ensure(slot);
            }
            return;
        };
        for slot in s..e {
            self.values[slot as usize] = Value::Number(0.0);
            self.state[slot as usize] = St::Iter;
        }
        for _ in 0..it.max_iterations.max(1) {
            let mut delta = 0.0f64;
            for slot in s..e {
                let old = self.values[slot as usize].clone();
                self.compute(slot);
                delta = delta.max(change(&old, &self.values[slot as usize]));
            }
            if delta < it.max_change {
                break;
            }
        }
        for slot in s..e {
            self.state[slot as usize] = St::Done;
        }
    }

    fn compute(&mut self, slot: u32) {
        let prog = self.prog;
        let def = &prog.cells[slot as usize];
        let CellKind::Formula(expr) = &def.kind else {
            return;
        };
        let prev = std::mem::replace(&mut self.at, def.at);
        let outer = std::mem::replace(&mut self.logical_fmt, FMT_NONE);
        let v = self.eval(expr);
        let (v, blank) = self.finalize(v);
        self.logical_fmt = outer;
        self.at = prev;
        let met = crate::notes::take();
        let v = if met.is_empty() {
            v
        } else {
            for what in met {
                self.violations.push((def.at, what));
            }
            // What LibreOffice gives here is not reproduced: an error value says so.
            Value::Error(ErrorKind::Unsupported)
        };
        self.values[slot as usize] = v;
        self.blank[slot as usize] = blank;
    }

    fn ensure(&mut self, slot: u32) {
        if self.state[slot as usize] == St::Pending {
            self.state[slot as usize] = St::Busy;
            self.compute(slot);
            self.state[slot as usize] = St::Done;
        }
    }

    /// Value of a cell, evaluating it first if it is still pending.
    pub fn cell_value(&mut self, c: CellRef) -> Value {
        match self.index.get(&c) {
            None => self.base.and_then(|b| b.get(c)).unwrap_or(Value::Empty),
            Some(&s) => {
                match self.state[s as usize] {
                    St::Busy => {
                        self.unstable_reads += 1;
                        return Value::Error(ErrorKind::Circ);
                    }
                    St::Iter => self.unstable_reads += 1,
                    _ => {}
                }
                self.ensure(s);
                self.values[s as usize].clone()
            }
        }
    }

    /// The part of `r` inside the sheet's populated extent, as `range_array` sees it.
    #[must_use]
    pub fn trimmed(&self, r: Rect) -> Option<Rect> {
        let (er, ec) = self.extent.get(r.sheet as usize).copied().unwrap_or((0, 0));
        (r.r0 < er && r.c0 < ec)
            .then(|| Rect::new(r.sheet, r.r0, r.c0, r.r1.min(er - 1), r.c1.min(ec - 1)))
    }

    /// Index of the key lane (first column, or first row if `horizontal`) of a
    /// range that is searched many times for exact matches; `None` when the range
    /// is small or its values are not settled yet (a cycle in progress).
    pub fn lookup_index(&mut self, r: Rect, horizontal: bool) -> Option<Arc<LookupIndex>> {
        if r.area() < SMALL_RANGE {
            return None;
        }
        if let Some(i) = self.lookup_cache.get(&(r, horizontal)) {
            return Some(Arc::clone(i));
        }
        let before = self.unstable_reads;
        let arr = self.range_array(r);
        if self.unstable_reads != before {
            return None;
        }
        let lanes = if horizontal { arr.cols } else { arr.rows };
        let mut index = LookupIndex::default();
        for i in 0..lanes {
            let v = if horizontal {
                arr.get(0, i)
            } else {
                arr.get(i, 0)
            };
            index.add(v, i as usize);
        }
        let index = Arc::new(index);
        self.lookup_cache
            .insert((r, horizontal), Arc::clone(&index));
        Some(index)
    }

    /// Materialize a rectangle, trimmed to the sheet's populated extent.
    pub fn range_array(&mut self, r: Rect) -> Arc<Array> {
        // Small ranges are cheaper to rebuild than to remember.
        let cacheable = r.area() >= SMALL_RANGE;
        if cacheable && let Some(a) = self.range_cache.get(&r) {
            return Arc::clone(a);
        }
        let before = self.unstable_reads;
        let a = Arc::new(self.build_range(r));
        if cacheable && self.unstable_reads == before {
            self.range_cache.insert(r, Arc::clone(&a));
        }
        a
    }

    /// A range as an operator or an element-wise function sees it in array evaluation: **all** of
    /// it, the cells beyond the sheet's populated extent included, as empty values. `LibreOffice`
    /// does the same, so `Data!A1:A40000*1` has 40,000 rows of numbers, the empty ones 0, and the
    /// binary search of `MATCH` or the pairs of `CORREL` see them. The rows past the data are not
    /// stored (see [`Ghost`](crate::value::Ghost)): a whole column is a million of them.
    pub fn range_array_full(&mut self, r: Rect) -> Arc<Array> {
        let cacheable = r.area() >= SMALL_RANGE;
        if cacheable && let Some(a) = self.full_cache.get(&r) {
            return Arc::clone(a);
        }
        let before = self.unstable_reads;
        let t = self.range_array(r);
        let (rows, cols) = (r.rows(), r.cols());
        let a = match self.trimmed(r) {
            Some(kept) if kept == r => t,
            Some(kept) => {
                // The columns past the data are stored (empty), the rows past it are not.
                let (srows, tcols) = (kept.rows() as usize, kept.cols() as usize);
                let mut data = vec![Value::Empty; srows * cols as usize];
                for rr in 0..srows {
                    let at = rr * cols as usize;
                    data[at..at + tcols].clone_from_slice(&t.data[rr * tcols..(rr + 1) * tcols]);
                }
                Arc::new(
                    Array::new(srows as u32, cols, data)
                        .with_ghost(rows - srows as u32, Value::Empty),
                )
            }
            None => {
                Arc::new(Array::filled(1, cols, Value::Empty).with_ghost(rows - 1, Value::Empty))
            }
        };
        if cacheable && self.unstable_reads == before {
            self.full_cache.insert(r, Arc::clone(&a));
        }
        a
    }

    fn build_range(&mut self, r: Rect) -> Array {
        let (er, ec) = self.extent.get(r.sheet as usize).copied().unwrap_or((0, 0));
        if r.r0 >= er || r.c0 >= ec {
            return Array::filled(1, 1, Value::Empty);
        }
        let r1 = r.r1.min(er - 1);
        let c1 = r.c1.min(ec - 1);
        let (rows, cols) = (r1 - r.r0 + 1, c1 - r.c0 + 1);
        let mut data = vec![Value::Empty; rows as usize * cols as usize];
        let area = u64::from(rows) * u64::from(cols);
        let slots = &self.by_sheet[r.sheet as usize];
        if self.base.is_some() {
            // Cells live in the program and in the base: read every position.
            for rr in 0..rows {
                for cc in 0..cols {
                    let at = CellRef::new(r.sheet, r.r0 + rr, r.c0 + cc);
                    data[(rr * cols + cc) as usize] = self.cell_value(at);
                }
            }
        } else if area <= slots.len() as u64 * 4 + 64 {
            for rr in 0..rows {
                for cc in 0..cols {
                    let at = CellRef::new(r.sheet, r.r0 + rr, r.c0 + cc);
                    if self.index.contains_key(&at) {
                        data[(rr * cols + cc) as usize] = self.cell_value(at);
                    }
                }
            }
        } else {
            let hits: Vec<u32> = slots
                .iter()
                .copied()
                .filter(|&s| {
                    let a = self.prog.cells[s as usize].at;
                    (r.r0..=r1).contains(&a.row) && (r.c0..=c1).contains(&a.col)
                })
                .collect();
            for s in hits {
                let a = self.prog.cells[s as usize].at;
                data[((a.row - r.r0) * cols + (a.col - r.c0)) as usize] = self.cell_value(a);
            }
        }
        Array::new(rows, cols, data)
    }

    /// Text collator for the locale; `case_sensitive` only where a setting asks.
    #[must_use]
    pub fn collator(&self, case_sensitive: bool) -> &crate::value::Coll {
        self.coll.get(case_sensitive)
    }

    #[must_use]
    pub fn sheet_id(&self, name: &str) -> Option<u16> {
        self.prog
            .sheets
            .iter()
            .position(|s| crate::addr::same_name(s, name))
            .map(|i| i as u16)
    }

    /// Whether the row is hidden in the source workbook.
    #[must_use]
    pub fn is_row_hidden(&self, sheet: u16, row: u32) -> bool {
        self.prog
            .hidden_rows
            .iter()
            .any(|&(s, a, b)| s == sheet && (a..=b).contains(&row))
    }

    /// Whether the cell's formula holds `SUBTOTAL` or `AGGREGATE`.
    #[must_use]
    pub fn is_subtotal_cell(&self, at: CellRef) -> bool {
        self.prog.subtotal_cells.binary_search(&at).is_ok()
    }

    pub fn rand(&mut self) -> f64 {
        // xorshift64*
        self.rng ^= self.rng >> 12;
        self.rng ^= self.rng << 25;
        self.rng ^= self.rng >> 27;
        (self.rng.wrapping_mul(0x2545_F491_4F6C_DD1D) >> 11) as f64 / (1u64 << 53) as f64
    }

    // ----- value access for functions -----

    /// Dereference to a scalar: a single-cell reference reads its cell; a
    /// larger one is intersected with the current cell; an array yields its
    /// top-left element.
    /// The names of the sheets, in order.
    #[must_use]
    pub fn sheet_names(&self) -> &[String] {
        &self.prog.sheets
    }

    /// Whether the formula is being evaluated element by element over arrays (an array formula or
    /// an argument of a function that takes arrays) instead of meeting ranges with its own cell.
    #[must_use]
    pub fn in_array_context(&self) -> bool {
        self.array_depth > 0
    }

    /// The cell of a range that the formula cell meets (implicit intersection, as `LibreOffice`
    /// does it and in any sheet): a range of one column is met by the row of the formula, one of
    /// one row by its column, any other by both. `None` when they do not meet (`#VALUE!`).
    #[must_use]
    pub fn intersection(&self, r: Rect) -> Option<CellRef> {
        let at = self.at;
        let row_in = (r.r0..=r.r1).contains(&at.row);
        let col_in = (r.c0..=r.c1).contains(&at.col);
        if r.cols() == 1 {
            row_in.then(|| CellRef::new(r.sheet, at.row, r.c0))
        } else if r.rows() == 1 {
            col_in.then(|| CellRef::new(r.sheet, r.r0, at.col))
        } else {
            (row_in && col_in).then(|| CellRef::new(r.sheet, at.row, at.col))
        }
    }

    pub fn scalar(&mut self, v: Value) -> Value {
        match v {
            Value::Ref(r) if r.is_single() => self.cell_value(r.top_left()),
            Value::Ref(r) => match self.intersection(r) {
                Some(cell) => self.cell_value(cell),
                None => Value::Error(ErrorKind::Value),
            },
            Value::Array(a) => a.data.first().cloned().unwrap_or_default(),
            v => v,
        }
    }

    /// The value a formula cell holds, and whether it is an *empty result*. A
    /// formula that ends in a blank (`=A9` with `A9` empty) holds the number 0 and
    /// `ISBLANK` is false, but `LibreOffice` keeps an empty result that still equals
    /// `""` and is not a number.
    fn finalize(&mut self, v: Value) -> (Value, bool) {
        // A plain reference to an empty result is an empty result too.
        if let Value::Ref(r) = &v
            && r.is_single()
        {
            let at = r.top_left();
            self.cell_value(at);
            if self.blank_result(at) {
                return (Value::Number(0.0), true);
            }
        }
        // A result that `IF` or `CHOOSE` took from branches of a boolean format is shown as a
        // boolean (`IF({1;0},TRUE())` is TRUE, and `IF({1;0},0,TRUE())` FALSE).
        let logical = matches!(&v, Value::Array(m) if m.logical);
        let from_array = matches!(&v, Value::Array(_));
        let jump = matches!(&v, Value::Array(m) if m.jump_false);
        match self.scalar(v) {
            Value::Empty => (Value::Number(0.0), true),
            // The first element of an array that `IF` left without a branch is `FALSE`.
            v if from_array && v.is_element_path() => (Value::Bool(false), false),
            v if from_array && v.is_empty_path() => {
                // The cell that holds what `IF` chose shows FALSE; an array another function
                // took it into shows an empty text.
                if jump {
                    (Value::Bool(false), false)
                } else {
                    (Value::text(""), false)
                }
            }
            // A boolean format on a number that is not 0 or 1 stays the number: a file holds it
            // so (`5` in an XLSX; an ODS writes `boolean-value="5"`).
            Value::Number(n)
                if (logical || (!from_array && self.logical_fmt != FMT_NONE))
                    && (n == 0.0 || n == 1.0) =>
            {
                (Value::Bool(n == 1.0), false)
            }
            v => (v, false),
        }
    }

    /// Whether the cell is a formula whose result is an empty result (see
    /// `finalize`).
    #[must_use]
    pub fn blank_result(&self, at: CellRef) -> bool {
        self.index.get(&at).is_some_and(|&s| self.blank[s as usize])
    }

    /// [`operand`](Self::operand) for an operator. In a comparison an empty
    /// result (see `finalize`) is empty again; elsewhere (`&`,
    /// arithmetic) `LibreOffice` uses its cached 0.
    fn comparand(&mut self, op: BinOp, v: Value) -> Value {
        let compare = matches!(
            op,
            BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Le | BinOp::Gt | BinOp::Ge
        );
        if compare
            && let Value::Ref(r) = &v
            && r.is_single()
            && self.blank_result(r.top_left())
        {
            return Value::Empty;
        }
        self.operand(v)
    }

    /// Operand of an operator: a scalar, or `Array` where array evaluation is
    /// in force. Elsewhere a multi-cell reference is intersected with the
    /// current cell (legacy implicit intersection, as in `LibreOffice` and
    /// non-dynamic Excel), which is `#VALUE!` when they do not meet.
    pub fn operand(&mut self, v: Value) -> Value {
        match v {
            Value::Ref(r) if r.is_single() => self.cell_value(r.top_left()),
            Value::Ref(r) if self.array_depth > 0 => Value::Array(self.range_array_full(r)),
            v @ Value::Ref(_) => self.scalar(v),
            v => v,
        }
    }

    /// Any value as an array (scalars become 1×1), whatever the context.
    pub fn array_of(&mut self, v: Value) -> Arc<Array> {
        let a = self.array_keep(v);
        if a.ghost.is_some() {
            Arc::new(a.expanded())
        } else {
            a
        }
    }

    /// [`array_of`](Self::array_of) with the rows past the data of an operator's array left unstored.
    pub fn array_keep(&mut self, v: Value) -> Arc<Array> {
        match v {
            Value::Ref(r) if r.is_single() => {
                Arc::new(Array::new(1, 1, vec![self.cell_value(r.top_left())]))
            }
            Value::Ref(r) => self.range_array(r),
            Value::Array(a) => a,
            s => Arc::new(Array::new(1, 1, vec![s])),
        }
    }

    /// # Errors
    ///
    /// The error value Excel gives when `v` is not a number, or the error `v` holds.
    pub fn num(&mut self, v: &Value) -> Res<f64> {
        let s = self.scalar(v.clone());
        coerce::to_num(&s, &self.opts.locale, self.opts.date_ctx())
    }

    /// # Errors
    ///
    /// The error `v` holds.
    pub fn text(&mut self, v: &Value) -> Res<String> {
        let s = self.scalar(v.clone());
        coerce::to_text(&s, &self.opts.locale)
    }

    /// # Errors
    ///
    /// The error value Excel gives when `v` is not a boolean, or the error `v` holds.
    pub fn boolean(&mut self, v: &Value) -> Res<bool> {
        let s = self.scalar(v.clone());
        coerce::to_bool(&s, &self.opts.locale)
    }

    /// Flatten arguments for aggregate functions. Numbers found in references
    /// and arrays are collected; text, booleans and blanks there are skipped.
    /// Direct scalar arguments are coerced. Errors propagate.
    ///
    /// # Errors
    ///
    /// The first error value among `args`, or `#VALUE!` for an argument that is not a number.
    pub fn numbers(&mut self, args: &[Value]) -> Res<Vec<f64>> {
        let mut out = Vec::new();
        for a in args {
            match a {
                // A small range is read into a reusable buffer instead of a shared array: every
                // cell is read before the first error returns, as when the array is built.
                Value::Ref(r) if r.area() < SMALL_RANGE => {
                    let mut buf = std::mem::take(&mut self.scratch);
                    buf.clear();
                    if let Some(t) = self.trimmed(*r) {
                        for row in t.r0..=t.r1 {
                            for col in t.c0..=t.c1 {
                                buf.push(self.cell_value(CellRef::new(t.sheet, row, col)));
                            }
                        }
                    }
                    let mut res = Ok(());
                    for v in &buf {
                        match v {
                            Value::Number(n) => out.push(*n),
                            Value::Error(e) => {
                                res = Err(*e);
                                break;
                            }
                            _ => {}
                        }
                    }
                    self.scratch = buf;
                    res?;
                }
                Value::Ref(_) | Value::Array(_) => {
                    let array = matches!(a, Value::Array(_));
                    for v in &self.array_of(a.clone()).data {
                        match v {
                            Value::Number(n) => out.push(*n),
                            // The elements of an array are numbers (`SUM({1,2}>1)` is 1).
                            Value::Bool(b) if array => out.push(f64::from(u8::from(*b))),
                            Value::Error(e) => return Err(*e),
                            _ => {}
                        }
                    }
                }
                Value::Empty => {}
                // LibreOffice does not coerce text given directly to SUM & co.
                Value::Text(_) => {
                    return Err(ErrorKind::Value);
                }
                v => out.push(coerce::to_num(v, &self.opts.locale, self.opts.date_ctx())?),
            }
        }
        Ok(out)
    }

    // ----- evaluation -----

    /// The arguments of a function that takes arrays, as an eager one registered with
    /// `force_array` has them.
    pub fn eval_array_args(&mut self, args: &[Expr]) -> Vec<Value> {
        self.array_depth += 1;
        let vals = args.iter().map(|a| self.eval(a)).collect();
        self.array_depth -= 1;
        vals
    }

    pub fn eval(&mut self, e: &Expr) -> Value {
        match e {
            Expr::Const(v) => v.clone(),
            Expr::Cell(c) => Value::Ref(Rect::single(*c)),
            Expr::Range(r) => Value::Ref(*r),
            Expr::Missing => Value::Empty,
            Expr::Unknown(..) => Value::Error(ErrorKind::Name),
            // `+x` hands `x` over untouched, an empty result included.
            Expr::Unary(UnOp::Plus, a) => self.eval(a),
            Expr::Unary(op, a) => {
                let v = self.eval(a);
                // A minus keeps the boolean format of an operand that is a boolean (`-TRUE()`), not
                // that of a number that kept it (`-VALUE(TRUE())`).
                let boolean = matches!(op, UnOp::Neg) && self.is_boolean(&v);
                let v = self.operand(v);
                let r = self.map_array(v, |eng, x| eng.unary(*op, x));
                self.logical_fmt = if boolean { FMT_NEGATED } else { FMT_NONE };
                r
            }
            Expr::Binary(op, a, b) => {
                let l = self.eval(a);
                let l = self.comparand(*op, l);
                let r = self.eval(b);
                let r = self.comparand(*op, r);
                let v = self.binary(*op, l, r);
                self.logical_fmt = FMT_NONE;
                v
            }
            Expr::Call(id, args) => {
                let def = id.def();
                let res = match def.imp {
                    crate::functions::Imp::Lazy(f) => {
                        if crate::functions::arity_refused(def, args.len()) {
                            Err(crate::functions::arity_error(def, args.len()))
                        } else {
                            let r = f(self, args);
                            // What `IF` and the like give is what one of their arguments gave,
                            // format included: they set it themselves.
                            if def.keep_format == 0 {
                                self.logical_fmt = FMT_NONE;
                            }
                            r
                        }
                    }
                    crate::functions::Imp::Eager(f) => {
                        if crate::functions::arity_refused(def, args.len()) {
                            Err(crate::functions::arity_error(def, args.len()))
                        } else {
                            let force = def.force_array;
                            let vals: Vec<Value> = args
                                .iter()
                                .enumerate()
                                .map(|(i, a)| {
                                    let forced = u32::from(force >> i.min(31) & 1 == 1);
                                    self.array_depth += forced;
                                    let v = self.eval(a);
                                    self.array_depth -= forced;
                                    v
                                })
                                .collect();
                            let vals = if def.ghost_ok {
                                vals
                            } else {
                                Self::unghost(vals)
                            };
                            // An error a calculation gave in an argument is the answer (the first one)
                            // for the functions that read the errors of the arguments before checking
                            // them (`SUMIF(NA(),1)`, `ACCRINT("a",…,NA())`).
                            if let Some(kind) = crate::functions::propagated_error(def, &vals) {
                                self.logical_fmt = FMT_NONE;
                                return Value::Error(kind);
                            }
                            // With an array among the arguments the formula is evaluated as a matrix
                            // and a plain value is a matrix of one element, which any argument takes.
                            let matrix_mode = vals.iter().any(|v| matches!(v, Value::Array(_)));
                            if !matrix_mode
                                && let Some(kind) = crate::functions::scalar_refused(def, &vals)
                            {
                                return Value::Error(kind);
                            }
                            if self.array_depth == 0
                                && let Some(kind) = crate::functions::range_refused(def, &vals)
                            {
                                return Value::Error(kind);
                            }
                            let mask = def.lift_mask;
                            // An array where the function takes one value is taken element by
                            // element (a function called for each, giving an array), in a formula that
                            // is not an array formula too; a range only in array evaluation, as it is
                            // met by the cell of the formula otherwise.
                            // An array among the arguments turns the ranges into arrays too.
                            let takes = |i: usize| mask == 0 || mask >> i.min(31) & 1 == 1;
                            // What `IF` or `CHOOSE` chose element by element is not taken element by
                            // element outside array evaluation: the function gets its first element.
                            let vals = if def.lift && self.array_depth == 0 {
                                vals.into_iter()
                                    .enumerate()
                                    .map(|(i, v)| match v {
                                        Value::Array(a) if a.jump && takes(i) => {
                                            Self::read_element(
                                                def,
                                                i,
                                                Self::element_of_a_call(a.logical(0, 0)),
                                            )
                                        }
                                        v => v,
                                    })
                                    .collect()
                            } else {
                                vals
                            };
                            let has_array = vals
                                .iter()
                                .enumerate()
                                .any(|(i, v)| takes(i) && matches!(v, Value::Array(_)));
                            let in_array = self.array_depth > 0;
                            let spread = def.lift
                                && (has_array
                                    || (in_array
                                        && vals.iter().enumerate().any(|(i, v)| {
                                            takes(i) && matches!(v, Value::Ref(r) if !r.is_single())
                                        })));
                            if spread {
                                let r = self.lifted(def, f, &vals, mask);
                                self.logical_fmt = FMT_NONE;
                                r
                            } else {
                                let result = f(self, &vals);
                                // A function that sets no number format leaves the boolean one of
                                // its argument on the number it gives (`VALUE(TRUE())`).
                                // The format the arguments left is the one of the last of them.
                                let after_args = self.logical_fmt;
                                self.logical_fmt = FMT_NONE;
                                if def.keep_format != 0
                                    && matches!(result, Ok(Value::Number(_)))
                                    && (after_args == FMT_KEPT
                                        || self.keeps_boolean(def.keep_format, &vals))
                                {
                                    self.logical_fmt = FMT_KEPT;
                                }
                                // An error handed to a function comes back as it is: a `#VALUE!`
                                // must not take the code of the domain of the function.
                                if matches!(result, Err(ErrorKind::Value))
                                    && vals
                                        .iter()
                                        .any(|v| matches!(v, Value::Error(ErrorKind::Value)))
                                {
                                    return Value::Error(ErrorKind::Value);
                                }
                                result
                            }
                        }
                    }
                };
                res.unwrap_or_else(|k| Self::error_value(k, def.name))
            }
        }
    }

    /// The error a function failure becomes: an argument outside a function's
    /// domain is `#NUM!` in Excel and `Err:502` in `LibreOffice` (an error that keeps its code).
    fn error_value(k: ErrorKind, function: &str) -> Value {
        use crate::functions::{LO_DOMAIN_NUM, LoError, lo_error};
        let domain = k == ErrorKind::Domain;
        Value::Error(match k {
            ErrorKind::Domain | ErrorKind::Value => match lo_error(function) {
                Some(LoError::Libre(code)) => ErrorKind::Libre(code),
                Some(LoError::Value) => ErrorKind::Value,
                Some(LoError::Num) => ErrorKind::Num,
                Some(LoError::DomainIs(code)) if domain => ErrorKind::Libre(code),
                _ if !domain => ErrorKind::Value,
                _ if LO_DOMAIN_NUM.contains(&function) => ErrorKind::Num,
                _ => ErrorKind::Libre(502),
            },
            k => k,
        })
    }

    /// Apply a scalar function element-wise, broadcasting like operators do.
    #[allow(clippy::unnecessary_wraps)] // same fallible signature as its siblings, which share call sites
    fn lifted(
        &mut self,
        def: &crate::functions::FuncDef,
        f: crate::functions::EagerFn,
        vals: &[Value],
        mask: u32,
    ) -> Res<Value> {
        let mut arrays: Vec<Option<Arc<Array>>> = vals
            .iter()
            .enumerate()
            .map(|(i, v)| match v {
                _ if mask != 0 && mask >> i & 1 == 0 => None,
                Value::Array(_) => Some(self.array_keep(v.clone())),
                Value::Ref(r) if !r.is_single() => Some(self.range_array_full(*r)),
                _ => None,
            })
            .collect();
        let ghost = Self::ghost_shape(arrays.iter().flatten().map(|a| &**a));
        if ghost.is_none() {
            for a in arrays.iter_mut().flatten() {
                if a.ghost.is_some() {
                    *a = Arc::new(a.expanded());
                }
            }
        }
        let rows = match ghost {
            Some((stored, _)) => stored,
            None => arrays.iter().flatten().map(|a| a.rows).max().unwrap_or(1),
        };
        let cols = arrays.iter().flatten().map(|a| a.cols).max().unwrap_or(1);
        let mut data = Vec::with_capacity(rows as usize * cols as usize);
        for i in 0..rows {
            for j in 0..cols {
                let call: Vec<Value> = vals
                    .iter()
                    .zip(&arrays)
                    .enumerate()
                    .map(|(k, (v, a))| match a {
                        // An array that is smaller than the others has no element there: `#N/A`.
                        Some(m) => {
                            let r = if m.logical_rows() == 1 { 0 } else { i };
                            let c = if m.cols == 1 { 0 } else { j };
                            if u64::from(r) < m.logical_rows() && c < m.cols {
                                match Self::element_of_a_call(m.logical(r, c)) {
                                    // The elements of an array are numbers; a boolean is one only in a cell.
                                    Value::Bool(b) if matches!(v, Value::Array(_)) => {
                                        Value::Number(f64::from(u8::from(b)))
                                    }
                                    x => Self::read_element(def, k, x),
                                }
                            } else {
                                Value::Error(ErrorKind::Na)
                            }
                        }
                        None => v.clone(),
                    })
                    .collect();
                let r = f(self, &call).unwrap_or_else(|k| Self::error_value(k, ""));
                let r = self.scalar(r);
                // The element a function hands back as it got it (`IFERROR`) is the one it was.
                data.push(if r.is_element_path() {
                    Value::empty_path()
                } else {
                    r
                });
            }
        }
        let mut out = Array::new(rows, cols, data);
        if let Some((_, extra)) = ghost {
            // The rows past the data all hold what the function gives for the value there.
            let call: Vec<Value> = vals
                .iter()
                .zip(&arrays)
                .map(|(v, a)| match a {
                    Some(m) => m
                        .ghost
                        .as_ref()
                        .map_or_else(|| m.get(0, 0).clone(), |g| g.value.clone()),
                    None => v.clone(),
                })
                .collect();
            let r = f(self, &call).unwrap_or_else(|k| Self::error_value(k, ""));
            let tail = self.scalar(r);
            out = out.with_ghost(extra, tail);
        }
        Ok(Value::Array(Arc::new(out)))
    }

    /// An element of an array as an operator or a function applied to each element reads it: the
    /// element that the missing branch of an `IF` gave is [`Value::element_path`] there (it is an
    /// empty text when it is taken out with `INDEX`, see [`Value::empty_path`]).
    pub(crate) fn element_of_a_call(x: &Value) -> Value {
        if x.is_empty_path() {
            Value::element_path()
        } else {
            x.clone()
        }
    }

    /// The element a missing `else` gave, as the argument `i` of `def` reads it: the number 0, the
    /// text `FALSE` or an empty text (`marker_table.rs`). Any other value as it is.
    fn read_element(def: &crate::functions::FuncDef, i: usize, x: Value) -> Value {
        if !x.is_element_path() {
            return x;
        }
        let bit = 1u32 << i.min(31);
        if def.marker_empty & bit != 0 {
            Value::text("")
        } else if def.marker_text & bit != 0 {
            Value::text("FALSE")
        } else {
            Value::Number(0.0)
        }
    }

    /// The arguments with the rows past the data of every array stored.
    fn unghost(vals: Vec<Value>) -> Vec<Value> {
        vals.into_iter()
            .map(|v| match v {
                Value::Array(a) if a.ghost.is_some() => Value::Array(Arc::new(a.expanded())),
                v => v,
            })
            .collect()
    }

    /// How an operation over `arrays` can keep the rows past their data unstored: every array
    /// is a single value or has such rows, and they have the same number of them. `(rows stored in
    /// the result, rows not stored)`; `None` when no array has them or they cannot.
    fn ghost_shape<'a>(arrays: impl Iterator<Item = &'a Array>) -> Option<(u32, u32)> {
        let mut logical: Option<u64> = None;
        let mut stored = 1u32;
        for a in arrays {
            if a.ghost.is_some() {
                let l = a.logical_rows();
                if logical.is_some_and(|x| x != l) {
                    return None;
                }
                logical = Some(l);
                stored = stored.max(a.rows);
            } else if !(a.rows == 1 && a.cols == 1) {
                return None;
            }
        }
        let l = logical?;
        Some((stored, (l - u64::from(stored)) as u32))
    }

    fn map_array(&mut self, v: Value, f: impl Fn(&mut Self, Value) -> Value + Copy) -> Value {
        match v {
            Value::Array(a) => {
                let data = a
                    .data
                    .iter()
                    .map(|x| f(self, Self::element_of_a_call(x)))
                    .collect();
                let mut out = Array::new(a.rows, a.cols, data).with_jump(a.jump);
                if let Some(g) = &a.ghost {
                    out = out.with_ghost(g.rows, f(self, g.value.clone()));
                }
                Value::Array(Arc::new(out))
            }
            v => f(self, v),
        }
    }

    fn unary(&mut self, op: UnOp, v: Value) -> Value {
        match op {
            UnOp::Plus => v,
            UnOp::Neg | UnOp::Percent => match self.num(&v) {
                Ok(n) => Value::Number(if op == UnOp::Neg { -n } else { n / 100.0 }),
                Err(e) => Value::Error(e),
            },
        }
    }

    fn binary(&mut self, op: BinOp, l: Value, r: Value) -> Value {
        match (&l, &r) {
            (Value::Array(_), _) | (_, Value::Array(_)) => {
                let both_arrays = matches!((&l, &r), (Value::Array(_), Value::Array(_)));
                let (mut a, mut b) = (self.array_keep(l), self.array_keep(r));
                let ghost = Self::ghost_shape([&*a, &*b].into_iter());
                if ghost.is_none() {
                    for m in [&mut a, &mut b] {
                        if m.ghost.is_some() {
                            *m = Arc::new(m.expanded());
                        }
                    }
                }
                let rows = match ghost {
                    Some((stored, _)) => stored,
                    None => a.rows.max(b.rows),
                };
                let cols = a.cols.max(b.cols);
                let mut data = Vec::with_capacity(rows as usize * cols as usize);
                for i in 0..rows {
                    for j in 0..cols {
                        let pick = |m: &Array| {
                            let lr = m.logical_rows();
                            let ri = if lr == 1 { 0 } else { i };
                            let ci = if m.cols == 1 { 0 } else { j };
                            (u64::from(ri) < lr && ci < m.cols)
                                .then(|| Self::element_of_a_call(m.logical(ri, ci)))
                        };
                        data.push(match (pick(&a), pick(&b)) {
                            // Between two arrays the element a missing `else` gave is an empty
                            // text in a concatenation (`IF({0;1},1)&IF({1;1},"x")` is `x`).
                            (Some(x), Some(y)) if op == BinOp::Concat && both_arrays => {
                                let empty = |v: Value| {
                                    if v.is_element_path() {
                                        Value::text("")
                                    } else {
                                        v
                                    }
                                };
                                self.scalar_binary(op, &empty(x), &empty(y))
                            }
                            (Some(x), Some(y)) => self.scalar_binary(op, &x, &y),
                            _ => Value::Error(ErrorKind::Na),
                        });
                    }
                }
                let mut out = Array::new(rows, cols, data).with_jump(a.jump || b.jump);
                if let Some((_, extra)) = ghost {
                    let tail = |m: &Array| {
                        m.ghost
                            .as_ref()
                            .map_or_else(|| m.get(0, 0).clone(), |g| g.value.clone())
                    };
                    let value = self.scalar_binary(op, &tail(&a), &tail(&b));
                    out = out.with_ghost(extra, value);
                }
                Value::Array(Arc::new(out))
            }
            _ => self.scalar_binary(op, &l, &r),
        }
    }

    /// Whether the value is a boolean: one, or a reference to a cell that holds one (a cell that
    /// `LibreOffice` shows as `TRUE` or `FALSE`).
    fn is_boolean(&mut self, v: &Value) -> bool {
        match v {
            Value::Bool(_) => true,
            Value::Ref(r) if r.is_single() => {
                matches!(self.cell_value(r.top_left()), Value::Bool(_))
            }
            _ => false,
        }
    }

    /// Whether one of the arguments of the positions in `mask` is a boolean.
    fn keeps_boolean(&mut self, mask: u32, vals: &[Value]) -> bool {
        vals.iter()
            .enumerate()
            .any(|(i, v)| mask >> i.min(31) & 1 == 1 && self.is_boolean(v))
    }

    fn scalar_binary(&self, op: BinOp, a: &Value, b: &Value) -> Value {
        if let Value::Error(e) = a {
            return Value::Error(*e);
        }
        if let Value::Error(e) = b {
            return Value::Error(*e);
        }
        let loc = &self.opts.locale;
        let sys = self.opts.date_ctx();
        // LibreOffice has no separate boolean type: TRUE is 1 and FALSE is 0.
        let as_num = |v: &Value| match v {
            Value::Bool(b) => Value::Number(f64::from(u8::from(*b))),
            v => v.clone(),
        };
        // In a comparison the element a missing `else` gave is empty (`=""`, `=0`, `=FALSE()` are true).
        let empty = Value::Empty;
        let (ca, cb) = (
            if a.is_element_path() { &empty } else { a },
            if b.is_element_path() { &empty } else { b },
        );
        match op {
            BinOp::Concat => match (coerce::to_text(a, loc), coerce::to_text(b, loc)) {
                (Ok(x), Ok(y)) => Value::text(x + &y),
                (Err(e), _) | (_, Err(e)) => Value::Error(e),
            },
            BinOp::Eq | BinOp::Ne if matches!((ca, cb), (Value::Text(_), Value::Text(_))) => {
                // Equality of two texts is not the order of the collation (see `casemap`).
                let (Value::Text(x), Value::Text(y)) = (ca, cb) else {
                    unreachable!()
                };
                let eq = crate::casemap::text_eq(x, y, self.opts.search.case_sensitive);
                Value::Bool(eq == (op == BinOp::Eq))
            }
            BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Le | BinOp::Gt | BinOp::Ge => {
                let ord = compare(
                    &as_num(ca),
                    &as_num(cb),
                    self.coll.get(self.opts.search.case_sensitive),
                );
                Value::Bool(match op {
                    BinOp::Eq => ord == Ordering::Equal,
                    BinOp::Ne => ord != Ordering::Equal,
                    BinOp::Lt => ord == Ordering::Less,
                    BinOp::Le => ord != Ordering::Greater,
                    BinOp::Gt => ord == Ordering::Greater,
                    _ => ord != Ordering::Less,
                })
            }
            _ => {
                let (x, y) = match (coerce::to_num(a, loc, sys), coerce::to_num(b, loc, sys)) {
                    (Ok(x), Ok(y)) => (x, y),
                    (Err(e), _) | (_, Err(e)) => return Value::Error(e),
                };
                let r = match op {
                    BinOp::Add => x + y,
                    BinOp::Sub => x - y,
                    BinOp::Mul => x * y,
                    BinOp::Div => {
                        if y == 0.0 {
                            return Value::Error(ErrorKind::DivZero);
                        }
                        x / y
                    }
                    _ => match power(x, y) {
                        Some(r) => r,
                        None => return Value::Error(ErrorKind::Num),
                    },
                };
                if r.is_finite() {
                    Value::Number(r)
                } else {
                    Value::Error(ErrorKind::Num)
                }
            }
        }
    }

    /// Final value of a cell by position (for reading outputs).
    pub fn get(&mut self, c: CellRef) -> Value {
        self.cell_value(c)
    }
}

/// `x^y` as `LibreOffice` computes it: `0^0` is 1 and `x^(1/n)` is the real root for odd `n`.
#[must_use]
pub fn power(x: f64, y: f64) -> Option<f64> {
    if x == 0.0 && y == 0.0 {
        return Some(1.0);
    }
    if x < 0.0 && y != 0.0 && y.fract() != 0.0 {
        let n = (1.0 / y).round();
        if (1.0 / y - n).abs() < 1e-12 && n % 2.0 != 0.0 {
            return Some(-(-x).powf(y));
        }
    }
    let r = x.powf(y);
    r.is_finite().then_some(r)
}

fn change(old: &Value, new: &Value) -> f64 {
    match (old, new) {
        (Value::Number(a), Value::Number(b)) => (a - b).abs(),
        (a, b) if a == b => 0.0,
        _ => f64::INFINITY,
    }
}
