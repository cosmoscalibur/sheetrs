//! The compiler: [`Model`] + input/output specification → [`Program`].
//!
//! Pipeline:
//!
//! 1. Parse every formula to IR and resolve statically-known `INDIRECT` /
//!    `OFFSET` targets into plain references.
//! 2. Build the dependency graph lazily from the requested outputs backwards,
//!    cutting at input cells. Anything outside that cone is dead and dropped.
//! 3. Order the cone with Tarjan SCC (dependencies first, cycles isolated).
//! 4. Mark cells that depend on an input, a volatile function or a dynamic
//!    reference as *live*; evaluate everything else now with the very same
//!    engine and keep only the resulting constants. What remains is the
//!    minimal run-time computation.

use crate::error::{Error, Result};
use crate::formula::{self, ParseCtx};
use crate::graph::{self, Csr};
use crate::model::{Model, ModelCell};
use sheetrs_engine::addr::{self, CellRef, Rect};
use sheetrs_engine::calculate::Overrides;
use sheetrs_engine::engine::{ConstSource, Engine};
use sheetrs_engine::fastmap::{FastMap, FastSet};
use sheetrs_engine::ir::{
    Binding, BuildInfo, BuildOptions, CellDef, CellKind, Expr, Meta, Mode, Program, Queryable,
};
use sheetrs_engine::options::{DateSystem, EngineOptions, Iteration, Locale, Search};
use sheetrs_engine::value::Value;
use std::borrow::Cow;
use std::collections::{BTreeSet, HashSet};

/// A cell or range declared as input or output.
#[derive(Debug, Clone)]
pub struct BindingSpec {
    /// `Sheet1!A1`, `Sheet1!A1:B3`, `A1` (first sheet) or a defined name.
    pub target: String,
    /// The JSON key; the canonical range text (`Sheet1!B2:C9`) when absent.
    pub name: Option<String>,
}

impl BindingSpec {
    #[must_use]
    pub fn parse(s: &str) -> Self {
        Self {
            target: s.trim().to_string(),
            name: None,
        }
    }

    /// A declared range, with the alias that names it in a request, if it has one.
    #[must_use]
    pub fn declared(name: Option<&str>, target: &str) -> Self {
        match name {
            Some(n) => Self::named(n, target),
            None => Self::parse(target),
        }
    }

    #[must_use]
    pub fn named(name: &str, target: &str) -> Self {
        Self {
            target: target.trim().to_string(),
            name: Some(name.trim().to_string()),
        }
    }
}

#[derive(Debug, Clone)]
#[allow(clippy::struct_excessive_bools)] // the independent options of a build, each with its own meaning
pub struct Spec {
    /// Declared inputs. Empty: no input is declared, and they are deduced (the constants the
    /// outputs start from, once the workbook is pruned to what the outputs read).
    pub inputs: Vec<BindingSpec>,
    /// The inputs were declared, even if the list is empty: there are none, so everything that
    /// does not read one is evaluated at build time.
    pub inputs_declared: bool,
    /// Declared outputs. Empty: none is declared and they are deduced, the same in every mode: the
    /// cells that are calculated (the formulas, not the constants), as the rectangles that hold them.
    pub outputs: Vec<BindingSpec>,
    /// Defaults to the date system of the source workbook.
    pub date_system: Option<DateSystem>,
    pub iteration: Option<Iteration>,
    pub locale: Option<Locale>,
    pub search: Option<Search>,
    pub now: Option<f64>,
    pub rand_seed: Option<u64>,
    /// What `INFO` and `CELL("filename")` say; the file name is the workbook's unless given.
    pub host: sheetrs_engine::options::Host,
    /// Evaluate input-independent cells at build time (default true). A Rust-library option for
    /// tests and low-level use (comparing a folded and an unfolded program): `sheetc`, `sheetc.toml`
    /// and the bindings always fold.
    pub fold: bool,
    /// What the program is for. A recalculated workbook (`Recalculate`) has every formula for an
    /// output and any constant for an input, whatever `inputs` and `outputs` say. An embedded one
    /// (`Embedded`) is any program with the workbook inside: it takes them as they are declared, and
    /// deduces what is not.
    pub mode: Mode,
    /// Settings-sheet rows, carried into the program for `--meta`.
    pub meta: Vec<Meta>,
    /// The settings sheet: not part of the computation's outputs or queries.
    pub skip_sheet: Option<u16>,
    /// Only inspecting (`sheetc inspect`): functions the engine lacks are reported, not an error.
    pub inspect: bool,
    /// Added to the program's diagnostics (for example, about the locale).
    pub notes: Vec<String>,
    /// Compatibility features (`enable_features`); every one is off unless enabled.
    pub features: Features,
    /// What the passes before a recalculation changed (see [`Cache`]): the formulas those changes
    /// reach, and the ones that depend on them, are calculated, and every other formula keeps the result
    /// the workbook caches. Only `Mode::Recalculate` reads it; the default is nothing changed.
    pub cache: Cache,
    /// Emulates `calculateAll()` of `LibreOffice`: the whole workbook is calculated first, the constants
    /// the build folds are the results of that calculation and those results are what the program
    /// stores as cache. Off by default: the values the workbook stores are the folded constants and the
    /// cache of the formulas the build keeps (the load of `LibreOffice`, which does not calculate what a
    /// file already has a result for). In `Mode::Recalculate` it ignores [`cache`](Self::cache) and
    /// calculates every formula.
    pub calculate_all: bool,
    /// Build only: the program stores no cached result for the formulas it keeps, so that it calculates
    /// all of them in every run (a lighter program). Without it a run calculates only what the inputs it
    /// is given change.
    pub force_calculate_all: bool,
    /// Only these cells are asked for: the outputs of the program are the formulas inside them (the
    /// program is pruned to what they read, and only they are written), where none is given the outputs
    /// are every formula. See [`Only`].
    pub only: Vec<Only>,
    /// A calculation whose answer is the workbook (`sheetc calculate` sets it, and only it: a build, embedded or not, keeps its
    /// own rules, whatever calls it): besides the outputs it declares (the names a request may use, and what an answer lists
    /// by default), every formula the calculation reaches is calculated and written, whether a declared output holds it or not,
    /// and the formulas it does not reach stay as the template caches them (the rules of the cache: what the inputs given and the
    /// formulas with no result or that are volatile reach, and what depends on them).
    pub whole_workbook: bool,
    /// The inputs a calculation was given, as the request names them (a declared name, a range or a defined name): they are
    /// the only cells that can differ from what the workbook holds, so they are the ones whose cone is calculated; the rest of
    /// the declared inputs keep their value and what reads only them keeps its result. `None` for a build, whose program is run
    /// later with any of the declared inputs.
    pub given_inputs: Option<Vec<String>>,
}

/// The cells of a sheet that a calculation is asked for: the whole sheet, some columns of it or some rows
/// (0-based, inclusive), as the conversion of formulas to values names them. A sheet that the workbook
/// has not is no cell.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Only {
    pub sheet: String,
    pub rows: Option<(u32, u32)>,
    pub columns: Option<(u32, u32)>,
}

/// What a workbook can depend on that the engine cannot calculate: the three kinds the build tells apart (see
/// [`FromCacheNote`]). `LibreOffice` keeps the result the file caches for a cell that depends on them until
/// something dirties the cell, and the build does the same (`sheetrs-compiler/docs/FIDELITY.md`, *What depends
/// on what the engine cannot calculate*).
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum Unsupported {
    /// A reference to another workbook (`[1]Sheet1!A1`, or a defined name that holds one): the result the file
    /// caches is the stored value of the link, which `LibreOffice` shows without updating it.
    External,
    /// A function that a macro of the workbook may define (a name no spreadsheet program has, in a workbook
    /// that carries macros): the macros are never run.
    Macro,
    /// A function the engine has no implementation of: an add-in or extension (`_xll.NAME`, an add-in
    /// service) or a function of `LibreOffice` that it lacks.
    Extension,
}

impl Unsupported {
    /// The name of the kind: `external`, `macro` or `extension`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::External => "external",
            Self::Macro => "macro",
            Self::Extension => "extension",
        }
    }
}

/// What a build left as the file caches it for a kind of [`Unsupported`]: `cells` formulas (and what
/// depends on them is calculated from those values), the first at `first` (`Sheet1!A1`), and what they
/// depend on.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FromCacheNote {
    pub kind: Unsupported,
    pub cells: usize,
    pub first: String,
    /// The references or the functions, without repeats, as the formulas write them (at most five).
    pub what: Vec<String>,
}

/// What a recalculation that uses the results a workbook caches is told, besides the workbook: the
/// changes the passes before it made that its formulas cannot show. Only the formulas those changes
/// reach, and the ones that depend on them, are calculated; every other formula is a constant that
/// holds the result the workbook caches, and keeps it.
///
/// Which formulas a change reaches is what `LibreOffice` recalculates when it makes the same change and the
/// file is saved and opened again (`sheetrs-edit/tests/assets/write_to`, the fixture `dirty`):
///
/// * a deletion of rows or columns reaches a formula that reads a deleted cell or a range that shrank (the
///   writer leaves such a formula with no result), the whole columns of a sheet that lost rows and the whole rows of
///   one that lost columns, `ROW` and `COLUMN` (of the formula's own cell or of a reference that moved), and the
///   formulas that read through a name whose target moved or shrank ([`names_reached`]); a formula that only reads
///   cells that moved keeps its result;
/// * a deletion of a sheet reaches every formula, and a deletion of a name every formula that reads through a name;
/// * a formula that was already stale (an error `#REF!` the file holds, a result of the file that is not the
///   formula's) keeps its result unless one of the above reaches it.
///
/// A formula with no cached result, a volatile one, and one that reads the cells of a change are calculated,
/// and so is everything that depends on them. A workbook that reads cells at run time (`INDIRECT`, `OFFSET`)
/// can read any of them: it is calculated whole. However many formulas are stale, the others keep their result
/// (the walk that finds them costs 1-1.3 s on 2.4 M formulas, `docs/performance.md`).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Cache {
    /// Cells whose content changed, as `(sheet, first row, first column, last row, last column)`
    /// (0-based, inclusive): the formulas that read them are stale.
    pub changed: Vec<(String, u32, u32, u32, u32)>,
    /// The rows whose visibility changed, as `(sheet, first, last)` (0-based, inclusive). `LibreOffice` calculates a `SUBTOTAL`
    /// that reads a range with one of them, and a formula that reads such a range and sits in a group of formulas (adjacent
    /// cells of a column with the same formula, as relative references).
    pub rows_hidden: Vec<(String, u32, u32)>,
    /// The bands of rows that were deleted, as `(sheet, first, last)` (0-based, inclusive, in the rows of the
    /// workbook before the deletion): a formula that reads whole columns of the sheet (a range to the last
    /// row, which keeps its text when rows go) is stale, and `ROW` and `COLUMN` of what moved.
    pub rows_deleted: Vec<(String, u32, u32)>,
    /// The bands of columns that were deleted (as `rows_deleted`): a formula that reads whole rows of the
    /// sheet is stale.
    pub columns_deleted: Vec<(String, u32, u32)>,
    /// A sheet was deleted: `LibreOffice` calculates every formula of the workbook again.
    pub sheets_deleted: bool,
    /// The defined names (in any case) that a formula reads through to be stale: those whose target the deletion moved or
    /// shrank ([`names_reached`]).
    pub stale_names: Vec<String>,
    /// A name was deleted: every formula that reads through a name is stale (the names are those of the workbook), and in
    /// an XLSX, `LibreOffice` also calculates again the formulas that hold a bare error `#REF!` (it reads it as a name it
    /// cannot resolve).
    pub names_deleted: bool,
}

/// One reference of the definition of a name: the sheet it names (`None`: it names none), and its rectangle
/// (0-based, inclusive; the whole columns and rows reach the last row and column).
type NameRef = (Option<String>, u32, u32, u32, u32);

/// What the definition of a name reads: its own references and the names it uses.
fn name_reads(definition: &str) -> (Vec<NameRef>, Vec<String>) {
    use sheetrs_formula::ast::RefArea;
    let Ok(ast) = sheetrs_formula::parse(definition.trim_start_matches('=')) else {
        return (Vec::new(), Vec::new());
    };
    let (mut refs, mut names) = (Vec::new(), Vec::new());
    ast.walk(&mut |x| match x {
        sheetrs_formula::Expr::Ref(r) => {
            let (r0, c0, r1, c1) = match r.area {
                RefArea::Cell(a) => (a.row, a.col, a.row, a.col),
                RefArea::Range(a, b) => (
                    a.row.min(b.row),
                    a.col.min(b.col),
                    a.row.max(b.row),
                    a.col.max(b.col),
                ),
                RefArea::Cols(a, b) => (0, a.col.min(b.col), addr::MAX_ROWS - 1, a.col.max(b.col)),
                RefArea::Rows(a, b) => (a.row.min(b.row), 0, a.row.max(b.row), addr::MAX_COLS - 1),
            };
            refs.push((r.sheet.clone(), r0, c0, r1, c1));
        }
        sheetrs_formula::Expr::Name(n) => names.push(n.to_lowercase()),
        _ => {}
    });
    (refs, names)
}

/// The defined names that `LibreOffice` recalculates the formulas of after the deletion of the bands of rows
/// (`rows`) and of columns (`columns`), as `(sheet, first, last)` in the coordinates before the deletion.
/// `names` are the definitions before it (name to formula text). Returned in lower case.
///
/// Observed (the file saved and opened again, `sheetrs-edit/tests/assets/write_to`, the fixture `dirty`, both
/// formats): a formula that reads through a name is recalculated when a reference in the definition of the name
/// moved (a band begins at or before its last row, or its last column), shrank or went; a name that is defined by other names is
/// recalculated when one of those lost cells (a band meets its target), but not when it only moved.
#[must_use]
pub fn names_reached<S: std::hash::BuildHasher>(
    names: &std::collections::HashMap<String, String, S>,
    rows: &[(String, u32, u32)],
    columns: &[(String, u32, u32)],
) -> Vec<String> {
    let same = |a: &Option<String>, b: &str| a.as_deref().is_some_and(|a| addr::same_name(a, b));
    // `all`: a band that begins at or before the reference moves it; otherwise only one that meets it counts.
    let reached = |(sheet, r0, c0, r1, c1): &NameRef, all: bool| {
        rows.iter().any(|(s, first, last)| {
            same(sheet, s)
                && if all {
                    first <= r1
                } else {
                    first <= r1 && last >= r0
                }
        }) || columns.iter().any(|(s, first, last)| {
            same(sheet, s)
                && if all {
                    first <= c1
                } else {
                    first <= c1 && last >= c0
                }
        })
    };
    let reads: std::collections::HashMap<String, (Vec<NameRef>, Vec<String>)> = names
        .iter()
        .map(|(name, definition)| (name.to_lowercase(), name_reads(definition)))
        .collect();
    let mut out: Vec<String> = Vec::new();
    for (name, (refs, inner)) in &reads {
        let mut hit = refs.iter().any(|r| reached(r, true));
        // The names it uses, and theirs, in turn.
        let (mut seen, mut stack) = (HashSet::from([name.clone()]), inner.clone());
        while let Some(next) = stack.pop() {
            if !seen.insert(next.clone()) {
                continue;
            }
            if let Some((refs, inner)) = reads.get(&next) {
                hit |= refs.iter().any(|r| reached(r, false));
                stack.extend(inner.iter().cloned());
            }
        }
        if hit {
            out.push(name.clone());
        }
    }
    out.sort();
    out
}

/// Optional compatibility behaviours, resolved while the program is built. All are off by default:
/// each one reads a formula differently from the reference (`LibreOffice` 26.8), so the user opts in.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Features {
    /// `_xlfn._LONGTEXT(…)` with constant text arguments: `LibreOffice` 26.2 wrote a text literal of
    /// more than 255 characters as pieces joined by it. The pieces are joined into one literal.
    pub longtext: bool,
}

impl Features {
    /// Every feature the configuration can name.
    pub const NAMES: &'static [&'static str] = &["LONGTEXT"];

    /// The features of an `enable_features` list (names in any case); only those listed are on.
    ///
    /// # Errors
    ///
    /// When a name is not one of the known features.
    pub fn parse(names: &[String]) -> Result<Self> {
        let mut f = Self::default();
        for n in names {
            match n.trim().to_ascii_uppercase().as_str() {
                "LONGTEXT" => f.longtext = true,
                other => {
                    return Err(Error::UnknownFeature {
                        name: other.to_string(),
                        known: Self::NAMES,
                    });
                }
            }
        }
        Ok(f)
    }
}

impl Default for Spec {
    fn default() -> Self {
        Self {
            inputs: Vec::new(),
            inputs_declared: false,
            outputs: Vec::new(),
            date_system: None,
            iteration: None,
            locale: None,
            search: None,
            now: None,
            host: sheetrs_engine::options::Host::default(),
            rand_seed: None,
            fold: true,
            mode: Mode::Values,
            meta: Vec::new(),
            skip_sheet: None,
            inspect: false,
            notes: Vec::new(),
            features: Features::default(),
            cache: Cache::default(),
            calculate_all: false,
            force_calculate_all: false,
            only: Vec::new(),
            whole_workbook: false,
            given_inputs: None,
        }
    }
}

#[derive(Debug, Default, Clone)]
pub struct Report {
    pub cells_total: usize,
    pub formulas_total: usize,
    /// Cells reachable from the outputs.
    pub cells_reachable: usize,
    /// Formulas evaluated at build time and replaced by constants.
    pub formulas_folded: usize,
    /// Formulas that remain in the program.
    pub formulas_kept: usize,
    /// Formulas that keep the result the workbook caches, because nothing reaches them (see
    /// [`Cache`]).
    pub formulas_cached: usize,
    pub cells_kept: usize,
    pub cycles: usize,
    /// A reference computed at run time forced the whole workbook to be kept.
    pub dynamic_refs: bool,
    pub unsupported: BTreeSet<String>,
    /// What the build took from the cache of the file, by kind (see [`FromCacheNote`]).
    pub from_cache: Vec<FromCacheNote>,
    /// Warnings of the build (text outside the scripts the collation orders as `LibreOffice` does…).
    pub warnings: Vec<String>,
    /// Every function called in the workbook's formulas and how many times
    /// (including the ones the engine does not implement).
    pub function_uses: std::collections::BTreeMap<String, usize>,
}

fn err<T>(msg: impl Into<String>) -> Result<T> {
    Err(Error::Compile(msg.into()))
}

const MAX_BINDING_CELLS: u64 = 10_000_000;

/// Resolve a user-facing target to a rectangle.
///
/// # Errors
///
/// When `target` is not a name or an address of the model, or names a sheet it does not have.
pub fn resolve_target(model: &Model, target: &str) -> Result<Rect> {
    let (sheet, tail) = addr::split_sheet(target.trim());
    let sheet_id = match &sheet {
        Some(n) => model.sheet_id(n),
        None => Some(0),
    };
    if let (Some(s), Some(area)) = (sheet_id, addr::parse_area(tail)) {
        return Ok(addr::bind(area, s));
    }
    if sheet.is_none()
        && let Some(def) = model
            .names
            .iter()
            .find(|(k, _)| addr::same_name(k, target.trim()))
            .map(|(_, v)| v.as_str())
    {
        let (sh, t) = addr::split_sheet(def.trim_start_matches('='));
        if let (Some(sh), Some(area)) = (
            sh.as_deref().and_then(|n| model.sheet_id(n)),
            addr::parse_area(t),
        ) {
            return Ok(addr::bind(area, sh));
        }
    }
    err(format!("cannot resolve `{target}` to a cell or range"))
}

/// What the declarations must satisfy, reported all at once:
/// an input holds constants (an input with a value keeps it as its default, and
/// one without is simply empty), and an output covers at least one value or
/// formula.
fn check_bindings(model: &Model, spec: &Spec) -> Result<()> {
    let label = |b: &BindingSpec| match &b.name {
        Some(n) => format!("`{n}` ({})", b.target),
        None => format!("`{}`", b.target),
    };
    let mut problems: Vec<String> = Vec::new();
    for b in &spec.inputs {
        let rect = resolve_target(model, &b.target)?;
        if rect.area() > MAX_BINDING_CELLS {
            return err(format!("input `{}` spans too many cells", b.target));
        }
        let mut formulas: Vec<String> = Vec::new();
        for r in rect.r0..=rect.r1 {
            for c in rect.c0..=rect.c1 {
                let at = CellRef::new(rect.sheet, r, c);
                if model.cells.get(&at).is_some_and(|mc| mc.formula.is_some()) {
                    formulas.push(addr::a1(r, c));
                }
            }
        }
        if !formulas.is_empty() {
            let shown = formulas
                .iter()
                .take(8)
                .cloned()
                .collect::<Vec<_>>()
                .join(", ");
            let more = formulas.len().saturating_sub(8);
            problems.push(format!(
                "input {} holds formulas in {shown}{}; an input must be a constant",
                label(b),
                if more > 0 {
                    format!(" and {more} more")
                } else {
                    String::new()
                }
            ));
        }
    }
    let total = model.cells.len() as u64;
    for b in &spec.outputs {
        let rect = resolve_target(model, &b.target)?;
        if rect.area() > MAX_BINDING_CELLS {
            return err(format!("output `{}` spans too many cells", b.target));
        }
        let filled = |at: &CellRef| {
            model
                .cells
                .get(at)
                .is_some_and(|c| c.formula.is_some() || c.value != Value::Empty)
        };
        let any = if rect.area() <= total * 4 + 64 {
            (rect.r0..=rect.r1)
                .any(|r| (rect.c0..=rect.c1).any(|c| filled(&CellRef::new(rect.sheet, r, c))))
        } else {
            model
                .cells
                .keys()
                .any(|at| rect.contains(at) && filled(&at))
        };
        if !any {
            problems.push(format!(
                "output {} is empty: it has no value and no formula",
                label(b)
            ));
        }
    }
    if problems.is_empty() {
        Ok(())
    } else {
        err(problems.join("\n"))
    }
}

#[derive(Default, Clone, Copy)]
struct Flags {
    dynamic: bool,
    volatile: bool,
    /// `LibreOffice` calculates the formula when it opens any file (see [`dependencies_into`]), which it does not for what a build
    /// folds: only the recalculation that keeps the cache looks at it.
    on_load: bool,
}

/// The peak resident memory of the process so far, in MB, where `/proc` reports it.
#[must_use]
pub fn peak_rss_mb() -> Option<u64> {
    let status = std::fs::read_to_string("/proc/self/status").ok()?;
    let kb = status
        .lines()
        .find_map(|l| l.strip_prefix("VmHWM:"))?
        .trim()
        .strip_suffix("kB")?
        .trim()
        .parse::<u64>()
        .ok()?;
    Some(kb / 1024)
}

/// The start of the stage timing, taken only when `SHEETC_TRACE` is set: reading the clock panics on
/// `wasm32-unknown-unknown`, where the variable is never set.
fn trace_start() -> Option<std::time::Instant> {
    std::env::var_os("SHEETC_TRACE").map(|_| std::time::Instant::now())
}

/// Phase timing and peak memory on stderr when `SHEETC_TRACE` is set.
fn trace(t0: Option<&std::time::Instant>, label: &str) {
    if let Some(t0) = t0 {
        eprintln!(
            "[compile {:>8.2?}] {label} (peak {} MB)",
            t0.elapsed(),
            peak_rss_mb().map_or("n/a".to_string(), |mb| mb.to_string())
        );
    }
}

/// A formula that depends on what the engine cannot calculate: the cell, the kind, what it is (the reference, the
/// function) and whether the engine answers it as `LibreOffice` does when the cell is calculated (an add-in that is
/// not installed is `#NAME?`).
type Found = (CellRef, Unsupported, String, bool);

/// What lowering a share of the formulas found.
#[derive(Default)]
struct Lowered {
    formulas: Vec<(CellRef, CellKind)>,
    subtotal_cells: Vec<CellRef>,
    cached: FastMap<CellRef, Value>,
    diags: Vec<String>,
    unsupported: BTreeSet<String>,
    /// The formulas that depend on what the engine cannot calculate: the cell, the kind and what it is
    /// (the reference, the function).
    found: Vec<Found>,
    /// Formulas whose syntax the parser does not read: the cell, the formula and why.
    syntax: Vec<(CellRef, String, String)>,
    /// Options of a function the engine does not support: the cell and what.
    options: Vec<(CellRef, String)>,
    function_uses: std::collections::BTreeMap<String, usize>,
}

fn lower_chunk(
    model: &Model,
    part: &[(CellRef, &ModelCell)],
    calc: sheetrs_formula::calc_print::CalcSyntax,
    features: Features,
    detect: &Detect,
) -> Lowered {
    let mut out = Lowered::default();
    for &(at, c) in part {
        let Some(f) = &c.formula else { continue };
        // A reference to another workbook is not read by the parser: what the formula depends on is told
        // before it is lowered.
        if let Some(what) = detect.external(f) {
            out.found.push((at, Unsupported::External, what, false));
            continue;
        }
        let ctx = ParseCtx {
            sheets: &model.sheets,
            current_sheet: at.sheet,
            names: &model.names,
            cells: Some(&model.cells),
            xlsx: model.from_xlsx,
            ods: model.from_file && !model.from_xlsx,
            calc,
            longtext: features.longtext,
        };
        let before = out.diags.len();
        match formula::parse(f, &ctx, &mut out.diags) {
            Ok(mut e) => {
                fold_refs(
                    &mut e,
                    at,
                    &model.sheets,
                    model.settings.string_ref_excel.unwrap_or(true),
                );
                if c.array
                    && let Some(id) =
                        sheetrs_engine::functions::lookup(sheetrs_engine::ir::Ns::Core, "_ARRAY")
                {
                    // A member of the block shows its own element of the result: the call is told
                    // how far it sits from the top-left cell.
                    let mut args = vec![e];
                    if let Some(&(rows, cols)) = model.array_members.get(&at) {
                        args.push(Expr::Const(Value::Number(f64::from(rows))));
                        args.push(Expr::Const(Value::Number(f64::from(cols))));
                    }
                    e = Expr::Call(id, args);
                }
                let mut holds_subtotal = false;
                e.walk(&mut |x| match x {
                    Expr::Unknown(call) => {
                        let name = &call.0;
                        out.unsupported.insert(name.to_string());
                        if let Some((kind, answers)) = detect.function(name) {
                            out.found.push((at, kind, name.to_string(), answers));
                        }
                        *out.function_uses.entry(name.to_string()).or_default() += 1;
                    }
                    Expr::Call(id, args) => {
                        let name = id.def().name;
                        if name == "TEXT"
                            && let Some(Expr::Const(Value::Text(code))) = args.get(1)
                            && sheetrs_engine::format::has_foreign_language_tag(code)
                        {
                            out.options.push((
                                at,
                                format!("TEXT code `{code}` names a language the engine has no month and day names for"),
                            ));
                        }
                        holds_subtotal |= matches!(name, "SUBTOTAL" | "AGGREGATE");
                        if !name.starts_with('_') {
                            *out.function_uses.entry(name.to_string()).or_default() += 1;
                        }
                    }
                    _ => {}
                });
                if holds_subtotal {
                    out.subtotal_cells.push(at);
                }
                out.formulas.push((at, CellKind::Formula(e)));
            }
            // Syntax the parser does not read (external workbook references, structured
            // references, `#` spills…): the build stops, see `compile`.
            Err(e) => {
                out.diags.push(format!("unsupported syntax ({e})"));
                out.syntax.push((at, f.clone(), e.to_string()));
            }
        }
        let sheet = model
            .sheets
            .get(at.sheet as usize)
            .map_or("?", String::as_str);
        for d in &mut out.diags[before..] {
            *d = format!("{sheet}!{}: {d}", addr::a1(at.row, at.col));
        }
    }
    out
}

/// What tells that a formula depends on what the engine cannot calculate (see [`Unsupported`]).
struct Detect {
    /// The workbook carries macros: a function no spreadsheet program has may be one of them.
    macros: bool,
    /// The defined names (uppercase) that hold a reference to another workbook.
    external_names: Vec<String>,
}

impl Detect {
    fn new(model: &Model) -> Self {
        let mut external_names: Vec<String> = model
            .names
            .iter()
            .filter(|(_, text)| external_reference(text).is_some())
            .map(|(name, _)| name.to_ascii_uppercase())
            .collect();
        external_names.sort();
        Self {
            macros: model.macros,
            external_names,
        }
    }

    /// The external reference (or the defined name that holds one) the formula text depends on.
    fn external(&self, formula: &str) -> Option<String> {
        if let Some(reference) = external_reference(formula) {
            return Some(reference);
        }
        if self.external_names.is_empty() {
            return None;
        }
        let up = formula.to_ascii_uppercase();
        self.external_names
            .iter()
            .find(|name| contains_word(&up, name))
            .cloned()
    }

    /// The kind of what a call to the function `name` (uppercase, prefixes taken off) depends on, when the
    /// engine does not have it, and whether the engine answers the call as `LibreOffice` does when the cell is
    /// calculated: an add-in that is not installed (`_xll.NAME`, `COM.SUN.STAR.SHEET.ADDIN.*`) is `#NAME?`
    /// (OBSERVED, both formats), and the engine gives that for a name it does not know. A function of
    /// `LibreOffice` that the engine lacks (`WEBSERVICE`, `FILTERXML`, `DDE`...) and one that a macro of the
    /// workbook may define have results that only `LibreOffice` can give. `None`: `LibreOffice` does not know the
    /// name either (`#NAME?`).
    fn function(&self, name: &str) -> Option<(Unsupported, bool)> {
        if name.starts_with("_XLL.") || name.starts_with("COM.SUN.STAR.SHEET.ADDIN.") {
            Some((Unsupported::Extension, true))
        } else if sheetrs_engine::functions::in_catalogue(name) {
            Some((Unsupported::Extension, false))
        } else if self.macros {
            Some((Unsupported::Macro, false))
        } else {
            None
        }
    }
}

/// The reference to another workbook in the text of a formula (`[1]Sheet1!A1` of an XLSX, which an ODS
/// is read as too), outside text between quotes: the sheet it names, with the index of the workbook
/// (`[1]Sheet1`). `[1]` before a sheet name and `!` is what an external workbook looks like; a structured
/// reference (`Table1[[#This Row],[1]]`) is not followed by `!`.
fn external_reference(text: &str) -> Option<String> {
    let bytes = text.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        match bytes[i] {
            b'"' => {
                i += 1;
                while i < bytes.len() {
                    if bytes[i] == b'"' {
                        if bytes.get(i + 1) == Some(&b'"') {
                            i += 2;
                            continue;
                        }
                        break;
                    }
                    i += 1;
                }
                i += 1;
            }
            b'[' => {
                let digits = bytes[i + 1..]
                    .iter()
                    .take_while(|b| b.is_ascii_digit())
                    .count();
                let close = i + 1 + digits;
                if digits > 0 && bytes.get(close) == Some(&b']') {
                    // The sheet name: up to the `!`, with no operator or space unless it is quoted.
                    let mut j = close + 1;
                    let mut quoted = false;
                    while j < bytes.len() {
                        match bytes[j] {
                            b'\'' => quoted = !quoted,
                            b'!' if !quoted => {
                                // With the cell or range it reads (`$A$1:B2`).
                                let tail = bytes[j + 1..]
                                    .iter()
                                    .take_while(|b| {
                                        b.is_ascii_alphanumeric() || matches!(b, b'$' | b':')
                                    })
                                    .count();
                                return Some(text[i..j + 1 + tail].to_string());
                            }
                            b'(' | b')' | b',' | b';' | b'+' | b'-' | b'*' | b'/' | b'^' | b'&'
                            | b'=' | b'<' | b'>' | b' ' | b'"' | b'[' | b']'
                                if !quoted =>
                            {
                                break;
                            }
                            _ => {}
                        }
                        j += 1;
                    }
                }
                i += 1;
            }
            _ => i += 1,
        }
    }
    None
}

/// Whether `word` is in `text` as a whole name (not inside another one).
fn contains_word(text: &str, word: &str) -> bool {
    let is_name = |b: u8| b.is_ascii_alphanumeric() || b == b'_' || b == b'.';
    let bytes = text.as_bytes();
    let mut from = 0;
    while let Some(at) = text[from..].find(word) {
        let start = from + at;
        let end = start + word.len();
        let before = start.checked_sub(1).map(|k| bytes[k]);
        let after = bytes.get(end).copied();
        if !before.is_some_and(is_name) && !after.is_some_and(is_name) {
            return true;
        }
        from = start + 1;
    }
    false
}

/// What [`compile`] takes as the model: `&Model` (the caller keeps it, unchanged) or a `Model` (the
/// compiler owns it and frees what it no longer needs).
pub trait ModelArg<'m> {
    fn into_cow(self) -> Cow<'m, Model>;
}

impl<'m> ModelArg<'m> for &'m Model {
    fn into_cow(self) -> Cow<'m, Model> {
        Cow::Borrowed(self)
    }
}

impl ModelArg<'_> for Model {
    fn into_cow(self) -> Cow<'static, Model> {
        Cow::Owned(self)
    }
}

/// # Errors
///
/// When the document defines a setting that the spec overrides, or the model cannot be compiled (an unknown name, a reference that cannot be resolved, a formula that does not parse).
///
/// # Panics
///
/// When a thread that lowers formulas panics.
pub fn compile<'m>(model: impl ModelArg<'m>, spec: &Spec) -> Result<(Program, Report)> {
    // A model passed by value is the compiler's: once every formula is lowered the text of the
    // formulas is freed (nothing after reads it), which lowers the peak of the rest of the build.
    // A model passed by reference is the caller's and stays as it is.
    let mut owned: Cow<'m, Model> = model.into_cow();
    // A workbook that carries the mark to calculate every formula when it is opened is a run with `calculate_all`:
    // what it caches is not current. One place decides it, for every path that builds or calculates.
    let marked;
    let spec = if owned.recalculate_on_load && !spec.calculate_all && !spec.inspect {
        marked = Spec {
            calculate_all: true,
            ..spec.clone()
        };
        &marked
    } else {
        spec
    };
    // `calculateAll()`: the whole workbook is calculated first and its results are the values the
    // workbook stores from here on (the folded constants and the cache of what is kept).
    if spec.calculate_all && spec.mode != Mode::Recalculate && !spec.inspect {
        let calculated = calculated_model(&owned, spec)?;
        owned = Cow::Owned(calculated);
    }
    let model: &Model = &owned;
    // The editor's recalculation (`Mode::Recalculate`, built by no command) has no declarations: every
    // formula is an output and any constant an input. An embedded program takes them as any program
    // does, and deduces what it does not declare.
    let cleared;
    let spec = if spec.mode == Mode::Recalculate
        && (!spec.inputs.is_empty() || spec.inputs_declared || !spec.outputs.is_empty())
    {
        let mut notes = spec.notes.clone();
        notes.push(
            "the inputs and outputs the settings declare are discarded: in this mode every formula is an output and any constant an input"
                .into(),
        );
        cleared = Spec {
            notes,
            inputs: Vec::new(),
            inputs_declared: false,
            outputs: Vec::new(),
            // What the settings declared as inputs and outputs is not exposed: the rest is kept.
            meta: spec
                .meta
                .iter()
                .filter(|m| !matches!(m.meta_type.as_str(), "input" | "output"))
                .cloned()
                .collect(),
            ..spec.clone()
        };
        &cleared
    } else {
        spec
    };
    // With no input declared they are deduced, in every mode: the dependency graph decides, the
    // inputs are the constants a path starts from. Only a recalculated workbook (no input is given
    // to it) keeps every constant as a node, as it is.
    let deduce_inputs = spec.inputs.is_empty() && !spec.inputs_declared;
    let every_constant_input = deduce_inputs && spec.mode == Mode::Recalculate;
    let analyse_inputs = deduce_inputs && !every_constant_input;
    // An embedded workbook with no outputs declared calculates every formula and writes them all:
    // the inputs it deduces (the constants some formula reads) need no node in the graph, which
    // would only taint what depends on them for a fold that has nothing to fold. They are listed as
    // the inputs, and the constants stay where they are, as in a workbook mode that takes them all.
    let lazy_inputs = analyse_inputs && spec.mode == Mode::Embedded && spec.outputs.is_empty();
    let fold = spec.fold && !every_constant_input && !lazy_inputs;
    let mut deduced: Vec<CellRef> = Vec::new();
    let t0 = trace_start();
    if model.from_file
        && let Some(what) = [
            (spec.date_system.is_some(), "date system"),
            (spec.iteration.is_some(), "iteration"),
            (spec.search.is_some(), "search settings"),
        ]
        .iter()
        .find_map(|(set, name)| set.then_some(*name))
    {
        return err(format!(
            "the {what} is defined by the document and cannot be overridden"
        ));
    }
    let mut report = Report::default();
    if model.settings.manual_calc {
        report.warnings.push(
            "the document is in manual calculation mode; the program calculates automatically (INFO(\"recalc\") is `Automatic`)"
                .into(),
        );
    }
    let mut diags: Vec<String> = model.diagnostics.clone();
    diags.extend(spec.notes.iter().cloned());
    let mut parse_diags: Vec<String> = Vec::new();

    // ----- 1. lower formulas -----
    // Only formulas and inputs become graph nodes. Constants stay in the model
    // and are read in place: a 50 MB workbook holds millions of them and none
    // needs ordering.
    let calc_syntax = sheetrs_formula::calc_print::CalcSyntax::for_decimal(
        spec.locale.as_ref().map_or('.', |l| l.decimal),
    );
    let mut formulas: Vec<(CellRef, CellKind)> = Vec::new();
    let mut subtotal_cells: Vec<CellRef> = Vec::new();
    // Formulas the parser rejects keep the value the file last computed.
    let mut cached: FastMap<CellRef, Value> = FastMap::default();
    report.cells_total += model.cells.len();
    let formula_cells: Vec<(CellRef, &ModelCell)> = model
        .cells
        .iter()
        .filter(|(_, c)| c.formula.is_some())
        .collect();
    report.formulas_total += formula_cells.len();
    // Where the formulas are, whatever becomes of each below (lowered, folded, or kept as the value
    // the file cached): the outputs nothing declares are the cells that are calculated.
    let mut formula_positions: Vec<CellRef> = formula_cells.iter().map(|(at, _)| *at).collect();
    // Lowering a formula reads only the model: the cells go to as many threads as
    // the machine has, and what each one found is merged afterwards (the result is
    // sorted below, so the order they finish in does not matter).
    let detect = Detect::new(model);
    let detect = &detect;
    let workers = std::thread::available_parallelism()
        .map_or(1, usize::from)
        .min(8);
    let chunk = formula_cells.len().div_ceil(workers).max(1);
    let lowered: Vec<Lowered> = if workers == 1 {
        // No thread to start: the only way on `wasm32-unknown-unknown`, which has none.
        vec![lower_chunk(
            model,
            &formula_cells,
            calc_syntax,
            spec.features,
            detect,
        )]
    } else {
        std::thread::scope(|scope| {
            let handles: Vec<_> = formula_cells
                .chunks(chunk)
                .map(|part| {
                    scope
                        .spawn(move || lower_chunk(model, part, calc_syntax, spec.features, detect))
                })
                .collect();
            handles
                .into_iter()
                .map(|h| h.join().expect("a lowering thread panicked"))
                .collect()
        })
    };
    let mut found: Vec<Found> = Vec::new();
    let mut syntax: Vec<(CellRef, String, String)> = Vec::new();
    let mut options: Vec<(CellRef, String)> = Vec::new();
    for part in lowered {
        formulas.extend(part.formulas);
        subtotal_cells.extend(part.subtotal_cells);
        cached.extend(part.cached);
        parse_diags.extend(part.diags);
        report.unsupported.extend(part.unsupported);
        syntax.extend(part.syntax);
        options.extend(part.options);
        found.extend(part.found);
        for (name, n) in part.function_uses {
            *report.function_uses.entry(name).or_default() += n;
        }
    }
    formulas.sort_unstable_by_key(|(at, _)| *at);
    subtotal_cells.sort_unstable();
    // A formula that reads another workbook is not in the program (the parser does not read it): it keeps the
    // result the file caches, which is the stored value of the link that `LibreOffice` shows.
    if !found.is_empty() && !spec.inspect {
        leave_external(
            model,
            &found,
            &mut subtotal_cells,
            &mut formula_positions,
            &mut cached,
        );
    }
    // The calls the engine cannot answer, in the cells that stay in the program: `keep_cached` takes out of it those
    // nothing reaches, which keep the result the file caches, and the rest is for [`settle_found`].
    let exempt: FastSet<CellRef> = found
        .iter()
        .filter(|f| f.1 != Unsupported::External && !f.3)
        .map(|f| f.0)
        .collect();
    // The default of every run is the load of `LibreOffice`: a formula that has a result and that
    // nothing reaches keeps it. A recalculation is told what changed ([`Cache`]); any other mode treats the
    // declared inputs as the cells that may change. With the inputs to be deduced every formula that
    // reads a constant is on a path from one, so there is nothing to cut.
    let analyse = !(spec.inspect || spec.mode == Mode::Recalculate && spec.calculate_all);
    // A recalculation cuts what nothing reaches; another mode cuts what no declared input reaches, and
    // with the inputs to be deduced every formula that reads a constant is on a path from one, so
    // nothing is cut (only the stale formulas are found, to keep no result for them).
    let cut = analyse
        && (spec.mode == Mode::Recalculate || !spec.inputs.is_empty() || spec.inputs_declared);
    let mut stale_formulas: Option<FastSet<CellRef>> = None;
    if analyse {
        let inputs: Vec<Rect> = if spec.mode == Mode::Recalculate {
            Vec::new()
        } else {
            let declared = |b: &BindingSpec| resolve_target(model, &b.target).ok();
            match &spec.given_inputs {
                // A calculation: the cells it was given are the only ones that may change.
                Some(given) => {
                    let rects: Option<Vec<Rect>> = given
                        .iter()
                        .map(|key| {
                            match spec.inputs.iter().find(|b| {
                                b.name
                                    .as_deref()
                                    .is_some_and(|n| n.eq_ignore_ascii_case(key))
                            }) {
                                Some(b) => declared(b),
                                None => resolve_target(model, key).ok(),
                            }
                        })
                        .collect();
                    // A key that is no cell the model has: any declared input may be what it means.
                    rects.unwrap_or_else(|| spec.inputs.iter().filter_map(declared).collect())
                }
                None => spec.inputs.iter().filter_map(declared).collect(),
            }
        };
        // With every formula calculated first, the workbook that is written has the results of all of
        // them: no formula keeps what the template stored.
        let keep_positions = !spec.calculate_all;
        let kept = keep_cached(
            model,
            &spec.cache,
            &inputs,
            cut,
            keep_positions,
            &formula_cells,
            &exempt,
            &mut formulas,
            &mut cached,
            &mut formula_positions,
        );
        report.formulas_cached = kept.count;
        stale_formulas = kept.stale;
    }
    if !found.is_empty() && !spec.inspect {
        settle_found(model, &found, &formulas, &mut report)?;
    }
    drop(formula_cells);
    if let Cow::Owned(own) = &mut owned {
        own.cells.release_formulas();
    }
    let model: &Model = &owned;

    trace(t0.as_ref(), "parsed formulas");
    if !model.unsupported_options.is_empty() && !spec.inspect {
        return Err(Error::Compile(format!(
            "the workbook uses an option this engine does not support, so it cannot replace LibreOffice for it: {}",
            model.unsupported_options.join("; ")
        )));
    }
    if !syntax.is_empty() && !spec.inspect {
        syntax.sort_by_key(|(at, ..)| *at);
        let list: Vec<String> = syntax
            .iter()
            .take(5)
            .map(|(at, f, why)| {
                let sheet = model
                    .sheets
                    .get(at.sheet as usize)
                    .map_or("?", String::as_str);
                format!(
                    "{sheet}!{} `{f}` ({why})",
                    sheetrs_engine::addr::a1(at.row, at.col)
                )
            })
            .collect();
        return Err(Error::Compile(format!(
            "{} formula(s) use syntax this engine does not support, so it cannot replace LibreOffice for the workbook; first: {}",
            syntax.len(),
            list.join("; ")
        )));
    }
    if !options.is_empty() && !spec.inspect {
        options.sort();
        let list: Vec<String> = options
            .iter()
            .take(5)
            .map(|(at, what)| {
                let sheet = model
                    .sheets
                    .get(at.sheet as usize)
                    .map_or("?", String::as_str);
                format!(
                    "{sheet}!{}: {what}",
                    sheetrs_engine::addr::a1(at.row, at.col)
                )
            })
            .collect();
        return Err(Error::Compile(format!(
            "{} formula(s) use an option this engine does not support, so it cannot replace LibreOffice for the workbook; first: {}",
            options.len(),
            list.join("; ")
        )));
    }
    check_bindings(model, spec)?;
    // ----- inputs / outputs -----
    let mut inputs = Vec::new();
    let mut input_nodes: Vec<(CellRef, CellKind)> = Vec::new();
    for b in &spec.inputs {
        let rect = resolve_target(model, &b.target)?;
        if rect.area() > MAX_BINDING_CELLS {
            return err(format!("input `{}` spans too many cells", b.target));
        }
        let name = b
            .name
            .clone()
            .unwrap_or_else(|| addr::qualified(&model.sheets, &rect));
        for r in rect.r0..=rect.r1 {
            for c in rect.c0..=rect.c1 {
                let at = CellRef::new(rect.sheet, r, c);
                let default = model
                    .cells
                    .get(&at)
                    .map_or(Value::Empty, |mc| mc.value.clone());
                input_nodes.push((at, CellKind::Input(default)));
            }
        }
        inputs.push(Binding { name, rect });
    }
    if lazy_inputs {
        deduced = constants_that_start_a_path(model, &formulas, spec.skip_sheet);
    } else if analyse_inputs {
        input_nodes.extend(
            constants_that_start_a_path(model, &formulas, spec.skip_sheet)
                .into_iter()
                .map(|at| {
                    let value = model
                        .cells
                        .get(&at)
                        .map_or(Value::Empty, |mc| mc.value.clone());
                    (at, CellKind::Input(value))
                }),
        );
    }
    trace(t0.as_ref(), "constants that start a path");
    if !input_nodes.is_empty() {
        // An input replaces whatever formula or constant sat in its cell. Both lists are in order
        // of position (`formulas` was sorted above), so one merge puts them together.
        input_nodes.sort_by_key(|(a, _)| *a);
        input_nodes.dedup_by_key(|(a, _)| *a);
        let mut merged: Vec<(CellRef, CellKind)> =
            Vec::with_capacity(formulas.len() + input_nodes.len());
        let mut rest = std::mem::take(&mut formulas).into_iter().peekable();
        for (at, kind) in input_nodes {
            while rest.peek().is_some_and(|(a, _)| *a < at) {
                merged.extend(rest.next());
            }
            if rest.peek().is_some_and(|(a, _)| *a == at) {
                rest.next();
                cached.remove(&at);
            }
            merged.push((at, kind));
        }
        merged.extend(rest);
        formulas = merged;
    }
    // The cells a calculation is asked for, when it is (`Spec::only`).
    let asked_areas: Vec<Rect> = spec
        .only
        .iter()
        .filter_map(|o| only_area(model, o))
        .collect();
    let asked = |at: &CellRef| {
        spec.only.is_empty()
            || asked_areas.iter().any(|r| {
                r.sheet == at.sheet
                    && (r.r0..=r.r1).contains(&at.row)
                    && (r.c0..=r.c1).contains(&at.col)
            })
    };
    let mut outputs = Vec::new();
    for b in &spec.outputs {
        let rect = resolve_target(model, &b.target)?;
        if rect.area() > MAX_BINDING_CELLS {
            return err(format!("output `{}` spans too many cells", b.target));
        }
        outputs.push(Binding {
            name: b
                .name
                .clone()
                .unwrap_or_else(|| addr::qualified(&model.sheets, &rect)),
            rect,
        });
    }
    if spec.outputs.is_empty() {
        // Nothing declared: the outputs are the cells that are calculated (the formulas, not the
        // constants), as the rectangles that hold them, the same in every mode.
        let positions: Vec<CellRef> = formula_positions
            .iter()
            .copied()
            .filter(|at| Some(at.sheet) != spec.skip_sheet && asked(at))
            .collect();
        let rects = formula_rects(&positions);
        if rects.len() > MAX_FORMULA_OUTPUTS {
            // A scatter of formulas that needs a rectangle for each few of them: the used area of
            // the formulas of each sheet, constants in between included, is one output per sheet.
            diags.push(format!(
                "the formulas lie in {} separate rectangles: each sheet's outputs are the area that holds its formulas, constants in it included",
                rects.len()
            ));
            let mut bbs: Vec<Option<Rect>> = vec![None; model.sheets.len()];
            for r in &rects {
                let bb = &mut bbs[usize::from(r.sheet)];
                *bb = Some(match *bb {
                    None => *r,
                    Some(b) => Rect::new(
                        r.sheet,
                        b.r0.min(r.r0),
                        b.c0.min(r.c0),
                        b.r1.max(r.r1),
                        b.c1.max(r.c1),
                    ),
                });
            }
            for rect in bbs.into_iter().flatten() {
                outputs.push(Binding {
                    name: addr::qualified(&model.sheets, &rect),
                    rect,
                });
            }
        } else {
            for rect in rects {
                outputs.push(Binding {
                    name: addr::qualified(&model.sheets, &rect),
                    rect,
                });
            }
        }
    }
    for list in [&inputs, &outputs] {
        let mut names = HashSet::new();
        for b in list {
            if !names.insert(b.name.as_str()) {
                return err(format!("`{}` is declared twice", b.name));
            }
        }
    }

    // The defined names that can be used as aliases: a name that points inside the outputs is an
    // alias of that output, one that points inside the declared inputs of those inputs (one inside
    // deduced inputs is added once they are known). None may point into the settings sheet.
    let mut named: Vec<Binding> = Vec::new();
    let mut outside: Vec<Binding> = Vec::new();
    {
        let mut skipped = Vec::new();
        let inside = |rect: &Rect, list: &[Binding]| {
            list.iter().any(|b| {
                let o = &b.rect;
                o.sheet == rect.sheet
                    && o.r0 <= rect.r0
                    && o.c0 <= rect.c0
                    && o.r1 >= rect.r1
                    && o.c1 >= rect.c1
            })
        };
        for name in model.names.keys() {
            match resolve_target(model, name) {
                Ok(rect) if Some(rect.sheet) == spec.skip_sheet => {}
                Ok(rect) if rect.area() <= MAX_BINDING_CELLS => {
                    let b = Binding {
                        name: name.clone(),
                        rect,
                    };
                    if inside(&rect, &outputs) || inside(&rect, &inputs) {
                        named.push(b);
                    } else {
                        outside.push(b);
                    }
                }
                _ => skipped.push(name.as_str()),
            }
        }
        if !skipped.is_empty() {
            skipped.sort_unstable();
            diags.push(format!(
                "defined names that are not a plain range cannot be used as aliases: {}",
                skipped.join(", ")
            ));
        }
    }

    trace(t0.as_ref(), "bindings");
    let n = formulas.len();
    let (ats, mut kinds): (Vec<CellRef>, Vec<CellKind>) = formulas.into_iter().unzip();
    let index_owned = NodeIndex::new(&ats, model.sheets.len());
    let index = &index_owned;
    let view = ConstView::new(model, &cached);

    trace(t0.as_ref(), "index built");
    // ----- 2. dependency graph, lazily, from the outputs -----
    let mut g = Lazy {
        index,
        kinds: &kinds,
        adj: vec![None; n],
        rect_ids: FastMap::default(),
        rect_members: Vec::new(),
        flags: vec![Flags::default(); n],
        deps: Vec::new(),
    };
    let in_outputs = |at: &CellRef| {
        outputs.iter().any(|b| {
            let r = &b.rect;
            r.sheet == at.sheet
                && (r.r0..=r.r1).contains(&at.row)
                && (r.c0..=r.c1).contains(&at.col)
        })
    };
    // What the program has to compute: the outputs.
    let mut wanted: Vec<Rect> = outputs.iter().map(|b| b.rect).collect();
    if spec.whole_workbook && !spec.outputs.is_empty() {
        // The formulas the declared outputs do not hold are calculated too.
        let outside: Vec<CellRef> = formula_positions
            .iter()
            .copied()
            .filter(|at| Some(at.sheet) != spec.skip_sheet && asked(at) && !in_outputs(at))
            .collect();
        wanted.extend(formula_rects(&outside));
    }
    let mut roots: Vec<u32> = Vec::new();
    for r in &wanted {
        g.members(r, &mut roots);
    }
    let mut visited = vec![false; n];
    let mut cone: Vec<u32> = Vec::new();
    let mut stack = roots;
    let mut all = false;
    while let Some(v) = stack.pop() {
        if Lazy::seen(&mut visited, v) {
            continue;
        }
        cone.push(v);
        g.adj_of(v);
        if (v as usize) < n && g.flags[v as usize].dynamic && !all {
            all = true;
            report.dynamic_refs = true;
            stack.extend(0..n as u32);
        }
        stack.extend_from_slice(g.adj_of(v));
    }
    // Compact ids over the cone.
    let mut compact = vec![u32::MAX; visited.len()];
    for (i, &v) in cone.iter().enumerate() {
        compact[v as usize] = i as u32;
    }
    // The graph over the compact ids, built straight from the adjacency of each node of the cone
    // (every one has been read above): no second copy of the adjacency in between.
    let csr = Csr::from_rows(
        cone.iter()
            .map(|&v| g.adj_known(v).iter().map(|&w| compact[w as usize])),
    );
    // The adjacency is in the CSR now; what the taint still reads of the graph is the flags.
    let flags = std::mem::take(&mut g.flags);
    drop(g);
    trace(t0.as_ref(), "cone + csr");
    let sccs = graph::tarjan(&csr);
    report.cells_reachable = cone.iter().filter(|&&v| (v as usize) < n).count();

    trace(t0.as_ref(), "scc");
    // ----- 3/4. taint and fold -----
    let mut tainted = vec![false; cone.len()];
    let mut comp_cyclic = Vec::new();
    for comp in sccs.components() {
        let mut t = false;
        for &m in comp {
            let v = cone[m as usize];
            if (v as usize) < n {
                match &kinds[v as usize] {
                    CellKind::Input(_) => t = true,
                    CellKind::Formula(e) => {
                        let f = flags[v as usize];
                        let _ = e;
                        t |= !fold || f.dynamic || f.volatile;
                    }
                    CellKind::Const(_) => {}
                }
            }
            t |= csr.edges(m).iter().any(|&d| tainted[d as usize]);
        }
        for &m in comp {
            tainted[m as usize] = t;
        }
        comp_cyclic.push(graph::is_cycle(&csr, comp));
    }
    drop(csr);
    // The formulas whose result is never kept: the ones that give another result each time they are
    // calculated, and the ones that read cells the graph does not show.
    let no_cache: FastSet<CellRef> = cone
        .iter()
        .filter(|&&v| (v as usize) < n && (flags[v as usize].volatile || flags[v as usize].dynamic))
        .map(|&v| ats[v as usize])
        .collect();

    // The locale is resolved by `locale::resolve` (the configuration, the workbook, the default of
    // its format) into `spec.locale`; a spec that is not built from a configuration has `en`. The
    // language of the document is told when the locale of the run is another one.
    let mut locale = spec.locale.clone().unwrap_or_default();
    if let Some(lang) = &model.settings.locale
        && !lang
            .to_lowercase()
            .starts_with(locale.tag.split('-').next().unwrap_or("en"))
    {
        diags.push(format!(
            "the document language is {lang}; text and number conversions follow the locale {}",
            locale.tag
        ));
    }
    // The two-digit-year window, in contrast, is stored in the document.
    if let Some(y) = model.settings.two_digit_year_start {
        locale.two_digit_year_start = y;
    }
    let opts = EngineOptions {
        date_system: spec.date_system.unwrap_or(model.date_system),
        iteration: spec.iteration.or(model.settings.iteration),
        locale,
        search: spec.search.or(model.settings.search).unwrap_or_default(),
        string_ref_excel: model.settings.string_ref_excel.unwrap_or(true),
        now: spec.now,
        rand_seed: spec.rand_seed,
        host: {
            let mut h = spec.host.clone();
            h.filename = h.filename.or_else(|| model.file_name.clone());
            h
        },
    };

    trace(t0.as_ref(), "taint");
    // Evaluate the input-independent formulas now, reading the workbook's
    // constants in place.
    let mut folded: Vec<Option<Value>> = vec![None; n];
    let mut blank_ids: FastSet<u32> = FastSet::default();
    {
        let mut cells = Vec::new();
        // The node of each cell handed to the engine, in the same order: the formulas move there
        // and come back (a copy of every folded formula was the largest cost of this stage).
        let mut handed: Vec<u32> = Vec::new();
        let mut cycles = Vec::new();
        for (ci, comp) in sccs.components().enumerate() {
            if tainted[comp[0] as usize] {
                continue;
            }
            let start = cells.len() as u32;
            for &m in comp {
                let v = cone[m as usize];
                if (v as usize) < n {
                    cells.push(CellDef {
                        at: ats[v as usize],
                        kind: std::mem::replace(
                            &mut kinds[v as usize],
                            CellKind::Const(Value::Empty),
                        ),
                    });
                    handed.push(v);
                }
            }
            if comp_cyclic[ci] {
                cycles.push((start, cells.len() as u32));
            }
        }
        let tmp = Program {
            sheets: model.sheets.clone(),
            options: opts.clone(),
            cells,
            hidden_rows: model.hidden_rows.clone(),
            subtotal_cells: subtotal_cells.clone(),
            cycles,
            cache: Vec::new(),
            inputs: Vec::new(),
            outputs: Vec::new(),
            queryable: Queryable::default(),
            mode: Mode::Values,
            formula_cells: Vec::new(),
            meta: Vec::new(),
            diagnostics: Vec::new(),
            build: BuildInfo::default(),
        };
        let mut eng = Engine::with_base(&tmp, opts.clone(), &view);
        eng.calculate();
        if !eng.violations.is_empty() && !spec.inspect {
            let mut kinds: Vec<&'static str> = eng.violations.iter().map(|v| v.1).collect();
            kinds.sort_unstable();
            kinds.dedup();
            let mut msgs = Vec::new();
            for what in kinds {
                let mut at: Vec<CellRef> = eng
                    .violations
                    .iter()
                    .filter(|v| v.1 == what)
                    .map(|v| v.0)
                    .collect();
                at.sort_unstable();
                at.dedup();
                let list: Vec<String> = at
                    .iter()
                    .take(5)
                    .map(|a| {
                        let sheet = model
                            .sheets
                            .get(a.sheet as usize)
                            .map_or("?", String::as_str);
                        format!("{sheet}!{}", sheetrs_engine::addr::a1(a.row, a.col))
                    })
                    .collect();
                msgs.push(format!(
                    "{} formula(s) {what}, so the build cannot reproduce them: {}",
                    at.len(),
                    list.join(", ")
                ));
            }
            return Err(Error::Compile(msgs.join("; ")));
        }
        for (d, &v) in tmp.cells.iter().zip(&handed) {
            report.formulas_folded += 1;
            folded[v as usize] = Some(eng.get(d.at));
            // An empty result is not a plain 0: keep its formula (see Engine).
            if eng.blank_result(d.at) {
                blank_ids.insert(v);
            }
        }
        drop(eng);
        for (d, v) in tmp.cells.into_iter().zip(handed) {
            kinds[v as usize] = d.kind;
        }
    }

    trace(t0.as_ref(), "fold evaluated");
    // What the surviving computation (and the outputs) still read: folded
    // formulas as constants, and the workbook's own constants.
    let mut need_rects: FastSet<Rect> = FastSet::default();
    // A rectangle inside one of the outputs adds nothing the output does not have: with the whole of
    // every sheet an output (nothing declared) that is almost every rectangle of the formulas, so
    // they are not kept: a set of every reference of every formula was the largest thing this stage
    // held, only to be skipped one by one.
    let inside_an_output = |r: &Rect| {
        !wanted.contains(r)
            && wanted.iter().any(|o| {
                o.sheet == r.sheet && o.r0 <= r.r0 && o.c0 <= r.c0 && o.r1 >= r.r1 && o.c1 >= r.c1
            })
    };
    let mut refs: Vec<Rect> = Vec::new();
    // With the inputs deduced, the constants the formulas read are the inputs already (nodes of the
    // graph, or listed apart when they are lazy): what these rectangles can still add is a folded
    // formula, so with every formula tainted (nothing folded) only the outputs may hold constants
    // of their own.
    let nothing_folded = tainted.iter().all(|&t| t);
    if !(lazy_inputs || analyse_inputs && nothing_folded) {
        for (i, &v) in cone.iter().enumerate() {
            if tainted[i]
                && (v as usize) < n
                && let CellKind::Formula(e) = &kinds[v as usize]
            {
                refs.clear();
                dependencies_into(e, &mut refs);
                need_rects.extend(refs.iter().copied().filter(|r| !inside_an_output(r)));
            }
        }
        for &id in &blank_ids {
            if let CellKind::Formula(e) = &kinds[id as usize] {
                refs.clear();
                dependencies_into(e, &mut refs);
                need_rects.extend(refs.iter().copied().filter(|r| !inside_an_output(r)));
            }
        }
    }
    need_rects.extend(wanted.iter().copied());
    let mut need_folded: FastSet<u32> = FastSet::default();
    let mut const_cells: Vec<CellRef> = Vec::new();
    // Where the constants that are not in order begin, when the first run is (see `all_consts`).
    let mut ordered_until = 0;
    let mut members = Vec::new();
    for r in &need_rects {
        members.clear();
        index.members(r, &mut members);
        need_folded.extend(members.iter().copied().filter(|&id| {
            compact[id as usize] != u32::MAX && !tainted[compact[id as usize] as usize]
        }));
        view.consts_in(r, index, &mut const_cells);
    }
    if all {
        need_folded.extend(
            cone.iter()
                .copied()
                .filter(|&v| (v as usize) < n && !tainted[compact[v as usize] as usize]),
        );
        const_cells.clear();
        ordered_until = view.all_consts(index, &mut const_cells);
    }
    // The constants of the model are in order: sort only what is not, and merge the two runs.
    let mut rest = const_cells.split_off(ordered_until);
    rest.sort_unstable();
    if ordered_until == 0 {
        const_cells = rest;
    } else {
        let mut merged = Vec::with_capacity(const_cells.len() + rest.len());
        let (mut a, mut b) = (const_cells.iter().peekable(), rest.iter().peekable());
        while let (Some(x), Some(y)) = (a.peek(), b.peek()) {
            merged.push(
                if x <= y { a.next() } else { b.next() }
                    .copied()
                    .expect("peeked"),
            );
        }
        merged.extend(a.copied());
        merged.extend(b.copied());
        const_cells = merged;
    }
    const_cells.dedup();
    if !deduced.is_empty() {
        // The deduced inputs are cells of the program, read or not: a name may cover them.
        let mut union = Vec::with_capacity(const_cells.len() + deduced.len());
        let (mut a, mut b) = (const_cells.iter().peekable(), deduced.iter().peekable());
        loop {
            let next = match (a.peek(), b.peek()) {
                (Some(x), Some(y)) => match x.cmp(y) {
                    std::cmp::Ordering::Less => a.next(),
                    std::cmp::Ordering::Greater => b.next(),
                    std::cmp::Ordering::Equal => {
                        b.next();
                        a.next()
                    }
                },
                (Some(_), None) => a.next(),
                (None, Some(_)) => b.next(),
                (None, None) => break,
            };
            union.extend(next.copied());
        }
        const_cells = union;
    }

    trace(t0.as_ref(), "needed");
    // ----- assemble -----
    // The constants of the workbook are in order already; the formulas folded to a value come out
    // of a hash set. The nodes are numbered in the order of their positions (`ats` is sorted), so
    // sorting their numbers (four bytes each) puts them in order, and the two runs are merged: no
    // sort of the fat pairs, and each value is built once.
    let mut folded_ids: Vec<u32> = need_folded
        .iter()
        .copied()
        .filter(|id| !blank_ids.contains(id))
        .collect();
    folded_ids.sort_unstable();
    let mut cells: Vec<CellDef> = Vec::with_capacity(const_cells.len() + folded_ids.len());
    let (mut next_const, mut next_folded) =
        (const_cells.iter().peekable(), folded_ids.iter().peekable());
    let mut next_input = deduced.iter().peekable();
    loop {
        let take_const = match (next_const.peek(), next_folded.peek()) {
            (Some(c), Some(&&f)) => **c <= ats[f as usize],
            (Some(_), None) => true,
            (None, Some(_)) => false,
            (None, None) => break,
        };
        let (at, v) = if take_const {
            let at = *next_const.next().expect("peeked");
            (at, view.get(at).unwrap_or_default())
        } else {
            let id = *next_folded.next().expect("peeked");
            let v = folded[id as usize].clone();
            (ats[id as usize], v.expect("a folded formula has its value"))
        };
        // In a workbook mode every constant is an input; otherwise the ones that start a path are
        // already nodes of the graph (`Input`), and the rest are constants.
        let is_input = take_const
            && (every_constant_input || {
                while next_input.peek().is_some_and(|d| **d < at) {
                    next_input.next();
                }
                next_input.peek().is_some_and(|d| **d == at)
            });
        let kind = if is_input {
            CellKind::Input(v)
        } else {
            CellKind::Const(v)
        };
        cells.push(CellDef { at, kind });
    }
    if every_constant_input {
        // The inputs are all the constants the program keeps, as the fewest rectangles that cover
        // exactly them; a name that points inside them is an alias of those inputs.
        for rect in coalesce(&const_cells) {
            inputs.push(Binding {
                name: addr::qualified(&model.sheets, &rect),
                rect,
            });
        }
        add_input_aliases(&mut named, outside, &const_cells);
    } else if lazy_inputs {
        // The inputs are the constants some formula reads (or a name covers), as the fewest
        // rectangles that cover exactly them.
        for rect in coalesce(&deduced) {
            inputs.push(Binding {
                name: addr::qualified(&model.sheets, &rect),
                rect,
            });
        }
        add_input_aliases(&mut named, outside, &deduced);
    } else if analyse_inputs {
        // The inputs are the constants a path of the kept computation starts from: those that
        // survived the pruning, as the fewest rectangles that cover exactly them.
        let mut started: Vec<CellRef> = cells
            .iter()
            .filter(|c| matches!(c.kind, CellKind::Input(_)))
            .map(|c| c.at)
            .collect();
        for comp in sccs.components() {
            for &m in comp {
                let v = cone[m as usize];
                if (v as usize) < n && matches!(kinds[v as usize], CellKind::Input(_)) {
                    started.push(ats[v as usize]);
                }
            }
        }
        started.sort_unstable();
        started.dedup();
        for rect in coalesce(&started) {
            inputs.push(Binding {
                name: addr::qualified(&model.sheets, &rect),
                rect,
            });
        }
        add_input_aliases(&mut named, outside, &started);
    }
    named.sort_by(|a, b| a.name.cmp(&b.name));
    let queryable = Queryable { names: named };
    let mut cycles = Vec::new();
    for (ci, comp) in sccs.components().enumerate() {
        if !tainted[comp[0] as usize] {
            continue;
        }
        let start = cells.len() as u32;
        for &m in comp {
            let v = cone[m as usize];
            if (v as usize) < n {
                let kind = std::mem::replace(&mut kinds[v as usize], CellKind::Const(Value::Empty));
                if matches!(kind, CellKind::Formula(_)) {
                    report.formulas_kept += 1;
                }
                cells.push(CellDef {
                    at: ats[v as usize],
                    kind,
                });
            }
        }
        if comp_cyclic[ci] {
            cycles.push((start, cells.len() as u32));
        }
    }
    // Folded formulas that ended in a blank stay formulas.
    let mut kept_blank: Vec<u32> = need_folded
        .iter()
        .copied()
        .filter(|id| blank_ids.contains(id))
        .collect();
    kept_blank.sort_unstable();
    for id in kept_blank {
        report.formulas_kept += 1;
        cells.push(CellDef {
            at: ats[id as usize],
            kind: std::mem::replace(&mut kinds[id as usize], CellKind::Const(Value::Empty)),
        });
    }
    report.cycles = comp_cyclic.iter().filter(|&&c| c).count();
    report.cells_kept = cells.len();

    // A TEXT code that is computed while the program runs may carry a language tag, which
    // LibreOffice honors and the engine has names for only in the languages of its table: the
    // code of a call that stays must be known now.
    if !spec.inspect {
        // The text a code expression has when the build can tell: a constant, a cell that is
        // one, or a join of such.
        fn code_text(
            e: &Expr,
            model: &Model,
            lookup: &dyn Fn(CellRef) -> Option<Option<Value>>,
        ) -> Option<String> {
            match e {
                Expr::Const(Value::Text(t)) => Some(t.to_string()),
                Expr::Const(_) => Some(String::new()),
                Expr::Cell(r) => match lookup(*r) {
                    Some(Some(Value::Text(t))) => Some(t.to_string()),
                    Some(Some(_)) => Some(String::new()),
                    Some(None) => None,
                    None => match model.cells.get(r).map(|c| &c.value) {
                        Some(Value::Text(t)) => Some(t.to_string()),
                        _ => Some(String::new()),
                    },
                },
                Expr::Binary(sheetrs_engine::ir::BinOp::Concat, ab) => {
                    Some(code_text(&ab[0], model, lookup)? + &code_text(&ab[1], model, lookup)?)
                }
                _ => None,
            }
        }
        let lookup = |r: CellRef| -> Option<Option<Value>> {
            let id = index.get(&r)?;
            let c = compact[id as usize];
            if c != u32::MAX && !tainted[c as usize] {
                Some(folded[id as usize].clone())
            } else {
                Some(None)
            }
        };
        let mut bad: Vec<(CellRef, String)> = Vec::new();
        // What the build cannot tell because it depends on data that changes: a warning, and an
        // error value (`#NOT_IMPLEMENTED!`) at run time if the case turns up.
        let mut unknown: Vec<(CellRef, String)> = Vec::new();
        // Where a call takes the text that is read as a regular expression when the document's
        // search mode is one.
        fn pattern_args(name: &str, n: usize) -> Vec<usize> {
            match name {
                "COUNTIF" | "SUMIF" | "AVERAGEIF" => vec![1],
                "MATCH" | "VLOOKUP" | "HLOOKUP" | "SEARCH" | "SEARCHB" => vec![0],
                "COUNTIFS" => (1..n).step_by(2).collect(),
                "SUMIFS" | "AVERAGEIFS" | "MAXIFS" | "MINIFS" => (2..n).step_by(2).collect(),
                _ => Vec::new(),
            }
        }
        let regex_mode = opts.search.mode == sheetrs_engine::options::SearchMode::Regex;
        for d in &cells {
            let CellKind::Formula(e) = &d.kind else {
                continue;
            };
            e.walk(&mut |x| {
                let Expr::Call(id, args) = x else { return };
                let name = id.def().name;
                let mut which = if regex_mode {
                    pattern_args(name, args.len())
                } else {
                    Vec::new()
                };
                // XLOOKUP reads its search text as a regular expression in match mode 3 only.
                if name == "XLOOKUP"
                    && matches!(args.get(4), Some(Expr::Const(Value::Number(n))) if *n == 3.0)
                {
                    which.push(0);
                }
                for i in which {
                    let Some(arg) = args.get(i) else { continue };
                    let Some(text) = code_text(arg, model, &lookup) else {
                        unknown.push((d.at, format!("{name} searches with a pattern that depends on data that changes, which may be a regular expression outside the subset that is reproduced")));
                        continue;
                    };
                    let rest = ["<=", ">=", "<>", "<", ">", "="]
                        .iter()
                        .find_map(|p| text.strip_prefix(p))
                        .unwrap_or(&text);
                    if sheetrs_engine::matching::may_be_regex(rest)
                        && let Err(why) = sheetrs_engine::regex_subset::translate(rest)
                    {
                        bad.push((d.at, format!("{name} searches with the regular expression `{rest}`, which has {why}")));
                    }
                }
                // XLOOKUP's binary searches over a range are not reproduced.
                if name == "XLOOKUP"
                    && matches!(args.get(1), Some(Expr::Range(_) | Expr::Cell(_)))
                    && let Some(mode) = args.get(5)
                {
                    let known = match mode {
                        Expr::Const(Value::Number(n)) => Some(*n),
                        Expr::Cell(r) => match lookup(*r) {
                            Some(Some(Value::Number(n))) => Some(n),
                            Some(_) => None,
                            None => match model.cells.get(r).map(|c| &c.value) {
                                Some(Value::Number(n)) => Some(*n),
                                _ => Some(1.0),
                            },
                        },
                        _ => None,
                    };
                    match known {
                        Some(n) if n.trunc().abs() == 2.0 => bad.push((
                            d.at,
                            "XLOOKUP search mode 2 or -2 (binary search) over a range".into(),
                        )),
                        None => unknown.push((
                            d.at,
                            "XLOOKUP has a search mode that depends on data that changes, and modes 2 and -2 over a range are not reproduced".into(),
                        )),
                        Some(_) => {}
                    }
                }
                if let Expr::Call(id, args) = x
                    && id.def().name == "TEXT"
                    && let Some(code) = args.get(1)
                {
                    match code_text(code, model, &lookup) {
                        None => unknown.push((d.at, "TEXT has a format code that depends on data that changes, which may name a language the engine has no names for".into())),
                        Some(t) if sheetrs_engine::format::has_foreign_language_tag(&t) => {
                            bad.push((d.at, format!("the code `{t}` names a language the engine has no month and day names for")));
                        }
                        Some(_) => {}
                    }
                }
            });
        }
        if !unknown.is_empty() {
            unknown.sort();
            let list: Vec<String> = unknown
                .iter()
                .take(5)
                .map(|(a, why)| {
                    let sheet = model
                        .sheets
                        .get(a.sheet as usize)
                        .map_or("?", String::as_str);
                    format!("{sheet}!{}: {why}", addr::a1(a.row, a.col))
                })
                .collect();
            report.warnings.push(format!(
                "{} call(s) use something that LibreOffice runs and the engine reproduces only in part, and the build cannot tell from the data it has whether it will be met: if it is, the cell is `#NOT_IMPLEMENTED!`; first: {}",
                unknown.len(),
                list.join("; ")
            ));
        }
        if !bad.is_empty() {
            bad.sort();
            let list: Vec<String> = bad
                .iter()
                .take(5)
                .map(|(a, why)| {
                    let sheet = model
                        .sheets
                        .get(a.sheet as usize)
                        .map_or("?", String::as_str);
                    format!("{sheet}!{}: {why}", addr::a1(a.row, a.col))
                })
                .collect();
            return err(format!(
                "{} call(s) use something that LibreOffice runs and the engine does not reproduce (a TEXT code with a language tag it has no names for, a regular expression outside the subset of ICU's syntax, XLOOKUP's binary search over a range); first: {}",
                bad.len(),
                list.join("; ")
            ));
        }
    }

    parse_diags.sort();
    parse_diags.dedup();
    diags.extend(parse_diags.into_iter().take(200));
    if report.cycles > 0 && spec.iteration.is_none() {
        diags.push(format!(
            "{} circular reference group(s) in the dependency graph; a cell that depends on itself when it is evaluated gives #CIRC! unless iteration is enabled (references the evaluation does not follow, such as an unchosen IF branch or the unselected cells of a SUMIF sum range, are not a cycle)",
            report.cycles
        ));
    }

    // The result the workbook caches for each formula the program keeps: a run that leaves a formula and
    // what it reads as they were takes it (`Engine::take_cached`). A formula that reads cells it cannot
    // name (`INDIRECT`, `OFFSET`) is calculated in every run (it is volatile for `LibreOffice` too, `no_cache`) and so is what
    // depends on it (`stale`), and what it reads holds the result the workbook caches; a program built with
    // `force_calculate_all` has none, nor has the program of an inspection.
    let cache: Vec<(u32, Value)> = match &stale_formulas {
        Some(stale)
            if !(spec.force_calculate_all || spec.inspect || spec.mode == Mode::Recalculate) =>
        {
            cells
                .iter()
                .enumerate()
                .filter(|(_, d)| {
                    matches!(d.kind, CellKind::Formula(_))
                        && !no_cache.contains(&d.at)
                        && !stale.contains(&d.at)
                })
                .filter_map(|(slot, d)| {
                    let cell = model.cells.get(&d.at)?;
                    (cell.formula.is_some() && cell.value != Value::Empty)
                        .then(|| (slot as u32, cell.value.clone()))
                })
                .collect()
        }
        _ => Vec::new(),
    };
    let meta = program_meta(model, spec, &queryable, &inputs, &outputs);
    let formula_cells = if spec.mode.is_workbook() {
        // The formulas the workbook answer writes: the ones inside the outputs (all of
        // them when nothing is declared). They come in the order of the positions.
        let declared = !spec.outputs.is_empty();
        let written: Vec<CellRef> = formula_positions
            .into_iter()
            .filter(|at| {
                Some(at.sheet) != spec.skip_sheet
                    && asked(at)
                    && (!declared || spec.whole_workbook || in_outputs(at))
            })
            .collect();
        sheetrs_engine::ir::CellRun::runs(written)
    } else {
        Vec::new()
    };
    Ok((
        Program {
            sheets: model.sheets.clone(),
            options: opts.clone(),
            cells,
            hidden_rows: model.hidden_rows.clone(),
            subtotal_cells,
            cycles,
            cache,
            inputs,
            outputs,
            mode: spec.mode,
            formula_cells,
            meta,
            queryable,
            diagnostics: diags,
            build: build_info(model, spec),
        },
        report,
    ))
}

/// The workbook with the result of every formula calculated again, as `calculateAll()` of `LibreOffice` leaves
/// it: the values a build then folds and stores as cache. The program of this calculation is the
/// recalculation of the editor (every formula an output and any constant an input, nothing folded); the
/// formulas that depend on what the engine cannot calculate keep the result the workbook caches (see
/// [`settle_found`]) as in any build.
fn calculated_model(model: &Model, spec: &Spec) -> Result<Model> {
    let all = Spec {
        mode: Mode::Recalculate,
        calculate_all: true,
        force_calculate_all: true,
        cache: Cache::default(),
        inputs: Vec::new(),
        inputs_declared: false,
        outputs: Vec::new(),
        only: Vec::new(),
        ..spec.clone()
    };
    let (program, _) = compile(model, &all)?;
    let mut engine = Engine::new(&program, program.effective_options(&Overrides::default()));
    engine.calculate();
    let mut calculated = model.clone();
    for run in &program.formula_cells {
        for at in run.cells() {
            if let Some(cell) = model.cells.get(&at) {
                let mut cell = cell.clone();
                cell.value = engine.get(at);
                calculated.cells.insert(at, cell);
            }
        }
    }
    Ok(calculated)
}

/// The formulas that read another workbook are not in the program (the parser does not read the reference): they keep the
/// result the file caches, as `LibreOffice` shows the stored value of the link (OBSERVED, both formats; the cell is not
/// updated unless the link is). What reads them is calculated from it.
fn leave_external(
    model: &Model,
    found: &[Found],
    subtotal_cells: &mut Vec<CellRef>,
    positions: &mut Vec<CellRef>,
    cached: &mut FastMap<CellRef, Value>,
) {
    let gone: FastSet<CellRef> = found
        .iter()
        .filter(|f| f.1 == Unsupported::External)
        .map(|f| f.0)
        .collect();
    if gone.is_empty() {
        return;
    }
    subtotal_cells.retain(|at| !gone.contains(at));
    positions.retain(|at| !gone.contains(at));
    for at in gone {
        let value = model.cells.get(&at).map(|c| c.value.clone());
        cached.insert(at, value.unwrap_or_default());
    }
}

/// What the build does with the formulas that depend on what the engine cannot calculate (`found`, see
/// [`Unsupported`]) once the analysis of what the run reaches has taken out of the program the ones that nothing reaches
/// (`formulas` is what stays to be calculated):
///
/// * a formula that is not in the program any more keeps the result the file caches, as `LibreOffice` does until something
///   dirties the cell; the report says so (a note per kind, and a warning), and what reads it is calculated from it;
/// * a formula that stays and that the engine answers as `LibreOffice` does (an add-in that is not installed: `#NAME?`) is
///   calculated;
/// * a formula that stays and that only `LibreOffice` can calculate (a function the engine lacks, one a macro may define) is an
///   error that names the cause: the result is not one the engine can give, and the cache is not the result of a cell that the
///   run calculates.
///
/// # Errors
///
/// When a formula that only `LibreOffice` can calculate is calculated by the run.
fn settle_found(
    model: &Model,
    found: &[Found],
    formulas: &[(CellRef, CellKind)],
    report: &mut Report,
) -> Result<()> {
    let name_of = |at: &CellRef| {
        let sheet = model
            .sheets
            .get(at.sheet as usize)
            .map_or("?", String::as_str);
        format!("{sheet}!{}", addr::a1(at.row, at.col))
    };
    let stays = |at: &CellRef| formulas.binary_search_by_key(at, |(cell, _)| *cell).is_ok();
    let mut found: Vec<&Found> = found.iter().collect();
    found.sort_by_key(|f| (f.1, f.0));
    for kind in [
        Unsupported::External,
        Unsupported::Macro,
        Unsupported::Extension,
    ] {
        let of_kind: Vec<&&Found> = found.iter().filter(|f| f.1 == kind).collect();
        let reached: Vec<&&&Found> = of_kind
            .iter()
            .filter(|f| !f.3 && f.1 != Unsupported::External && stays(&f.0))
            .collect();
        if let Some(first) = reached.first() {
            let mut what: Vec<String> = Vec::new();
            for f in &reached {
                if !what.contains(&f.2) {
                    what.push(f.2.clone());
                }
            }
            what.truncate(5);
            let (cause, why) = match kind {
                Unsupported::Macro => (
                    "call a function that a macro of the workbook may define",
                    "macros are never run, and LibreOffice gives `#MACRO?` or the result of the macro, by its security level",
                ),
                _ => (
                    "call a function that the engine does not implement (an add-in or extension of LibreOffice, or one of its functions that the engine lacks)",
                    "its result is the one only LibreOffice gives",
                ),
            };
            return Err(Error::Unsupported {
                kind,
                message: format!(
                    "{} formula(s) {cause} and the run calculates them (an edit reaches them, or every formula is calculated), so it cannot replace LibreOffice for the workbook: {why}; first: {} ({})",
                    reached.len(),
                    name_of(&first.0),
                    what.join(", ")
                ),
            });
        }
        // The ones that keep what the file caches: the cells that are not in the program.
        let kept: Vec<&&Found> = of_kind
            .iter()
            .filter(|f| f.1 == Unsupported::External || !stays(&f.0))
            .copied()
            .collect();
        let Some(first) = kept.first() else {
            continue;
        };
        let mut what: Vec<String> = Vec::new();
        for f in &kept {
            if !what.contains(&f.2) {
                what.push(f.2.clone());
            }
        }
        what.truncate(5);
        let mut cells: Vec<CellRef> = kept.iter().map(|f| f.0).collect();
        cells.dedup();
        report.from_cache.push(FromCacheNote {
            kind,
            cells: cells.len(),
            first: name_of(&first.0),
            what,
        });
        report.warnings.push(format!(
            "{} formula(s) that depend on {} keep the result the file caches (the engine cannot calculate them; LibreOffice keeps it until something dirties the cell): first {}",
            cells.len(),
            kind.as_str(),
            name_of(&first.0)
        ));
    }
    report.formulas_cached += report.from_cache.iter().map(|n| n.cells).sum::<usize>();
    Ok(())
}

/// The area of the cells of `only`, `None` when the sheet is not in the workbook.
fn only_area(model: &Model, only: &Only) -> Option<Rect> {
    let sheet = model.sheet_id(&only.sheet)?;
    let (r0, r1) = only.rows.unwrap_or((0, addr::MAX_ROWS - 1));
    let (c0, c1) = only.columns.unwrap_or((0, addr::MAX_COLS - 1));
    Some(Rect::new(sheet, r0, c0, r1, c1))
}

/// The moment of a build in UTC (`2026-10-30T12:00:00Z`): `SOURCE_DATE_EPOCH` (seconds since 1970)
/// when it is set, so that a build can be repeated exactly; else the clock. A value that is not a
/// number of seconds is an error: a build that was meant to repeat must not silently use the clock.
///
/// # Errors
///
/// When `SOURCE_DATE_EPOCH` is set and is not a number of seconds.
pub fn build_time() -> Result<String> {
    let seconds = match std::env::var("SOURCE_DATE_EPOCH") {
        Ok(value) => value
            .trim()
            .parse::<i64>()
            .map_err(|_| Error::BuildTime(value.clone()))?,
        Err(_) => sheetrs_engine::clock::unix_now().as_secs().cast_signed(),
    };
    Ok(format_utc(seconds))
}

/// `seconds` since 1970-01-01 as `YYYY-MM-DDThh:mm:ssZ`.
fn format_utc(seconds: i64) -> String {
    let (days, rest) = (seconds.div_euclid(86_400), seconds.rem_euclid(86_400));
    // Civil date from the days since 1970-01-01 (proleptic Gregorian calendar).
    let z = days + 719_468;
    let era = z.div_euclid(146_097);
    let day_of_era = z.rem_euclid(146_097);
    let year_of_era =
        (day_of_era - day_of_era / 1460 + day_of_era / 36_524 - day_of_era / 146_096) / 365;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let shifted_month = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * shifted_month + 2) / 5 + 1;
    let month = if shifted_month < 10 {
        shifted_month + 3
    } else {
        shifted_month - 9
    };
    let year = year_of_era + era * 400 + i64::from(month <= 2);
    format!(
        "{year:04}-{month:02}-{day:02}T{:02}:{:02}:{:02}Z",
        rest / 3600,
        rest % 3600 / 60,
        rest % 60
    )
}

/// Most rectangles the formula cells of a workbook may be outputs as, before the outputs
/// become the used area of the formulas of each sheet.
const MAX_FORMULA_OUTPUTS: usize = 4096;

/// The rectangles that hold exactly the cells `at` (sorted by position): the runs of consecutive
/// cells of a row, and one under another when two rows have the same run.
fn formula_rects(at: &[CellRef]) -> Vec<Rect> {
    let mut rects: Vec<Rect> = Vec::new();
    // The rectangle that ends in the row before, by sheet and columns.
    let mut open: FastMap<(u16, u32, u32), usize> = FastMap::default();
    let mut i = 0;
    while i < at.len() {
        let first = at[i];
        let mut last = first.col;
        let mut j = i + 1;
        while j < at.len()
            && at[j].sheet == first.sheet
            && at[j].row == first.row
            && at[j].col == last + 1
        {
            last = at[j].col;
            j += 1;
        }
        let key = (first.sheet, first.col, last);
        match open.get(&key).copied() {
            Some(k) if first.row > 0 && rects[k].r1 + 1 == first.row => rects[k].r1 = first.row,
            _ => {
                open.insert(key, rects.len());
                rects.push(Rect::new(
                    first.sheet,
                    first.row,
                    first.col,
                    first.row,
                    last,
                ));
            }
        }
        i = j;
    }
    rects.sort_unstable_by_key(|r| (r.sheet, r.r0, r.c0));
    rects
}

/// How the program is built: the compiler, the options that shape it and the workbook. The time
/// is left empty: the command that builds sets it with [`build_time`].
fn build_info(model: &Model, spec: &Spec) -> BuildInfo {
    BuildInfo {
        compiler: format!("sheetrs-compiler {}", env!("CARGO_PKG_VERSION")),
        target: format!("{}-{}", std::env::consts::ARCH, std::env::consts::OS),
        built_at: String::new(),
        options: BuildOptions {
            outputs: if spec.outputs.is_empty() {
                "deduced"
            } else {
                "declared"
            }
            .to_string(),
            inputs: if spec.inputs.is_empty() {
                "deduced"
            } else {
                "declared"
            }
            .to_string(),
            locale: spec
                .locale
                .as_ref()
                .map_or_else(String::new, |l| l.tag.clone()),
            fold: spec.fold,
            calculate_all: spec.calculate_all,
            force_calculate_all: spec.force_calculate_all,
            features: if spec.features.longtext {
                vec!["LONGTEXT".to_string()]
            } else {
                Vec::new()
            },
        },
        source: model.source.clone(),
    }
}

/// What `--meta` lists, in this order: the properties the document declares (`document`), the rows
/// of the settings sheet, the inputs and the outputs that were not declared (so that what the
/// program exposes is always listed: `input` and `output`, the value is the range), and the defined
/// names that can be used as aliases (`name`).
fn program_meta(
    model: &Model,
    spec: &Spec,
    queryable: &Queryable,
    inputs: &[Binding],
    outputs: &[Binding],
) -> Vec<Meta> {
    let entry = |key: &str, value: &str, meta_type: &str| Meta {
        key: key.to_string(),
        value: serde_json::Value::String(value.to_string()),
        description: String::new(),
        meta_type: meta_type.to_string(),
    };
    let mut meta: Vec<Meta> = model
        .properties
        .iter()
        .map(|(key, value)| entry(key, value, "document"))
        .collect();
    meta.extend(spec.meta.iter().cloned());
    for (declared, list, kind) in [
        (
            !spec.inputs.is_empty() || spec.inputs_declared,
            inputs,
            "input",
        ),
        (!spec.outputs.is_empty(), outputs, "output"),
    ] {
        // A declared one is in the settings already (a row of the sheet, an entry of the file).
        if !declared || !spec.meta.iter().any(|m| m.meta_type == kind) {
            meta.extend(list.iter().map(|b| {
                let range = addr::qualified(&model.sheets, &b.rect);
                entry(&b.name, &range, kind)
            }));
        }
    }
    let mut names: Vec<&String> = queryable.names.iter().map(|b| &b.name).collect();
    names.sort();
    meta.extend(
        names
            .into_iter()
            .filter_map(|name| Some(entry(name, model.names.get(name)?, "name"))),
    );
    meta
}

/// The names that point inside the inputs are aliases of them: those whose every cell is one.
fn add_input_aliases(named: &mut Vec<Binding>, outside: Vec<Binding>, inputs: &[CellRef]) {
    for b in outside {
        let all_inputs = b.rect.area() <= inputs.len() as u64
            && (b.rect.r0..=b.rect.r1).all(|r| {
                (b.rect.c0..=b.rect.c1).all(|c| {
                    inputs
                        .binary_search(&CellRef::new(b.rect.sheet, r, c))
                        .is_ok()
                })
            });
        if all_inputs {
            named.push(b);
        }
    }
}

/// The constants a path of the dependency graph starts from, in order: those a formula reads (a
/// reference or a range) and those inside a defined name or a table, which are exposed under that
/// name. A constant that no formula reads and no name covers starts no path, so it is not an input
/// (it may still be an output). A workbook that resolves references while it runs (`INDIRECT`,
/// `OFFSET`) can read any cell, so there every constant counts.
fn constants_that_start_a_path(
    model: &Model,
    formulas: &[(CellRef, CellKind)],
    skip_sheet: Option<u16>,
) -> Vec<CellRef> {
    let key =
        |at: CellRef| (u64::from(at.sheet) << 34) | (u64::from(at.row) << 14) | u64::from(at.col);
    let unkey = |k: u64| {
        CellRef::new(
            (k >> 34) as u16,
            ((k >> 14) & 0xF_FFFF) as u32,
            (k & 0x3FFF) as u32,
        )
    };
    let is_constant = |at: &CellRef| {
        Some(at.sheet) != skip_sheet
            && model
                .cells
                .get(at)
                .is_some_and(|c| c.formula.is_none() && !matches!(c.value, Value::Empty))
    };
    // The references to one cell (almost all of them) go straight to a list of positions, and only
    // the ranges are kept as rectangles: a set of millions of rectangles is the cost of this pass.
    let mut keys: Vec<u64> = Vec::new();
    let mut rects: FastSet<Rect> = FastSet::default();
    let mut dynamic = false;
    let mut deps: Vec<Rect> = Vec::new();
    for (_, kind) in formulas {
        if let CellKind::Formula(e) = kind {
            deps.clear();
            dynamic |= dependencies_into(e, &mut deps).dynamic;
            for r in &deps {
                if r.r0 == r.r1 && r.c0 == r.c1 {
                    keys.push(key(CellRef::new(r.sheet, r.r0, r.c0)));
                } else {
                    rects.insert(*r);
                }
            }
        }
    }
    for name in model.names.keys() {
        if let Ok(rect) = resolve_target(model, name)
            && rect.area() <= MAX_BINDING_CELLS
        {
            rects.insert(rect);
        }
    }
    let total = model.cells.len() as u64;
    if dynamic {
        keys.clear();
        keys.extend(model.cells.keys().filter(|at| is_constant(at)).map(key));
    } else {
        keys.sort_unstable();
        keys.dedup();
        keys.retain(|k| is_constant(&unkey(*k)));
        for r in &rects {
            if r.area() <= total * 4 + 64 {
                for row in r.r0..=r.r1 {
                    for col in r.c0..=r.c1 {
                        let at = CellRef::new(r.sheet, row, col);
                        if is_constant(&at) {
                            keys.push(key(at));
                        }
                    }
                }
            } else {
                keys.extend(
                    model
                        .cells
                        .keys()
                        .filter(|at| r.contains(*at) && is_constant(at))
                        .map(key),
                );
            }
        }
    }
    keys.sort_unstable();
    keys.dedup();
    keys.into_iter().map(unkey).collect()
}

/// The fewest rectangles that cover exactly the cells (sorted, without repeats): the runs of a row,
/// joined with the run of the row above that has the same columns.
fn coalesce(cells: &[CellRef]) -> Vec<Rect> {
    use std::collections::HashMap;
    let mut rects: Vec<Rect> = Vec::new();
    // The rectangle that ends in the previous row, by sheet and columns.
    let mut open: HashMap<(u16, u32, u32), usize> = HashMap::new();
    let mut i = 0;
    while i < cells.len() {
        let (sheet, row, c0) = (cells[i].sheet, cells[i].row, cells[i].col);
        let mut c1 = c0;
        let mut j = i + 1;
        while j < cells.len()
            && cells[j].sheet == sheet
            && cells[j].row == row
            && cells[j].col == c1 + 1
        {
            c1 += 1;
            j += 1;
        }
        match open.get(&(sheet, c0, c1)).copied() {
            Some(k) if rects[k].r1 + 1 == row => rects[k].r1 = row,
            _ => {
                open.insert((sheet, c0, c1), rects.len());
                rects.push(Rect::new(sheet, row, c0, row, c1));
            }
        }
        i = j;
    }
    rects.sort_by_key(|r| (r.sheet, r.r0, r.c0));
    rects
}

/// The workbook's constants, read in place (see `ConstSource`).
struct ConstView<'m> {
    model: &'m Model,
    cached: &'m FastMap<CellRef, Value>,
    extent: Vec<(u32, u32)>,
}

impl<'m> ConstView<'m> {
    fn new(model: &'m Model, cached: &'m FastMap<CellRef, Value>) -> Self {
        let mut extent = vec![(0u32, 0u32); model.sheets.len().max(1)];
        for at in model.cells.keys() {
            let e = &mut extent[at.sheet as usize];
            e.0 = e.0.max(at.row + 1);
            e.1 = e.1.max(at.col + 1);
        }
        Self {
            model,
            cached,
            extent,
        }
    }

    /// Positions of constants inside `r` that are not graph nodes.
    fn consts_in(&self, r: &Rect, nodes: &NodeIndex, out: &mut Vec<CellRef>) {
        let total = self.model.cells.len() as u64;
        if r.area() <= total * 4 + 64 {
            for row in r.r0..=r.r1 {
                for col in r.c0..=r.c1 {
                    let at = CellRef::new(r.sheet, row, col);
                    if !nodes.contains(&at) && self.get(at).is_some() {
                        out.push(at);
                    }
                }
            }
        } else {
            out.extend(
                self.model
                    .cells
                    .keys()
                    .filter(|at| r.contains(*at) && !nodes.contains(at))
                    .filter(|at| self.get(*at).is_some()),
            );
        }
    }

    /// Every constant, in `out`: first those of the model, which come in the order of the
    /// positions, then the cached results of folded formulas, which come in any order. Returns where
    /// the second run starts.
    fn all_consts(&self, nodes: &NodeIndex, out: &mut Vec<CellRef>) -> usize {
        out.extend(
            self.model
                .cells
                .keys()
                .filter(|at| !nodes.contains(at) && self.get(*at).is_some()),
        );
        let ordered = out.len();
        out.extend(self.cached.keys().filter(|at| !nodes.contains(at)));
        ordered
    }
}

impl ConstSource for ConstView<'_> {
    fn get(&self, at: CellRef) -> Option<Value> {
        if let Some(v) = self.cached.get(&at) {
            return Some(v.clone());
        }
        self.model
            .cells
            .get(&at)
            .filter(|c| c.formula.is_none())
            .map(|c| c.value.clone())
    }

    fn extent(&self, sheet: u16) -> (u32, u32) {
        self.extent.get(sheet as usize).copied().unwrap_or((0, 0))
    }
}

/// Where the cells that are graph nodes sit: node `i` is at `ats[i]`. The nodes come in order of
/// position (sheet, row, column), so those of a sheet are consecutive and those of a row too: a
/// cell is found by the first node of its row and a short search along the row, with no hash table
/// (four bytes a node less than one would take, and no hashing of every position to build it).
struct NodeIndex<'a> {
    ats: &'a [CellRef],
    sheets: Vec<SheetNodes>,
}

#[derive(Clone, Default)]
struct SheetNodes {
    /// The nodes of the sheet are `start..end`.
    start: u32,
    end: u32,
    /// `rows[r]` is the first node at row `r` or after, up to one past the last row of the sheet.
    /// Empty for a sheet whose nodes lie so far apart that the table would outweigh them: a search
    /// over the sheet's nodes does then.
    rows: Vec<u32>,
}

impl<'a> NodeIndex<'a> {
    /// `ats` must be sorted by position, without repeats.
    fn new(ats: &'a [CellRef], sheets: usize) -> Self {
        debug_assert!(ats.windows(2).all(|w| w[0] < w[1]));
        let mut out = vec![SheetNodes::default(); sheets.max(1)];
        let mut i = 0;
        while i < ats.len() {
            let sheet = ats[i].sheet;
            let mut j = i;
            while j < ats.len() && ats[j].sheet == sheet {
                j += 1;
            }
            let last_row = ats[j - 1].row as usize;
            let mut rows = Vec::new();
            if last_row <= 8 * (j - i) + 4096 {
                rows.reserve_exact(last_row + 2);
                let mut k = i;
                for row in 0..=last_row + 1 {
                    while k < j && (ats[k].row as usize) < row {
                        k += 1;
                    }
                    rows.push(k as u32);
                }
            }
            out[usize::from(sheet)] = SheetNodes {
                start: i as u32,
                end: j as u32,
                rows,
            };
            i = j;
        }
        Self { ats, sheets: out }
    }

    /// The node at `at`, if the cell is one.
    fn get(&self, at: &CellRef) -> Option<u32> {
        let sheet = self.sheets.get(usize::from(at.sheet))?;
        let (lo, hi) = if sheet.rows.is_empty() {
            (sheet.start, sheet.end)
        } else {
            let row = at.row as usize;
            if row + 1 >= sheet.rows.len() {
                return None;
            }
            (sheet.rows[row], sheet.rows[row + 1])
        };
        let found = self.ats[lo as usize..hi as usize]
            .binary_search_by_key(&(at.row, at.col), |c| (c.row, c.col))
            .ok()?;
        Some(lo + found as u32)
    }

    fn contains(&self, at: &CellRef) -> bool {
        self.get(at).is_some()
    }

    /// Existing cells inside `r`, in order of position.
    fn members(&self, r: &Rect, out: &mut Vec<u32>) {
        let Some(sheet) = self.sheets.get(usize::from(r.sheet)) else {
            return;
        };
        if !sheet.rows.is_empty() {
            // Row by row: the cost follows the rows of `r` (one for a cell), not the nodes of the sheet.
            let last_row = sheet.rows.len() - 2;
            if r.r0 as usize > last_row {
                return;
            }
            for row in r.r0 as usize..=(r.r1 as usize).min(last_row) {
                let (lo, hi) = (sheet.rows[row] as usize, sheet.rows[row + 1] as usize);
                let nodes = &self.ats[lo..hi];
                let from = nodes.partition_point(|c| c.col < r.c0);
                out.extend(
                    nodes[from..]
                        .iter()
                        .take_while(|c| c.col <= r.c1)
                        .enumerate()
                        .map(|(k, _)| (lo + from + k) as u32),
                );
            }
            return;
        }
        let nodes = &self.ats[sheet.start as usize..sheet.end as usize];
        // A sparse sheet: the nodes of the rows of `r`, found by two searches, and among them those in its columns.
        let first = nodes.partition_point(|c| c.row < r.r0);
        let last = first + nodes[first..].partition_point(|c| c.row <= r.r1);
        out.extend(
            (first..last)
                .filter(|&i| (r.c0..=r.c1).contains(&nodes[i].col))
                .map(|i| sheet.start + i as u32),
        );
    }
}

/// Lazily-built dependency adjacency over cell nodes (`0..n`) and shared
/// rectangle nodes (`n..`).
struct Lazy<'a> {
    index: &'a NodeIndex<'a>,
    kinds: &'a [CellKind],
    adj: Vec<Option<Vec<u32>>>,
    rect_ids: FastMap<Rect, u32>,
    rect_members: Vec<Vec<u32>>,
    /// What the functions of each formula (node `0..n`) call; the default for one not read yet.
    flags: Vec<Flags>,
    /// The references of the formula being read, reused by every formula.
    deps: Vec<Rect>,
}

impl Lazy<'_> {
    fn n(&self) -> u32 {
        self.kinds.len() as u32
    }

    fn seen(visited: &mut Vec<bool>, v: u32) -> bool {
        if visited.len() <= v as usize {
            visited.resize(v as usize + 1, false);
        }
        std::mem::replace(&mut visited[v as usize], true)
    }

    fn members(&self, r: &Rect, out: &mut Vec<u32>) {
        self.index.members(r, out);
    }

    fn rect_node(&mut self, r: Rect) -> Option<u32> {
        if let Some(&id) = self.rect_ids.get(&r) {
            return Some(id);
        }
        let mut m = Vec::new();
        self.members(&r, &mut m);
        match m.len() {
            0 => None,
            1 => Some(m[0]),
            _ => {
                let id = self.n() + self.rect_members.len() as u32;
                self.rect_members.push(m);
                self.adj.push(None);
                self.rect_ids.insert(r, id);
                Some(id)
            }
        }
    }

    /// The adjacency of a node whose `adj_of` has been called already.
    fn adj_known(&self, v: u32) -> &[u32] {
        if v >= self.n() {
            return &self.rect_members[(v - self.n()) as usize];
        }
        self.adj[v as usize].as_deref().unwrap_or(&[])
    }

    fn adj_of(&mut self, v: u32) -> &[u32] {
        // The members of a range node are already stored: no copy.
        if v >= self.n() {
            return &self.rect_members[(v - self.n()) as usize];
        }
        if self.adj.len() <= v as usize {
            self.adj.resize(v as usize + 1, None);
        }
        if self.adj[v as usize].is_none() {
            let deps = match &self.kinds[v as usize] {
                CellKind::Formula(e) => {
                    let mut deps = std::mem::take(&mut self.deps);
                    deps.clear();
                    let flags = dependencies_into(e, &mut deps);
                    self.flags[v as usize] = flags;
                    let mut out = Vec::with_capacity(deps.len());
                    for &r in &deps {
                        // One cell is its own node: no rectangle to share.
                        if r.is_single() {
                            out.extend(self.index.get(&r.top_left()));
                        } else {
                            out.extend(self.rect_node(r));
                        }
                    }
                    self.deps = deps;
                    out.sort_unstable();
                    out.dedup();
                    out.shrink_to_fit();
                    out
                }
                _ => Vec::new(),
            };
            self.adj[v as usize] = Some(deps);
        }
        self.adj[v as usize].as_deref().unwrap_or(&[])
    }
}

/// The functions whose result depends on where a cell is, not on what it holds, and that `LibreOffice` calculates
/// again when the cell moves (it keeps the result of `ROWS`, `COLUMNS`, `AREAS`, `SHEET`, which ask about the size
/// of a range or the sheet).
const POSITIONAL: [&str; 2] = ["ROW", "COLUMN"];

/// What a formula asks of another cell that is answered while it is lowered (`FORMULA`, `ISFORMULA`,
/// `FORMULATEXT` of a literal reference): the formula keeps no reference to the cell, so the text says it.
const ABOUT_FORMULAS: [&str; 3] = ["formula", "isformula", "formulatext"];

fn overlaps(a: &Rect, b: &Rect) -> bool {
    a.sheet == b.sheet && a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1
}

/// A formula as its cell reads it: the references that are not absolute as the distance to the cell, so that two cells of a column
/// whose formulas read the same cells relative to themselves have the same text (`SUM(A1:A3)` in `D4` and `SUM(A2:A4)` in `D5`).
fn same_shape(text: &str, at: CellRef) -> Option<String> {
    use sheetrs_formula::ast::RefArea;
    let mut ast = sheetrs_formula::parse(text.trim_start_matches('=')).ok()?;
    let (row, col) = (at.row, at.col);
    ast.walk_mut(&mut |x| {
        if let sheetrs_formula::Expr::Ref(r) = x {
            let cell = |a: &mut sheetrs_formula::ast::CellAddr| {
                if !a.row_abs {
                    a.row = a.row.wrapping_sub(row);
                }
                if !a.col_abs {
                    a.col = a.col.wrapping_sub(col);
                }
            };
            match &mut r.area {
                RefArea::Cell(a) => cell(a),
                RefArea::Range(a, b) => {
                    cell(a);
                    cell(b);
                }
                RefArea::Cols(a, b) => {
                    for c in [a, b] {
                        if !c.abs {
                            c.col = c.col.wrapping_sub(col);
                        }
                    }
                }
                RefArea::Rows(a, b) => {
                    for c in [a, b] {
                        if !c.abs {
                            c.row = c.row.wrapping_sub(row);
                        }
                    }
                }
            }
        }
    });
    Some(format!("{ast:?}"))
}

/// Whether the formula of `at` has, in the cell above or below it, a formula of the same shape (`LibreOffice` makes a group of
/// the adjacent formulas of a column that read the same relative cells, and a group listens to the area it reads as one).
fn in_a_group(model: &Model, at: CellRef) -> bool {
    let text = |at: CellRef| {
        model
            .cells
            .get(&at)
            .and_then(|c| c.formula.as_deref())
            .filter(|f| !f.is_empty())
            .and_then(|f| same_shape(f, at))
    };
    let Some(mine) = text(at) else {
        return false;
    };
    let neighbor = |row: Option<u32>| {
        row.is_some_and(|row| text(CellRef::new(at.sheet, row, at.col)).as_ref() == Some(&mine))
    };
    neighbor(at.row.checked_sub(1)) || neighbor(at.row.checked_add(1))
}

/// Whether a formula that asks what a cell holds (`ISFORMULA`, `FORMULA`, `FORMULATEXT`) asks about a cell of `changed`: the
/// references it is given overlap one, or it is given something that is not a reference (any cell may be the one).
fn asks_about_a_changed_cell(text: &str, sheet: u16, model: &Model, changed: &[Rect]) -> bool {
    use sheetrs_formula::ast::RefArea;
    let Ok(ast) = sheetrs_formula::parse(text.trim_start_matches('=')) else {
        return true;
    };
    let mut stale = false;
    ast.walk(&mut |x| {
        let sheetrs_formula::Expr::Call(call) = x else {
            return;
        };
        let name = call.name.to_lowercase();
        if !ABOUT_FORMULAS.contains(&name.trim_start_matches("_xlfn.")) {
            return;
        }
        for arg in &call.args {
            match arg {
                Some(sheetrs_formula::Expr::Ref(r)) => {
                    let sheet = match &r.sheet {
                        Some(name) => model.sheet_id(name),
                        None => Some(sheet),
                    };
                    let Some(sheet) = sheet else {
                        stale = true;
                        continue;
                    };
                    let (r0, c0, r1, c1) = match r.area {
                        RefArea::Cell(a) => (a.row, a.col, a.row, a.col),
                        RefArea::Range(a, b) => (
                            a.row.min(b.row),
                            a.col.min(b.col),
                            a.row.max(b.row),
                            a.col.max(b.col),
                        ),
                        RefArea::Cols(a, b) => {
                            (0, a.col.min(b.col), addr::MAX_ROWS - 1, a.col.max(b.col))
                        }
                        RefArea::Rows(a, b) => {
                            (a.row.min(b.row), 0, a.row.max(b.row), addr::MAX_COLS - 1)
                        }
                    };
                    let asked = Rect::new(sheet, r0, c0, r1, c1);
                    stale |= changed.iter().any(|c| overlaps(c, &asked));
                }
                _ => stale = true,
            }
        }
    });
    stale
}

/// Whether `text` names one of `names` (in lower case), however it is written.
fn mentions(text: &str, names: &HashSet<String>, buffer: &mut String) -> bool {
    text.split(|c: char| !(c.is_alphanumeric() || c == '_' || c == '.'))
        .filter(|word| !word.is_empty())
        .any(|word| {
            buffer.clear();
            buffer.extend(word.chars().flat_map(char::to_lowercase));
            names.contains(buffer.as_str())
                || buffer
                    .strip_prefix("_xlfn.")
                    .is_some_and(|word| names.contains(word))
        })
}

/// What [`keep_cached`] found.
struct Kept {
    /// How many formulas keep their result, none when the whole workbook is calculated.
    count: usize,
    /// The formulas whose cached result is stale (the ones a change reaches, and what depends on them), which
    /// a program must not keep as a result. `None` when the analysis did not run (the workbook is
    /// calculated whole, or reads cells it cannot name): the program keeps no result at all.
    stale: Option<FastSet<CellRef>>,
}

/// Cuts the formulas to those a recalculation has to calculate (see [`Cache`]): the others become
/// constants that hold the result the workbook caches (`cached`) and are no longer outputs
/// (`positions`), when `cut`; without it nothing is cut and only the stale formulas are found.
///
/// The ones to calculate are the stale ones and the formulas that depend on them, directly or
/// through a range, found on the graph of the formulas. A change to a cell that is no formula
/// reaches the formulas that read it, and those are where the walk starts. The declared `inputs` are
/// cells that may change when the program runs: the formulas that read them are kept (they have a
/// result to keep, and a run that leaves the inputs as they are takes it), but their results are not
/// stale.
#[allow(clippy::too_many_arguments)] // the state of the build, which the pieces of it share
fn keep_cached(
    model: &Model,
    cache: &Cache,
    inputs: &[Rect],
    cut: bool,
    keep_positions: bool,
    texts: &[(CellRef, &ModelCell)],
    exempt: &FastSet<CellRef>,
    formulas: &mut Vec<(CellRef, CellKind)>,
    cached: &mut FastMap<CellRef, Value>,
    positions: &mut Vec<CellRef>,
) -> Kept {
    let n = formulas.len();
    let nothing = Kept {
        count: 0,
        stale: None,
    };
    let changed: Vec<Rect> = cache
        .changed
        .iter()
        .filter_map(|(sheet, r0, c0, r1, c1)| {
            Some(Rect::new(model.sheet_id(sheet)?, *r0, *c0, *r1, *c1))
        })
        .collect();
    // Where rows or columns went: the sheets that lost some (the whole columns or rows of one are stale), and
    // the first row or column that went in each (what is at or after it moved: `ROW` and `COLUMN` of it).
    let mut rows_deleted: Vec<u16> = Vec::new();
    let mut columns_deleted: Vec<u16> = Vec::new();
    let mut first_row: FastMap<u16, u32> = FastMap::default();
    let mut first_column: FastMap<u16, u32> = FastMap::default();
    for (bands, sheets, first) in [
        (&cache.rows_deleted, &mut rows_deleted, &mut first_row),
        (
            &cache.columns_deleted,
            &mut columns_deleted,
            &mut first_column,
        ),
    ] {
        for (sheet, at, _) in bands {
            if let Some(id) = model.sheet_id(sheet) {
                sheets.push(id);
                let slot = first.entry(id).or_insert(*at);
                *slot = (*slot).min(*at);
            }
        }
    }
    // A deleted sheet: `LibreOffice` calculates every formula again.
    if cache.sheets_deleted {
        return nothing;
    }
    // The words that make a formula stale by its text: the names it reads through that the deletions reached,
    // and, whenever a cell changed, the functions that ask what a cell holds.
    let mut names: HashSet<String> = cache
        .stale_names
        .iter()
        .map(|name| name.to_lowercase())
        .collect();
    // A deleted name recompiles every formula that reads through a name.
    if cache.names_deleted {
        names.extend(model.names.keys().map(|name| name.to_lowercase()));
    }
    // The cells that changed, and the ones a calculation is given: what `ISFORMULA`, `FORMULA` and `FORMULATEXT` ask about (they are
    // answered while the formula is lowered, so they keep no reference to the cell).
    let changed_or_given: Vec<Rect> = changed.iter().chain(inputs).copied().collect();
    let about: HashSet<String> = if changed_or_given.is_empty() {
        HashSet::new()
    } else {
        ABOUT_FORMULAS.map(String::from).into_iter().collect()
    };
    // `FORMULATEXT` (`FORMULA` of an ODS) is calculated by `LibreOffice` when it opens any file (OBSERVED, both formats; `ISFORMULA` is not).
    // It is answered while the formula is lowered, so the call is not seen: the text says it.
    let on_load_words: HashSet<String> = ["formula", "formulatext"]
        .map(String::from)
        .into_iter()
        .collect();
    // What the deletions moved: the cell or range (of the workbook after them) is at or after the first row or
    // column that went in its sheet.
    let moved = |r: &Rect| {
        first_row.get(&r.sheet).is_some_and(|first| r.r1 >= *first)
            || first_column
                .get(&r.sheet)
                .is_some_and(|first| r.c1 >= *first)
    };
    // The cell of the model that holds each formula. Both lists are in the order of the cells; `formulas` lacks the
    // ones that were taken from the cache or kept as the value of the file, so the positions are matched, not the indexes.
    let mut cell_of: Vec<Option<&ModelCell>> = Vec::with_capacity(n);
    let mut next = 0;
    for (at, _) in formulas.iter() {
        while texts.get(next).is_some_and(|(text_at, _)| text_at < at) {
            next += 1;
        }
        cell_of.push(
            texts
                .get(next)
                .filter(|(text_at, _)| text_at == at)
                .map(|(_, c)| *c),
        );
    }
    let hidden: Vec<Rect> = cache
        .rows_hidden
        .iter()
        .filter_map(|(sheet, first, last)| {
            Some(Rect::new(
                model.sheet_id(sheet)?,
                *first,
                0,
                *last,
                addr::MAX_COLS - 1,
            ))
        })
        .collect();
    let no_result = |i: usize| cell_of[i].is_none_or(|cell| cell.value == Value::Empty);
    let (mut deps, mut buffer) = (Vec::new(), String::new());
    // `hard`: the result is stale. `soft`: the formula reads an input, whose value may change.
    let mut hard = vec![false; n];
    let mut soft = vec![false; n];
    // The formulas `LibreOffice` calculates when it loads the file (an ODS: see `ModelCell::recalculated_on_load`): they are
    // calculated, and what reads them is not (OBSERVED: a formula that caches a result of a type it keeps shows it, whatever
    // the cell it reads gives), so they are no start of the walk.
    // The calls the engine cannot answer (`exempt`) are not: their result is what `LibreOffice` calculates with what only it
    // has (a macro, a function the engine lacks), which the file caches.
    let on_load: Vec<bool> = cell_of
        .iter()
        .zip(formulas.iter())
        .map(|(cell, (at, _))| {
            cell.is_some_and(|cell| cell.recalculated_on_load) && !exempt.contains(at)
        })
        .collect();
    for (i, (at, kind)) in formulas.iter().enumerate() {
        let CellKind::Formula(e) = kind else {
            continue;
        };
        deps.clear();
        let flags = dependencies_into(e, &mut deps);
        // `INDIRECT` and `OFFSET` that read a cell the compiler cannot name are volatile for `LibreOffice` too: it calculates them
        // when it opens the file and what reads them with them, and what they read keeps the result the file caches (the cells the
        // cut leaves hold it, and a dynamic read of one finds it).
        let mut stale = flags.volatile
            || flags.on_load
            || flags.dynamic
            || no_result(i)
            || changed
                .iter()
                .any(|r| r.contains(*at) || deps.iter().any(|d| overlaps(d, r)))
            || deps.iter().any(|d| {
                d.r1 == addr::MAX_ROWS - 1 && rows_deleted.contains(&d.sheet)
                    || d.c1 == addr::MAX_COLS - 1 && columns_deleted.contains(&d.sheet)
            });
        // Hidden or shown rows (OBSERVED, `write_to/probe_hide_groups.py`, both formats): `LibreOffice` calculates a `SUBTOTAL` that
        // reads a range with one of them, and a formula that reads such a range (not a single cell) when it is in a group of
        // adjacent formulas of one shape; any other formula keeps its result.
        // In an ODS (not in an XLSX, OBSERVED, `visibility.{xlsx,ods}`: the same formulas give another table) a `SUBTOTAL` that reads a
        // single cell of a row that flipped is calculated too.
        if !(stale || hidden.is_empty()) {
            let (mut reads, mut subtotal, mut reads_a_cell) = (false, false, false);
            e.walk(&mut |x| match x {
                Expr::Range(r) => reads |= hidden.iter().any(|h| overlaps(r, h)),
                Expr::Cell(c) => reads_a_cell |= hidden.iter().any(|h| h.contains(*c)),
                Expr::Call(id, _) => subtotal |= id.def().name == "SUBTOTAL",
                _ => {}
            });
            stale = reads && (subtotal || in_a_group(model, *at))
                || !model.from_xlsx && subtotal && reads_a_cell;
        }
        // `ROW` and `COLUMN` give where a cell is: they are calculated again when the formula's own cell moved
        // or the cell or range they are given did (`LibreOffice` keeps the result of `ROWS`, `COLUMNS`,
        // `ISFORMULA` and the rest of what asks about cells when they only move).
        if !(stale || first_row.is_empty() && first_column.is_empty()) {
            e.walk(&mut |x| {
                if let Expr::Call(id, args) = x
                    && POSITIONAL.contains(&id.def().name)
                {
                    stale |= moved(&Rect::single(*at));
                    for arg in args {
                        arg.walk(&mut |y| match y {
                            Expr::Cell(c) => stale |= moved(&Rect::single(*c)),
                            Expr::Range(r) => stale |= moved(r),
                            _ => {}
                        });
                    }
                }
            });
        }
        // A formula of an XLSX that holds a name LibreOffice cannot resolve (a bare `#REF!` that a deletion or a shifted
        // shared formula left, an undefined name) and caches an error is calculated when the file is opened: it shows
        // that error of the name (`oracle`, `shared_beyond`). One that caches a number keeps it.
        if !stale
            && model.from_xlsx
            && let Some(cell) = cell_of[i]
            && matches!(cell.value, Value::Error(_))
        {
            e.walk(&mut |x| {
                stale |= matches!(
                    x,
                    Expr::Const(Value::Error(sheetrs_engine::value::ErrorKind::Name))
                );
            });
        }
        if !stale
            && cache.names_deleted
            && model.from_xlsx
            && let Some(cell) = cell_of[i]
            && let Some(text) = &cell.formula
        {
            // A bare one: `Data!#REF!` (a qualified error, as Excel writes it) is not read as a name.
            let bytes = text.as_bytes();
            stale = bytes.windows(5).enumerate().any(|(at, w)| {
                w.eq_ignore_ascii_case(b"#REF!") && (at == 0 || bytes[at - 1] != b'!')
            });
        }
        if !stale
            && !names.is_empty()
            && let Some(cell) = cell_of[i]
            && let Some(text) = &cell.formula
        {
            stale = mentions(text, &names, &mut buffer);
        }
        // A formula that asks what a cell holds is stale when that cell is one that changed (the formula keeps no
        // reference to it: the text says which).
        if !stale
            && let Some(cell) = cell_of[i]
            && let Some(text) = &cell.formula
            && mentions(text, &on_load_words, &mut buffer)
        {
            stale = true;
        }
        if !stale
            && !about.is_empty()
            && let Some(cell) = cell_of[i]
            && let Some(text) = &cell.formula
            && mentions(text, &about, &mut buffer)
        {
            stale = asks_about_a_changed_cell(text, at.sheet, model, &changed_or_given);
        }
        hard[i] = stale;
        soft[i] = !stale
            && inputs
                .iter()
                .any(|r| r.contains(*at) || deps.iter().any(|d| overlaps(d, r)));
    }
    let all = std::mem::take(formulas);
    let (ats, kinds): (Vec<CellRef>, Vec<CellKind>) = all.into_iter().unzip();
    let index = NodeIndex::new(&ats, model.sheets.len());
    // The cells of an array formula are calculated together: the writer drops the result of the
    // first one only, and what is stale for one is for all.
    let group = |at: &CellRef, offset: &(u32, u32)| -> Option<(usize, usize)> {
        let master = CellRef::new(
            at.sheet,
            at.row.checked_sub(offset.0)?,
            at.col.checked_sub(offset.1)?,
        );
        Some((index.get(at)? as usize, index.get(&master)? as usize))
    };
    for roots in [&mut hard, &mut soft] {
        for (at, offset) in &model.array_members {
            if let Some((member, master)) = group(at, offset) {
                roots[master] |= roots[member];
            }
        }
        for (at, offset) in &model.array_members {
            if let Some((member, master)) = group(at, offset) {
                roots[member] = roots[master];
            }
        }
    }
    let wanted = (hard.iter().any(|&r| r), soft.iter().any(|&r| r) && cut);
    let reach = |reached: &mut Vec<bool>, readers: &[u32], start: &[u32], roots: &[bool]| {
        let mut stack: Vec<u32> = (0..n as u32).filter(|&v| roots[v as usize]).collect();
        while let Some(v) = stack.pop() {
            if !std::mem::replace(&mut reached[v as usize], true) {
                let (from, to) = (start[v as usize] as usize, start[v as usize + 1] as usize);
                stack.extend_from_slice(&readers[from..to]);
            }
        }
    };
    let (stale_reached, reached);
    if wanted == (false, false) {
        // Nothing reaches anything: no graph to build.
        stale_reached = vec![false; n];
        reached = vec![false; n];
    } else {
        let mut graph = Lazy {
            index: &index,
            kinds: &kinds,
            adj: vec![None; n],
            rect_ids: FastMap::default(),
            rect_members: Vec::new(),
            flags: vec![Flags::default(); n],
            deps: Vec::new(),
        };
        for v in 0..n as u32 {
            graph.adj_of(v);
        }
        // Who reads each node (the formulas, and the rectangles that stand for ranges), as a CSR: the
        // walk goes from a stale node to what depends on it.
        let total = n + graph.rect_members.len();
        let mut start = vec![0u32; total + 2];
        for v in 0..total as u32 {
            for &w in graph.adj_known(v) {
                start[w as usize + 2] += 1;
            }
        }
        for i in 2..start.len() {
            start[i] += start[i - 1];
        }
        let mut readers = vec![0u32; start[total + 1] as usize];
        for v in 0..total as u32 {
            for &w in graph.adj_known(v) {
                let slot = &mut start[w as usize + 1];
                readers[*slot as usize] = v;
                *slot += 1;
            }
        }
        drop(graph);
        let mut hard_reached = vec![false; total];
        reach(&mut hard_reached, &readers, &start, &hard);
        let mut all_reached = hard_reached.clone();
        let both: Vec<bool> = hard.iter().zip(&soft).map(|(&h, &s)| h || s).collect();
        reach(&mut all_reached, &readers, &start, &both);
        stale_reached = hard_reached;
        reached = all_reached;
    }
    let (mut stale_reached, mut reached) = (stale_reached, reached);
    for (i, &load) in on_load.iter().enumerate() {
        stale_reached[i] |= load;
        reached[i] |= load;
    }
    let stale: FastSet<CellRef> = ats
        .iter()
        .enumerate()
        .filter(|(i, _)| stale_reached[*i])
        .map(|(_, at)| *at)
        .collect();
    if !cut {
        // Nothing is cut: the formulas go back as they came.
        formulas.extend(ats.into_iter().zip(kinds));
        return Kept {
            count: 0,
            stale: Some(stale),
        };
    }
    if keep_positions {
        positions.retain(|at| index.get(at).is_some_and(|id| reached[id as usize]));
    }
    let mut kept = 0;
    for (i, (at, kind)) in ats.iter().zip(kinds).enumerate() {
        if reached[i] {
            formulas.push((*at, kind));
        } else {
            let value = model.cells.get(at).map(|c| c.value.clone());
            cached.insert(*at, value.unwrap_or_default());
            kept += 1;
        }
    }
    Kept {
        count: kept,
        stale: Some(stale),
    }
}

/// Appends the references of `e` to `deps` and returns the flags of the functions it calls.
fn dependencies_into(e: &Expr, deps: &mut Vec<Rect>) -> Flags {
    let mut flags = Flags::default();
    e.walk(&mut |x| match x {
        Expr::Cell(c) => deps.push(Rect::single(*c)),
        Expr::Range(r) => deps.push(*r),
        Expr::Call(id, args) => {
            let d = id.def();
            flags.dynamic |= d.dynamic;
            flags.volatile |= d.volatile;
            // `CELL("filename")` says where the program is; the other types are the cell's.
            if d.name == "CELL" {
                flags.volatile |= !matches!(args.first(),
                    Some(Expr::Const(sheetrs_engine::value::Value::Text(t))) if !t.eq_ignore_ascii_case("FILENAME"));
            }
            // `LibreOffice` calculates these when it opens any file, and what reads them with them (OBSERVED, the default load
            // of both formats, `sheetrs-edit/tests/assets/write_to/probe_dirty.py classes`): `HYPERLINK`, every type of `CELL`
            // (`"filename"` too), `INFO` and `FORMULATEXT`, besides the volatile functions and `INDIRECT` and `OFFSET`.
            // The whole-column reads, `ROW`, `COLUMN`, `ISFORMULA` and `SHEET` keep what the file caches.
            flags.on_load |= matches!(d.name, "HYPERLINK" | "CELL" | "INFO" | "FORMULATEXT" | "FORMULA");
        }
        _ => {}
    });
    flags
}

/// Resolve `INDIRECT`/`OFFSET`/`_RANGE` whose arguments are known at build
/// time into plain references, so the graph (and the optimizer) sees them.
pub fn fold_refs(e: &mut Expr, at: CellRef, sheets: &[String], excel_refs: bool) {
    e.walk_mut(&mut |x| {
        let Expr::Call(id, args) = x else { return };
        let folded = match id.def().name {
            "_RANGE" => match (rect_of(&args[0]), rect_of(&args[1])) {
                (Some(p), Some(q)) if p.sheet == q.sheet => Some(Rect::new(
                    p.sheet,
                    p.r0.min(q.r0),
                    p.c0.min(q.c0),
                    p.r1.max(q.r1),
                    p.c1.max(q.c1),
                )),
                _ => None,
            },
            "OFFSET" => offset_const(args),
            "INDIRECT" => indirect_const(args, at, sheets, excel_refs),
            _ => None,
        };
        if let Some(r) = folded {
            *x = if r.is_single() {
                Expr::Cell(r.top_left())
            } else {
                Expr::Range(r)
            };
        }
    });
}

fn indirect_const(args: &[Expr], at: CellRef, sheets: &[String], excel_refs: bool) -> Option<Rect> {
    let Expr::Const(Value::Text(t)) = &args[0] else {
        return None;
    };
    if !matches!(args.get(1), None | Some(Expr::Const(Value::Bool(true)))) {
        return None;
    }
    let (sheet, tail) = addr::string_ref(t, excel_refs)?;
    let sheet = match sheet {
        Some(n) => sheets.iter().position(|s| addr::same_name(s, &n))? as u16,
        None => at.sheet,
    };
    addr::parse_area(&tail).map(|a| addr::bind(a, sheet))
}

fn rect_of(e: &Expr) -> Option<Rect> {
    match e {
        Expr::Cell(c) => Some(Rect::single(*c)),
        Expr::Range(r) => Some(*r),
        _ => None,
    }
}

fn const_num(e: &Expr) -> Option<f64> {
    match e {
        Expr::Const(Value::Number(n)) => Some(*n),
        _ => None,
    }
}

fn offset_const(args: &[Expr]) -> Option<Rect> {
    let base = rect_of(&args[0])?;
    let dr = const_num(&args[1])?.trunc() as i64;
    let dc = const_num(&args[2])?.trunc() as i64;
    let h = match args.get(3) {
        None | Some(Expr::Missing) => i64::from(base.rows()),
        Some(e) => const_num(e)?.trunc() as i64,
    };
    let w = match args.get(4) {
        None | Some(Expr::Missing) => i64::from(base.cols()),
        Some(e) => const_num(e)?.trunc() as i64,
    };
    let (r0, c0) = (i64::from(base.r0) + dr, i64::from(base.c0) + dc);
    let (r1, c1) = (r0 + h - 1, c0 + w - 1);
    if h < 1
        || w < 1
        || r0 < 0
        || c0 < 0
        || r1 >= i64::from(addr::MAX_ROWS)
        || c1 >= i64::from(addr::MAX_COLS)
    {
        return None;
    }
    Some(Rect::new(
        base.sheet, r0 as u32, c0 as u32, r1 as u32, c1 as u32,
    ))
}

#[cfg(test)]
mod build_time_tests {
    use super::{format_utc, formula_rects};
    use sheetrs_engine::addr::{self, CellRef};

    #[test]
    fn the_time_is_utc() {
        for (epoch, text) in [
            (0, "1970-01-01T00:00:00Z"),
            (951_782_400, "2000-02-29T00:00:00Z"),
            (1_793_361_600, "2026-10-30T12:00:00Z"),
            (1_709_251_199, "2024-02-29T23:59:59Z"),
            (-86_400, "1969-12-31T00:00:00Z"),
        ] {
            assert_eq!(format_utc(epoch), text, "{epoch}");
        }
    }

    #[test]
    fn formula_cells_become_the_rectangles_that_hold_exactly_them() {
        let at = |sheet, row, col| CellRef::new(sheet, row, col);
        // Sheet 0: a column (B1:B3), a cell (D1), a block (F1:G2) and a lone cell below the
        // column (B5, not joined to it); sheet 1: one cell.
        let cells = [
            at(0, 0, 1),
            at(0, 0, 3),
            at(0, 0, 5),
            at(0, 0, 6),
            at(0, 1, 1),
            at(0, 1, 5),
            at(0, 1, 6),
            at(0, 2, 1),
            at(0, 4, 1),
            at(1, 0, 0),
        ];
        let rects = formula_rects(&cells);
        let text: Vec<String> = rects.iter().map(addr::rect_text).collect();
        assert_eq!(text, ["B1:B3", "D1", "F1:G2", "B5", "A1"]);
        let mut held: Vec<(u16, u32, u32)> = rects
            .iter()
            .flat_map(|r| {
                (r.r0..=r.r1).flat_map(move |row| (r.c0..=r.c1).map(move |c| (r.sheet, row, c)))
            })
            .collect();
        held.sort_unstable();
        let mut wanted: Vec<(u16, u32, u32)> =
            cells.iter().map(|c| (c.sheet, c.row, c.col)).collect();
        wanted.sort_unstable();
        assert_eq!(held, wanted, "no constant, no formula left out");
    }
}

#[cfg(test)]
mod node_index_tests {
    use super::NodeIndex;
    use sheetrs_engine::addr::{CellRef, Rect};

    /// The nodes `ats` of `NodeIndex`, found as a plain scan of the positions would find them.
    fn check_index(ats: &[CellRef], sheets: usize) {
        let index = NodeIndex::new(ats, sheets);
        for (i, at) in ats.iter().enumerate() {
            assert_eq!(index.get(at), Some(i as u32), "{at:?}");
        }
        let last = ats
            .iter()
            .map(|a| (a.row, a.col))
            .fold((0, 0), |m, x| (m.0.max(x.0), m.1.max(x.1)));
        for sheet in 0..=(sheets as u16) {
            for row in 0..last.0 + 3 {
                for col in 0..last.1 + 3 {
                    let at = CellRef::new(sheet, row, col);
                    assert_eq!(
                        index.get(&at),
                        ats.iter().position(|a| *a == at).map(|i| i as u32)
                    );
                }
            }
        }
        for sheet in 0..=(sheets as u16) {
            for (r0, r1) in [(0, 0), (1, 1), (0, 3), (2, 9), (5, 5), (0, u32::MAX)] {
                for (c0, c1) in [(0, 0), (1, 1), (0, 2), (1, 7), (0, u32::MAX)] {
                    let r = Rect::new(sheet, r0, c0, r1, c1);
                    let mut got = Vec::new();
                    index.members(&r, &mut got);
                    let want: Vec<u32> = ats
                        .iter()
                        .enumerate()
                        .filter(|(_, a)| r.contains(**a))
                        .map(|(i, _)| i as u32)
                        .collect();
                    assert_eq!(got, want, "{r:?}");
                }
            }
        }
    }

    #[test]
    fn the_node_index_finds_cells_and_rectangles_as_a_scan_does() {
        let ats = [
            (0, 0, 0),
            (0, 0, 3),
            (0, 2, 1),
            (0, 2, 2),
            (0, 5, 0),
            (1, 1, 1),
            (1, 3, 0),
            (1, 3, 7),
            (3, 0, 0),
        ]
        .map(|(s, r, c)| CellRef::new(s, r, c));
        check_index(&ats, 4);
        check_index(&[], 2);
    }

    #[test]
    fn the_node_index_of_a_sparse_sheet_has_no_row_table_and_still_finds_cells() {
        // Two nodes a million rows apart: a table of rows would outweigh them.
        let ats = [
            CellRef::new(0, 1, 1),
            CellRef::new(0, 900_000, 2),
            CellRef::new(1, 4, 4),
        ];
        let index = NodeIndex::new(&ats, 2);
        assert!(index.sheets[0].rows.is_empty());
        assert!(!index.sheets[1].rows.is_empty());
        assert_eq!(index.get(&CellRef::new(0, 900_000, 2)), Some(1));
        assert_eq!(index.get(&CellRef::new(0, 900_000, 3)), None);
        let mut got = Vec::new();
        index.members(&Rect::new(0, 0, 0, 1_048_575, 16_383), &mut got);
        assert_eq!(got, vec![0, 1]);
        got.clear();
        index.members(&Rect::new(0, 2, 0, 899_999, 16_383), &mut got);
        assert!(got.is_empty());
    }
}
