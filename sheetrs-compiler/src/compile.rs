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
    /// Evaluate input-independent cells at build time.
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
    pub cells_kept: usize,
    pub cycles: usize,
    /// A reference computed at run time forced the whole workbook to be kept.
    pub dynamic_refs: bool,
    pub unsupported: BTreeSet<String>,
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

/// What lowering a share of the formulas found.
#[derive(Default)]
struct Lowered {
    formulas: Vec<(CellRef, CellKind)>,
    subtotal_cells: Vec<CellRef>,
    cached: FastMap<CellRef, Value>,
    diags: Vec<String>,
    unsupported: BTreeSet<String>,
    /// Functions of `LibreOffice` the engine lacks, with the first cell that calls each.
    missing: std::collections::BTreeMap<String, CellRef>,
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
) -> Lowered {
    let mut out = Lowered::default();
    for &(at, c) in part {
        let Some(f) = &c.formula else { continue };
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
                    e = Expr::Call(id, vec![e]);
                }
                let mut holds_subtotal = false;
                e.walk(&mut |x| match x {
                    Expr::Unknown(call) => {
                        let name = &call.0;
                        out.unsupported.insert(name.to_string());
                        if sheetrs_engine::functions::in_catalogue(name) {
                            out.missing.entry(name.to_string()).or_insert(at);
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
    let formula_positions: Vec<CellRef> = formula_cells.iter().map(|(at, _)| *at).collect();
    // Lowering a formula reads only the model: the cells go to as many threads as
    // the machine has, and what each one found is merged afterwards (the result is
    // sorted below, so the order they finish in does not matter).
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
        )]
    } else {
        std::thread::scope(|scope| {
            let handles: Vec<_> = formula_cells
                .chunks(chunk)
                .map(|part| {
                    scope.spawn(move || lower_chunk(model, part, calc_syntax, spec.features))
                })
                .collect();
            handles
                .into_iter()
                .map(|h| h.join().expect("a lowering thread panicked"))
                .collect()
        })
    };
    let mut missing: std::collections::BTreeMap<String, CellRef> =
        std::collections::BTreeMap::default();
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
        for (name, at) in part.missing {
            let first = missing.entry(name).or_insert(at);
            *first = (*first).min(at);
        }
        for (name, n) in part.function_uses {
            *report.function_uses.entry(name).or_default() += n;
        }
    }
    formulas.sort_unstable_by_key(|(at, _)| *at);
    subtotal_cells.sort_unstable();
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
    if !missing.is_empty() && !spec.inspect {
        let list: Vec<String> = missing
            .iter()
            .map(|(name, at)| {
                let sheet = model
                    .sheets
                    .get(at.sheet as usize)
                    .map_or("?", String::as_str);
                format!(
                    "{name} ({}x, first at {sheet}!{})",
                    report.function_uses.get(name).copied().unwrap_or(0),
                    sheetrs_engine::addr::a1(at.row, at.col)
                )
            })
            .collect();
        return Err(Error::Compile(format!(
            "the workbook calls {} function(s) of LibreOffice that this engine does not implement, so it cannot replace LibreOffice for it: {}",
            missing.len(),
            list.join(", ")
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
            .filter(|at| Some(at.sheet) != spec.skip_sheet)
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
    // What the program has to compute: the outputs.
    let wanted: Vec<Rect> = outputs.iter().map(|b| b.rect).collect();
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

    // Precedence: explicit option > what the document states > default.
    // The locale is the user's regional setting, which a file does not carry
    // (LibreOffice evaluates a document whose language says `es-ES` with the
    // system locale), so it comes from the caller and defaults to `en`. The
    // document's language is informational.
    let mut locale = spec.locale.clone().unwrap_or_default();
    if let Some(lang) = &model.settings.locale
        && !lang
            .to_lowercase()
            .starts_with(locale.tag.split('-').next().unwrap_or("en"))
    {
        diags.push(format!(
            "the document language is {lang}; text and number conversions follow the locale {} (build configuration)",
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
            inputs: Vec::new(),
            outputs: Vec::new(),
            queryable: Queryable::default(),
            mode: Mode::Values,
            formula_cells: Vec::new(),
            meta: Vec::new(),
            diagnostics: Vec::new(),
            build: BuildInfo::default(),
            layout: sheetrs_engine::engine::Layout::default(),
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

    let meta = program_meta(model, spec, &queryable, &inputs, &outputs);
    let formula_cells = if spec.mode.is_workbook() {
        // The formulas the workbook answer writes: the ones inside the outputs (all of
        // them when nothing is declared). They come in the order of the positions.
        let declared = !spec.outputs.is_empty();
        let written: Vec<CellRef> = formula_positions
            .into_iter()
            .filter(|at| {
                Some(at.sheet) != spec.skip_sheet
                    && (!declared
                        || outputs.iter().any(|b| {
                            let r = &b.rect;
                            r.sheet == at.sheet
                                && (r.r0..=r.r1).contains(&at.row)
                                && (r.c0..=r.c1).contains(&at.col)
                        }))
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
            inputs,
            outputs,
            mode: spec.mode,
            formula_cells,
            meta,
            queryable,
            diagnostics: diags,
            build: build_info(model, spec),
            layout: sheetrs_engine::engine::Layout::default(),
        },
        report,
    ))
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
